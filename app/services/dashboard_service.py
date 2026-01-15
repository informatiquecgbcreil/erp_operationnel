from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Tuple

from flask import url_for
from werkzeug.routing import BuildError
from sqlalchemy import func

from app.extensions import db
from app.models import (
    Subvention,
    Depense,
    LigneBudget,
    SessionActivite,
    PresenceActivite,
    Participant,
)


def _month_key(d: date) -> str:
    return f"{d.year:04d}-{d.month:02d}"


def _last_n_months(n: int, today: date | None = None) -> List[Tuple[int, int]]:
    today = today or date.today()
    y, m = today.year, today.month
    out: List[Tuple[int, int]] = []
    for _ in range(n):
        out.append((y, m))
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    out.reverse()
    return out


def build_dashboard_context(user, *, days: int = 90) -> Dict[str, Any]:
    """Construit un contexte riche pour le dashboard.

    - Ne modifie pas la DB.
    - Doit rester robuste : aucune url_for sur une route à paramètres obligatoires.
    """

    def _safe(endpoint: str, fallback: str = "#", **values) -> str:
        try:
            return url_for(endpoint, **values)
        except BuildError:
            return fallback

    has_perm = getattr(user, "has_perm", None)
    has_scope_all = callable(has_perm) and has_perm("scope:all_secteurs")
    has_business_access = callable(has_perm) and any(
        has_perm(p) for p in ("subventions:view", "projets:view", "stats:view", "statsimpact:view")
    )
    if callable(has_perm) and has_perm("admin:users") and not has_business_access:
        return {
            "mode": "admin_tech",
            "kpis": {},
            "alerts": [],
            "shortcuts": [
                {"label": "Gérer l’équipe", "url": _safe("admin.users"), "icon": "🛠️"},
            ],
            "recents": {"depenses": [], "sessions": [], "participants": []},
            "charts": {},
            "days": days,
        }

    # --- périmètre ---
    subs_q = Subvention.query.filter_by(est_archive=False)
    if not has_scope_all:
        subs_q = subs_q.filter(Subvention.secteur == user.secteur_assigne)
    subs = subs_q.all()

    # --- KPIs budget (déjà calculés via propriétés sur Subvention) ---
    total_attribue = sum(float(s.montant_attribue or 0) for s in subs)
    total_recu = sum(float(s.montant_recu or 0) for s in subs)
    total_engage = sum(float(s.total_engage or 0) for s in subs)
    total_reste = sum(float(s.total_reste or 0) for s in subs)
    taux = 0.0
    if total_attribue > 0:
        taux = round((total_engage / total_attribue) * 100, 1)

    # --- Alertes (pilotage) ---
    alerts: List[Dict[str, Any]] = []
    for s in subs:
        recu = float(s.montant_recu or 0)
        reel_lignes = float(s.total_reel_lignes or 0)
        engage = float(s.total_engage or 0)
        reste = float(s.total_reste or 0)

        # reçu mais pas ventilé
        if recu > 0 and reel_lignes == 0:
            alerts.append({
                "level": "danger",
                "text": f"{s.nom} : reçu {recu:.2f}€ mais lignes réel = 0€ (ventilation manquante).",
                "url": _safe("main.subvention_pilotage", subvention_id=s.id),
            })
        # engagé > réel lignes
        if reel_lignes > 0 and engage > reel_lignes:
            alerts.append({
                "level": "danger",
                "text": f"{s.nom} : engagé {engage:.2f}€ > lignes réel {reel_lignes:.2f}€ (dépassement).",
                "url": _safe("main.subvention_pilotage", subvention_id=s.id),
            })
        # proche du plafond
        if float(s.montant_attribue or 0) > 0:
            pct = (engage / float(s.montant_attribue or 0)) * 100
            if pct >= 80:
                alerts.append({
                    "level": "warning",
                    "text": f"{s.nom} : {pct:.0f}% consommé (reste {reste:.2f}€).",
                    "url": _safe("main.subvention_pilotage", subvention_id=s.id),
                })

    # --- Activité (fenêtre) ---
    since = datetime.utcnow() - timedelta(days=days)

    # Sessions / présences / uniques
    sessions_q = SessionActivite.query.filter_by(is_deleted=False)
    pres_q = PresenceActivite.query
    if not has_scope_all:
        sessions_q = sessions_q.filter(SessionActivite.secteur == user.secteur_assigne)
        pres_q = pres_q.join(SessionActivite).filter(SessionActivite.secteur == user.secteur_assigne)

    sessions_recent = sessions_q.filter(SessionActivite.created_at >= since).count()
    uniques_recent = pres_q.join(Participant).filter(PresenceActivite.created_at >= since).with_entities(Participant.id).distinct().count()

    # --- Graphiques ---
    months = _last_n_months(6)
    month_labels = [f"{y}-{m:02d}" for (y, m) in months]

    # Dépenses par mois (date_paiement sinon created_at)
    dep_q = Depense.query.filter_by(est_supprimee=False)
    if not has_scope_all:
        # LigneBudget n'a pas de colonne 'secteur' : le secteur est porté par la
        # Subvention (et/ou par les Projets). On filtre donc via Subvention.secteur.
        dep_q = (
            dep_q.join(LigneBudget)
            .join(Subvention, LigneBudget.subvention_id == Subvention.id)
            .filter(Subvention.secteur == user.secteur_assigne)
        )
    dep_rows = dep_q.with_entities(Depense.montant, Depense.date_paiement, Depense.created_at).all()

    dep_by_month = {k: 0.0 for k in month_labels}
    for montant, date_paiement, created_at in dep_rows:
        d = date_paiement or (created_at.date() if created_at else None)
        if not d:
            continue
        mk = _month_key(d)
        if mk in dep_by_month:
            dep_by_month[mk] += float(montant or 0)

    # Sessions par mois (réalisées / créées)
    sess_rows = sessions_q.with_entities(SessionActivite.created_at).all()
    sess_by_month = {k: 0 for k in month_labels}
    for (created_at,) in sess_rows:
        if not created_at:
            continue
        mk = _month_key(created_at.date())
        if mk in sess_by_month:
            sess_by_month[mk] += 1

    # Répartition des participants (uniques) par type_public sur la période
    pub_counts = {"H": 0, "S": 0, "B": 0, "A": 0, "P": 0, "?": 0}
    pub_rows = (
        pres_q.join(Participant)
        .filter(PresenceActivite.created_at >= since)
        .with_entities(Participant.id, Participant.type_public)
        .distinct()
        .all()
    )
    for _pid, tp in pub_rows:
        key = (tp or "?").strip().upper()
        if key not in pub_counts:
            key = "?"
        pub_counts[key] += 1

    charts = {
        "budget_donut": {
            "labels": ["Engagé", "Disponible"],
            "values": [round(total_engage, 2), round(max(total_attribue - total_engage, 0.0), 2)],
        },
        "depenses_bar": {
            "labels": month_labels,
            "values": [round(dep_by_month[k], 2) for k in month_labels],
        },
        "sessions_line": {
            "labels": month_labels,
            "values": [sess_by_month[k] for k in month_labels],
        },
        "public_pie": {
            "labels": ["Habitants", "Seniors", "Bénévoles", "Allophones", "Parents", "Autre"],
            "values": [pub_counts["H"], pub_counts["S"], pub_counts["B"], pub_counts["A"], pub_counts["P"], pub_counts["?"]],
        },
    }

    # --- récents ---
    recent_depenses = dep_q.order_by(Depense.created_at.desc()).limit(6).all()
    recent_sessions = sessions_q.order_by(SessionActivite.created_at.desc()).limit(6).all()
    recent_participants_q = Participant.query
    if not has_scope_all:
        recent_participants_q = recent_participants_q.filter(Participant.created_secteur == user.secteur_assigne)
    recent_participants = recent_participants_q.order_by(Participant.created_at.desc()).limit(6).all()

    shortcuts = [
        {"label": "Nouvelle dépense", "url": _safe("budget.depense_new"), "icon": "➕"},
        # route à paramètres -> on renvoie vers la liste des ateliers
        {"label": "Nouvelle session", "url": _safe("activite.index"), "icon": "📅"},
        {"label": "Participants", "url": _safe("activite.participants", fallback=_safe("activite.index")), "icon": "👥"},
        {"label": "Inventaire", "url": _safe("inventaire_materiel.list_items"), "icon": "📦"},
        {"label": "Données activités", "url": _safe("statsimpact.dashboard"), "icon": "📊"},
        {"label": "Stats & bilans", "url": _safe("main.stats_bilans", fallback=_safe("main.dashboard")), "icon": "🧾"},
    ]

    return {
        "mode": "global" if has_scope_all else "secteur",
        "days": days,
        "kpis": {
            "attribue": round(total_attribue, 2),
            "recu": round(total_recu, 2),
            "engage": round(total_engage, 2),
            "reste": round(total_reste, 2),
            "taux": taux,
            "sessions": sessions_recent,
            "uniques": uniques_recent,
        },
        "alerts": alerts[:12],
        "shortcuts": shortcuts,
        "recents": {
            "depenses": recent_depenses,
            "sessions": recent_sessions,
            "participants": recent_participants,
        },
        "charts": charts,
    }
