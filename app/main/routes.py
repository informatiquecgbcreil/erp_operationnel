import csv
from io import StringIO
from datetime import date


def _parse_iso_date(s: str):
    try:
        if not s:
            return None
        return date.fromisoformat(s)
    except Exception:
        return None

def _indicator_date_range(params: dict, selected_annee: int | None):
    period = (params.get("period") or "context").strip()
    if period == "custom":
        d1 = _parse_iso_date(params.get("start") or "")
        d2 = _parse_iso_date(params.get("end") or "")
        if d1 and d2 and d2 < d1:
            d1, d2 = d2, d1
        return d1, d2
    if period == "year" or period == "context":
        if selected_annee:
            return date(selected_annee, 1, 1), date(selected_annee, 12, 31)
    return None, None

def _indicator_target_status(value, target, op: str):
    if target is None or value is None:
        return None
    try:
        v = float(value)
        t = float(target)
    except Exception:
        return None
    if t == 0:
        return None
    op = (op or "ge").strip()

    # ge : on veut v >= t ; le : on veut v <= t
    if op == "le":
        ratio = t / v if v != 0 else float("inf")
        ok = v <= t
    else:
        ratio = v / t
        ok = v >= t

    if ok:
        return "ok"
    if ratio >= 0.75:
        return "warn"
    return "bad"
from flask import (
    Blueprint, render_template, request, redirect, url_for, flash,
    abort, current_app, Response, jsonify
)
from flask_login import login_required, current_user
from app.rbac import require_perm, can, can_access_secteur
from app.secteurs import get_secteur_labels

from app.extensions import db
from app.models import Subvention, LigneBudget, Depense, Projet, SubventionProjet, AtelierActivite, SessionActivite, PresenceActivite, ProjetAtelier, ProjetIndicateur
from app.services.dashboard_service import build_dashboard_context

bp = Blueprint("main", __name__)

# --------- Permissions ---------
def can_see_secteur(secteur: str) -> bool:
    return can_access_secteur(secteur)


def _compute_prorata(lignes, montant_cible: float):
    """
    Calcule une répartition pro-rata sur montant_base.
    Ne modifie pas la DB : retourne un dict {ligne_id: montant_theorique}
    Ajuste la dernière ligne pour tomber pile au centime.
    """
    lignes = list(lignes)
    out = {}
    if not lignes:
        return out

    total_base = sum(float(l.montant_base or 0) for l in lignes)
    if total_base <= 0:
        for l in lignes:
            out[l.id] = 0.0
        return out

    ratio = float(montant_cible or 0) / total_base

    cumul = 0.0
    for i, l in enumerate(lignes):
        base = float(l.montant_base or 0)
        part = round(base * ratio, 2)
        if i == len(lignes) - 1:
            part = round(float(montant_cible or 0) - cumul, 2)
        out[l.id] = float(part)
        cumul += float(part)

    return out


# --------- Setup start ---------
@bp.route("/setup-start")
def setup_start():
    # simple page de diagnostic / aide
    return render_template("controle.html")


# --------- Dashboard ---------
@bp.route("/dashboard")
@login_required
@require_perm("dashboard:view")
def dashboard():
    # Période "activité" (utilisée pour les KPIs atelier/participants)
    try:
        days = int(request.args.get("days") or 90)
    except Exception:
        days = 90

    ctx = build_dashboard_context(current_user, days=days)
    return render_template("dashboard.html", **ctx)


# --------- List subventions ---------
@bp.route("/subventions")
@login_required
@require_perm("subventions:view")
def subventions_list():
    secteurs = get_secteur_labels(active_only=False)

    subs_q = Subvention.query.filter_by(est_archive=False)
    if not can("scope:all_secteurs"):
        subs_q = subs_q.filter(Subvention.secteur == current_user.secteur_assigne)

    subs = subs_q.order_by(Subvention.annee_exercice.desc(), Subvention.nom.asc()).all()
    return render_template("subventions_list.html", subs=subs, secteurs=secteurs)


@bp.route("/subvention/nouvelle", methods=["POST"])
@login_required
@require_perm('subventions:edit')
def subvention_create():
    nom = (request.form.get("nom") or "").strip()
    secteur = (request.form.get("secteur") or "").strip()
    annee = int(request.form.get("annee_exercice") or 2025)

    montant_demande = float(request.form.get("montant_demande") or 0)
    montant_attribue = float(request.form.get("montant_attribue") or 0)
    montant_recu = float(request.form.get("montant_recu") or 0)

    if not can("scope:all_secteurs"):
        secteur = current_user.secteur_assigne

    if not nom or not secteur:
        flash("Nom + secteur obligatoires.", "danger")
        return redirect(url_for("main.subventions_list"))

    if not can_see_secteur(secteur):
        abort(403)

    s = Subvention(
        nom=nom,
        secteur=secteur,
        annee_exercice=annee,
        montant_demande=montant_demande,
        montant_attribue=montant_attribue,
        montant_recu=montant_recu,
    )
    db.session.add(s)
    db.session.commit()

    flash("Subvention créée.", "success")
    return redirect(url_for("main.subvention_pilotage", subvention_id=s.id))


# --------- Pilotage subvention ---------
@bp.route("/subvention/<int:subvention_id>/pilotage", methods=["GET", "POST"])
@login_required
@require_perm("subventions:view")
def subvention_pilotage(subvention_id):
    sub = Subvention.query.get_or_404(subvention_id)
    if not can_see_secteur(sub.secteur):
        abort(403)

    if request.method == "POST":
        if not can("subventions:edit"):
            abort(403)
        action = request.form.get("action") or ""

        # --- Montants globaux ---
        if action == "update_montants":
            sub.montant_demande = float(request.form.get("montant_demande") or 0)
            sub.montant_attribue = float(request.form.get("montant_attribue") or 0)
            sub.montant_recu = float(request.form.get("montant_recu") or 0)
            db.session.commit()
            flash("Montants mis à jour.", "success")
            return redirect(url_for("main.subvention_pilotage", subvention_id=sub.id))

        # --- Ajouter une ligne ---
        if action == "add_ligne":
            compte = (request.form.get("compte") or "60").strip()
            libelle = (request.form.get("libelle") or "").strip()
            montant_base = float(request.form.get("montant_base") or 0)
            montant_reel = float(request.form.get("montant_reel") or 0)
            nature = (request.form.get("nature") or "charge").strip()


            if not libelle:
                flash("Libellé obligatoire.", "danger")
                return redirect(url_for("main.subvention_pilotage", subvention_id=sub.id))

            l = LigneBudget(
                subvention_id=sub.id,
                nature=nature,
                compte=compte,
                libelle=libelle,
                montant_base=montant_base,
                montant_reel=montant_reel,
            )
            db.session.add(l)
            db.session.commit()
            flash("Ligne ajoutée.", "success")
            return redirect(url_for("main.subvention_pilotage", subvention_id=sub.id))

        # --- Ventilation automatique (écrit dans montant_reel) ---
        if action == "auto_ventilation":
            mode = (request.form.get("mode") or "copy_base").strip()
            target = (request.form.get("target") or "recu").strip()

            if target == "attribue":
                montant_cible = float(sub.montant_attribue or 0)
            else:
                montant_cible = float(sub.montant_recu or 0)

            lignes = list(sub.lignes)
            if not lignes:
                flash("Aucune ligne à ventiler.", "warning")
                return redirect(url_for("main.subvention_pilotage", subvention_id=sub.id))

            if mode == "reset":
                for l in lignes:
                    l.montant_reel = 0.0
                db.session.commit()
                flash("Ventilation réinitialisée (réel = 0).", "warning")
                return redirect(url_for("main.subvention_pilotage", subvention_id=sub.id))

            if mode == "copy_base":
                for l in lignes:
                    l.montant_reel = float(l.montant_base or 0)
                db.session.commit()
                flash("Ventilation : base copiée vers réel.", "success")
                return redirect(url_for("main.subvention_pilotage", subvention_id=sub.id))

            if mode == "prorata_base":
                total_base = sum(float(l.montant_base or 0) for l in lignes)
                if total_base <= 0:
                    flash("Impossible : total des bases = 0.", "danger")
                    return redirect(url_for("main.subvention_pilotage", subvention_id=sub.id))

                theor = _compute_prorata(lignes, montant_cible)
                for l in lignes:
                    l.montant_reel = float(theor.get(l.id, 0.0))
                db.session.commit()
                flash(f"Ventilation pro-rata effectuée sur {montant_cible:.2f}€.", "success")
                return redirect(url_for("main.subvention_pilotage", subvention_id=sub.id))

            abort(400)

        abort(400)

    # --- GET : données pour page ---
    projets_q = Projet.query.filter(Projet.secteur == sub.secteur)
    if not can("scope:all_secteurs"):
        projets_q = projets_q.filter(Projet.secteur == current_user.secteur_assigne)

    projets = projets_q.order_by(Projet.nom.asc()).all()
    linked_ids = set(sp.projet_id for sp in sub.projets)

    lignes = list(sub.lignes)
    total_base = round(sum(float(l.montant_base or 0) for l in lignes), 2)

    theor_recu = _compute_prorata(lignes, float(sub.montant_recu or 0))
    theor_attribue = _compute_prorata(lignes, float(sub.montant_attribue or 0))

    recu = float(sub.montant_recu or 0)
    reel_lignes = float(sub.total_reel_lignes or 0)
    engage = float(sub.total_engage or 0)

    warnings = []
    if recu > 0 and reel_lignes == 0:
        warnings.append("Tu as un montant reçu, mais aucune ventilation en lignes réel : utilise la ventilation auto ou renseigne le réel par ligne.")
    if recu > 0 and reel_lignes > 0 and reel_lignes < recu:
        warnings.append("Ventilation partielle : total lignes réel < montant reçu. Il manque une répartition.")
    if reel_lignes > 0 and engage > reel_lignes:
        warnings.append("Attention : engagé > total lignes réel (dépenses au-dessus de l'enveloppe ventilée).")

    return render_template(
        "budget_pilotage.html",
        sub=sub,
        projets=projets,
        linked_ids=linked_ids,
        total_base=total_base,
        theor_recu=theor_recu,
        theor_attribue=theor_attribue,
        warnings=warnings
    )


@bp.route("/subvention/<int:subvention_id>/delete", methods=["POST"])
@login_required
@require_perm('subventions:delete')
def subvention_delete(subvention_id):
    sub = Subvention.query.get_or_404(subvention_id)
    if not can_see_secteur(sub.secteur):
        abort(403)

    db.session.delete(sub)
    db.session.commit()
    flash("Subvention supprimée.", "warning")
    return redirect(url_for("main.subventions_list"))


# --------- Edit / Delete lignes ---------
@bp.route("/ligne/<int:ligne_id>/edit", methods=["POST"])
@login_required
@require_perm("subventions:edit")
def ligne_edit(ligne_id):
    l = LigneBudget.query.get_or_404(ligne_id)
    sub = l.source_sub
    if not can_see_secteur(sub.secteur):
        abort(403)

    l.nature = (request.form.get("nature") or getattr(l, "nature", "charge")).strip()
    l.compte = (request.form.get("compte") or l.compte).strip()
    l.libelle = (request.form.get("libelle") or l.libelle).strip()
    l.montant_base = float(request.form.get("montant_base") or l.montant_base or 0)
    l.montant_reel = float(request.form.get("montant_reel") or l.montant_reel or 0)
    db.session.commit()

    flash("Ligne modifiée.", "success")
    return redirect(url_for("main.subvention_pilotage", subvention_id=sub.id))


@bp.route("/ligne/<int:ligne_id>/delete", methods=["POST"])
@login_required
@require_perm('budget:delete')
def ligne_delete(ligne_id):
    l = LigneBudget.query.get_or_404(ligne_id)
    sub = l.source_sub
    if not can_see_secteur(sub.secteur):
        abort(403)

    db.session.delete(l)
    db.session.commit()

    flash("Ligne supprimée.", "warning")
    return redirect(url_for("main.subvention_pilotage", subvention_id=sub.id))


# --------- Lier / délier subvention à projet ---------
@bp.route("/subvention/<int:subvention_id>/toggle_projet", methods=["POST"])
@login_required
@require_perm("subventions:link")
def subvention_toggle_projet(subvention_id):
    sub = Subvention.query.get_or_404(subvention_id)
    if not can_see_secteur(sub.secteur):
        abort(403)

    projet_id = int(request.form.get("projet_id") or 0)
    projet = Projet.query.get_or_404(projet_id)

    if projet.secteur != sub.secteur:
        abort(400)

    link = SubventionProjet.query.filter_by(projet_id=projet.id, subvention_id=sub.id).first()
    if link:
        db.session.delete(link)
        db.session.commit()
        flash("Subvention retirée du projet.", "warning")
    else:
        link = SubventionProjet(projet_id=projet.id, subvention_id=sub.id)
        db.session.add(link)
        db.session.commit()
        flash("Subvention ajoutée au projet.", "success")

    return redirect(url_for("main.subvention_pilotage", subvention_id=sub.id))


# --------- APIs pour dropdowns dépenses ---------
@bp.route("/api/subvention/<int:subvention_id>/comptes")
@login_required
@require_perm("subventions:view")
def api_comptes(subvention_id):
    sub = Subvention.query.get_or_404(subvention_id)
    if not can_see_secteur(sub.secteur):
        abort(403)

    nature = request.args.get("nature")

    q = LigneBudget.query.filter_by(subvention_id=sub.id)
    if nature:
        q = q.filter(LigneBudget.nature == nature)

    comptes = sorted({l.compte for l in q if l.compte})
    return jsonify({"comptes": comptes})


@bp.route("/api/subvention/<int:subvention_id>/lignes")
@login_required
@require_perm("subventions:view")
def api_lignes(subvention_id):
    sub = Subvention.query.get_or_404(subvention_id)
    if not can_see_secteur(sub.secteur):
        abort(403)

    compte = (request.args.get("compte") or "").strip()
    nature = request.args.get("nature")

    q = LigneBudget.query.filter_by(subvention_id=sub.id)
    if nature:
        q = q.filter(LigneBudget.nature == nature)
    if compte:
        q = q.filter(LigneBudget.compte == compte)

    lignes = q.order_by(LigneBudget.compte.asc(), LigneBudget.libelle.asc()).all()

    out = []
    for l in lignes:
        out.append({
            "id": l.id,
            "compte": l.compte,
            "libelle": l.libelle,
            "montant_reel": float(l.montant_reel or 0),
            "engage": float(l.engage or 0),
            "reste": float(l.reste or 0),
        })

    return jsonify({"lignes": out})


# --------- Stats ---------
@bp.route("/stats")
@login_required
@require_perm("stats:view")
def stats():
    """
    Vue synthèse des budgets avec représentation graphique.

    On peut filtrer par année et/ou secteur via des paramètres GET.
    Responsable de secteur : le filtre secteur est forcé sur son secteur.
    Option : filtre projet (projet_id) pour croiser finance + indicateurs participants.
    """
    has_global_scope = can("stats:view_all") or can("scope:all_secteurs")

    # --- Lecture filtres (année, secteur, projet) ---
    annee_raw = (request.args.get("annee") or "").strip()
    secteur_raw = (request.args.get("secteur") or "").strip()
    projet_id_raw = (request.args.get("projet_id") or "").strip()

    selected_annee: int | None = None
    if annee_raw:
        try:
            selected_annee = int(annee_raw)
        except ValueError:
            selected_annee = None

    selected_secteur: str | None = secteur_raw or None

    selected_projet_id: int | None = None
    if projet_id_raw:
        try:
            selected_projet_id = int(projet_id_raw)
        except ValueError:
            selected_projet_id = None

    # --- Base query ---
    sub_q = Subvention.query.filter_by(est_archive=False)
    proj_q = Projet.query

    # Filtre année
    if selected_annee:
        sub_q = sub_q.filter(Subvention.annee_exercice == selected_annee)

    # Filtre secteur
    if selected_secteur:
        sub_q = sub_q.filter(Subvention.secteur == selected_secteur)
        proj_q = proj_q.filter(Projet.secteur == selected_secteur)

    # Filtre projet (finance)
    if selected_projet_id:
        proj_q = proj_q.filter(Projet.id == selected_projet_id)
        sub_q = sub_q.join(SubventionProjet, SubventionProjet.subvention_id == Subvention.id)                   .filter(SubventionProjet.projet_id == selected_projet_id)

    # Restriction responsable secteur
    if not has_global_scope:
        sub_q = sub_q.filter(Subvention.secteur == current_user.secteur_assigne)
        proj_q = proj_q.filter(Projet.secteur == current_user.secteur_assigne)
        selected_secteur = current_user.secteur_assigne

        # On blind le projet : seulement ceux de son secteur
        if selected_projet_id:
            p_tmp = Projet.query.get(selected_projet_id)
            if not p_tmp or p_tmp.secteur != current_user.secteur_assigne:
                selected_projet_id = None

    subs = sub_q.order_by(Subvention.annee_exercice.desc(), Subvention.nom.asc()).all()
    projets = proj_q.order_by(Projet.nom.asc()).all()

    # Pré-calcul des années disponibles (pour sélecteur)
    all_annees = sorted({s.annee_exercice for s in Subvention.query.filter_by(est_archive=False).all()}, reverse=True)
    all_secteurs = get_secteur_labels(active_only=False)

    # --- Totaux globaux ---
    total_recu = round(sum(float(s.montant_recu or 0) for s in subs), 2)
    total_engage = round(sum(float(s.total_engage or 0) for s in subs), 2)
    total_reste = round(sum(float(s.total_reste or 0) for s in subs), 2)

    # --- Agrégation par secteur ---
    by_secteur: dict[str, dict[str, float]] = {}
    for s in subs:
        d = by_secteur.setdefault(s.secteur, {"recu": 0.0, "engage": 0.0, "reste": 0.0})
        d["recu"] += float(s.montant_recu or 0)
        d["engage"] += float(s.total_engage or 0)
        d["reste"] += float(s.total_reste or 0)
    for sec, vals in by_secteur.items():
        vals["recu"] = round(vals.get("recu", 0.0), 2)
        vals["engage"] = round(vals.get("engage", 0.0), 2)
        vals["reste"] = round(vals.get("reste", 0.0), 2)

    # --- Agrégation par compte ---
    by_compte: dict[str, dict[str, float]] = {}
    for s in subs:
        for l in s.lignes:
            d = by_compte.setdefault(l.compte, {"reel": 0.0, "engage": 0.0, "reste": 0.0})
            d["reel"] += float(l.montant_reel or 0)
            d["engage"] += float(l.engage or 0)
            d["reste"] += float(l.reste or 0)
    for comp, vals in by_compte.items():
        vals["reel"] = round(vals.get("reel", 0.0), 2)
        vals["engage"] = round(vals.get("engage", 0.0), 2)
        vals["reste"] = round(vals.get("reste", 0.0), 2)

    # --- Détails par projet ---
    by_projet: list[dict[str, float | str]] = []
    for p in projets:
        by_projet.append({
            "id": p.id,
            "nom": p.nom,
            "secteur": p.secteur,
            "demande": p.total_demande,
            "attribue": p.total_attribue,
            "recu": p.total_recu,
            "reel_lignes": p.total_reel_lignes,
            "engage": p.total_engage,
            "reste": p.total_reste,
        })

    # Valeurs max pour barres proportionnelles
    max_secteur_total = max([v["recu"] + v["engage"] + v["reste"] for v in by_secteur.values()] + [0.0])
    max_compte_total = max([v["reel"] + v["engage"] + v["reste"] for v in by_compte.values()] + [0.0])
    max_projet_total = max([p["recu"] + p["engage"] + p["reste"] for p in by_projet] + [0.0])

    # --- Indicateurs projet (si projet sélectionné) ---
    project_indicators = []
    selected_projet = None

    if selected_projet_id:
        selected_projet = Projet.query.get(selected_projet_id)

    if selected_projet and can_see_secteur(selected_projet.secteur):
        atelier_ids = [lnk.atelier_id for lnk in ProjetAtelier.query.filter_by(projet_id=selected_projet.id).all()]

        # borne temporelle : année si fournie, sinon pas de filtre
        date_min = date_max = None
        if selected_annee:
            date_min = date(selected_annee, 1, 1)
            date_max = date(selected_annee, 12, 31)


        # Valeurs calculées "par défaut" sur l'année sélectionnée (context)
        base_date_min = base_date_max = None
        if selected_annee:
            base_date_min = date(selected_annee, 1, 1)
            base_date_max = date(selected_annee, 12, 31)

        def _compute_participants_metrics(atelier_ids_scope, dmin, dmax):
            out = {
                "participants_uniques": 0,
                "presences_totales": 0,
                "sessions_totales": 0,
                "recurrence_2plus": 0,
            }
            if not atelier_ids_scope:
                return out

            sess_q = SessionActivite.query.filter(SessionActivite.atelier_id.in_(atelier_ids_scope))                 .filter(SessionActivite.is_deleted == False)                 .filter(SessionActivite.statut != "annulee")

            session_date = db.func.coalesce(SessionActivite.date_session, SessionActivite.rdv_date)
            if dmin and dmax:
                sess_q = sess_q.filter(session_date >= dmin).filter(session_date <= dmax)

            out["sessions_totales"] = int(sess_q.count())

            sess_ids = [r[0] for r in sess_q.with_entities(SessionActivite.id).all()]
            if not sess_ids:
                return out

            pres_q = PresenceActivite.query.filter(PresenceActivite.session_id.in_(sess_ids))
            out["presences_totales"] = int(pres_q.count())
            out["participants_uniques"] = int(
                pres_q.with_entities(db.func.count(db.distinct(PresenceActivite.participant_id))).scalar() or 0
            )

            sub = pres_q.with_entities(
                PresenceActivite.participant_id,
                db.func.count(PresenceActivite.id).label("c"),
            ).group_by(PresenceActivite.participant_id).having(db.func.count(PresenceActivite.id) >= 2)
            out["recurrence_2plus"] = int(sub.count())
            return out

        # Finances : charges / produits sur les subventions déjà filtrées (année/secteur/projet)
        dep = 0.0
        rec = 0.0
        for s in subs:
            for l in s.lignes:
                mt = float(l.montant_reel or 0)
                if (l.nature or "").lower() == "charge":
                    dep += mt
                elif (l.nature or "").lower() == "produit":
                    rec += mt
        dep = round(dep, 2)
        rec = round(rec, 2)

        inds = ProjetIndicateur.query.filter_by(projet_id=selected_projet.id, is_active=True)             .order_by(ProjetIndicateur.created_at.asc()).all()

        unit_map = {
            "depenses_totales": "€",
            "recettes_totales": "€",
            "cout_par_participant": "€",
            "cout_par_presence": "€",
        }

        for ind in inds:
            params = ind.params() or {}
            dmin, dmax = _indicator_date_range(params, selected_annee)

            # scope atelier (si défini)
            scope_atelier_ids = list(atelier_ids)
            atelier_id = params.get("atelier_id")
            try:
                if atelier_id:
                    atelier_id_int = int(atelier_id)
                    if atelier_id_int in atelier_ids:
                        scope_atelier_ids = [atelier_id_int]
            except Exception:
                pass

            metrics = _compute_participants_metrics(scope_atelier_ids, dmin, dmax)

            # Valeur selon le code
            val = None
            code = ind.code

            if code in ("participants_uniques", "presences_totales", "sessions_totales", "recurrence_2plus"):
                val = metrics.get(code, 0)

            elif code == "depenses_totales":
                val = dep

            elif code == "recettes_totales":
                val = rec

            elif code == "cout_par_participant":
                u = metrics.get("participants_uniques", 0) or 0
                val = round(dep / u, 2) if u else None

            elif code == "cout_par_presence":
                u = metrics.get("presences_totales", 0) or 0
                val = round(dep / u, 2) if u else None

            # objectifs (optionnels)
            target = params.get("target", None)
            op = params.get("target_op", "ge")
            status = _indicator_target_status(val, target, op)

            project_indicators.append({
                "label": ind.label,
                "code": ind.code,
                "value": val,
                "unit": unit_map.get(ind.code, ""),
                "target": target,
                "target_op": op,
                "status": status,
                "period": (params.get("period") or "context"),
                "start": params.get("start"),
                "end": params.get("end"),
                "atelier_id": params.get("atelier_id"),
            })
        inds = ProjetIndicateur.query.filter_by(projet_id=selected_projet.id, is_active=True).order_by(ProjetIndicateur.created_at.asc()).all()

    return render_template(
        "stats.html",
        total_recu=total_recu,
        total_engage=total_engage,
        total_reste=total_reste,
        by_secteur=by_secteur,
        by_compte=by_compte,
        by_projet=by_projet,
        max_secteur_total=max_secteur_total,
        max_compte_total=max_compte_total,
        max_projet_total=max_projet_total,
        all_annees=all_annees,
        all_secteurs=all_secteurs,
        selected_annee=selected_annee,
        selected_secteur=selected_secteur,
        selected_projet_id=selected_projet_id,
        projets_for_filter=projets,
        selected_projet=selected_projet,
        project_indicators=project_indicators,
    )


# --- Hub ergonomique : 1 menu "Stats & bilans" ---
@bp.route("/stats-bilans")
@login_required
@require_perm("stats:view")
def stats_bilans():
    # Objectif : éviter la confusion entre 2 écrans. Ici on explique la différence
    # et on redirige vers l'écran choisi (sans casser l'existant).
    return render_template("stats_bilans.html")



# --------- Contrôle ---------
@bp.route("/controle")
@login_required
@require_perm("controle:view")
def controle():
    return render_template("controle.html")


# --------- Bilan financeurs (global) ---------
from sqlalchemy import distinct

@bp.route("/bilan")
@login_required
@require_perm("bilans:view")
def bilan_global():
    has_global_scope = can("scope:all_secteurs")

    # --- Lecture filtres ---
    annee_raw = (request.args.get("annee") or "").strip()
    secteur_raw = (request.args.get("secteur") or "").strip()
    projet_id_raw = (request.args.get("projet_id") or "").strip()

    selected_annee = None
    if annee_raw:
        try:
            selected_annee = int(annee_raw)
        except ValueError:
            selected_annee = None

    selected_secteur = secteur_raw or None
    selected_projet_id = None
    if projet_id_raw:
        try:
            selected_projet_id = int(projet_id_raw)
        except ValueError:
            selected_projet_id = None

    # --- RESPONSABLE SECTEUR : on force le secteur, mais on autorise le filtre projet
    # uniquement si le projet appartient au même secteur.
    if not has_global_scope:
        selected_secteur = current_user.secteur_assigne
        if selected_projet_id:
            pjt = Projet.query.get(selected_projet_id)
            sec_pjt = (pjt.secteur or "").strip().lower() if pjt else ""
            sec_user = (selected_secteur or "").strip().lower()
            if (not pjt) or (sec_pjt != sec_user):
                selected_projet_id = None

    # --- Base query ---
    q = Subvention.query.filter_by(est_archive=False)

    if selected_annee is not None:
        q = q.filter(Subvention.annee_exercice == selected_annee)

    if selected_secteur:
        q = q.filter(Subvention.secteur == selected_secteur)

    # --- Filtre projet (si demandé) ---
    # Important : on ne doit pas créer de cartesian product -> join propre
    if selected_projet_id:
        # join via table d'association
        q = q.join(SubventionProjet, SubventionProjet.subvention_id == Subvention.id)\
             .filter(SubventionProjet.projet_id == selected_projet_id)

    subs = q.order_by(
        Subvention.annee_exercice.desc(),
        Subvention.secteur.asc(),
        Subvention.nom.asc()
    ).all()

    # --- Totaux ---
    totals = {
        "demande": round(sum(float(s.montant_demande or 0) for s in subs), 2),
        "attribue": round(sum(float(s.montant_attribue or 0) for s in subs), 2),
        "recu": round(sum(float(s.montant_recu or 0) for s in subs), 2),
        "reel_lignes": round(sum(float(s.total_reel_lignes or 0) for s in subs), 2),
        "engage": round(sum(float(s.total_engage or 0) for s in subs), 2),
        "reste": round(sum(float(s.total_reste or 0) for s in subs), 2),
    }

    # --- Alertes simples (optionnel mais utile) ---
    alertes = []
    for s in subs:
        recu = float(s.montant_recu or 0)
        reel_lignes = float(s.total_reel_lignes or 0)
        engage = float(s.total_engage or 0)

        if recu > 0 and reel_lignes == 0:
            alertes.append(f"{s.nom} : reçu {recu:.2f}€ mais lignes réel = 0€ (ventilation manquante).")
        if recu > 0 and reel_lignes > 0 and reel_lignes < recu:
            alertes.append(f"{s.nom} : reçu {recu:.2f}€ mais lignes réel = {reel_lignes:.2f}€ (ventilation incomplète).")
        if reel_lignes > 0 and engage > reel_lignes:
            alertes.append(f"{s.nom} : engagé {engage:.2f}€ > lignes réel {reel_lignes:.2f}€ (dépassement).")

    # --- Listes de filtres affichées (secteurs / projets) ---
    # secteurs : soit config, soit distinct en base, MAIS filtré par rôle
    secteurs = get_secteur_labels(active_only=False)
    if not secteurs:
        secteurs = [r[0] for r in db.session.query(distinct(Subvention.secteur)).all() if r[0]]

    if not has_global_scope:
        secteurs = [current_user.secteur_assigne]

    # projets : uniquement ceux visibles
    projets_q = Projet.query
    if not has_global_scope:
        projets_q = projets_q.filter(Projet.secteur == current_user.secteur_assigne)
    projets = projets_q.order_by(Projet.secteur.asc(), Projet.nom.asc()).all()

    return render_template(
        "bilan.html",
        subs=subs,
        totals=type("Obj", (), totals),  # petit hack pour totals.demande etc si ton template utilise des attributs
        alertes=alertes,
        secteurs=secteurs,
        projets=projets,
        selected_annee=selected_annee,
        selected_secteur=selected_secteur,
        selected_projet_id=selected_projet_id
    )


# Alias de compat (si ton layout appelle encore main.bilan)

@bp.route("/bilan-global")
@login_required
def bilan():
    return redirect(url_for("main.bilan_global"))


# --------- Exports simples ---------
@bp.route("/export/depenses.csv")
@login_required
@require_perm("depenses:view")
def export_depenses_csv():
    dep_q = Depense.query.join(LigneBudget).join(Subvention)
    if not can("scope:all_secteurs"):
        dep_q = dep_q.filter(Subvention.secteur == current_user.secteur_assigne)

    deps = dep_q.all()

    out = StringIO()
    writer = csv.writer(out, delimiter=";")
    writer.writerow(["secteur", "subvention", "annee", "compte", "ligne", "depense", "montant", "date_paiement", "type"])

    for d in deps:
        l = d.budget_source
        s = l.source_sub
        writer.writerow([
            s.secteur,
            s.nom,
            s.annee_exercice,
            l.compte,
            l.libelle,
            d.libelle,
            f"{float(d.montant or 0):.2f}".replace(".", ","),
            d.date_paiement.isoformat() if d.date_paiement else "",
            d.type_depense or ""
        ])

    content = out.getvalue().encode("utf-8-sig")  # Excel friendly
    filename = f"depenses_{date.today().isoformat()}.csv"
    return Response(content, mimetype="text/csv", headers={
        "Content-Disposition": f"attachment; filename={filename}"
    })


@bp.route("/export/subvention/<int:subvention_id>.csv")
@login_required
@require_perm("subventions:view")
def export_subvention_csv(subvention_id):
    s = Subvention.query.get_or_404(subvention_id)
    if not can_see_secteur(s.secteur):
        abort(403)

    out = StringIO()
    writer = csv.writer(out, delimiter=";")
    writer.writerow(["subvention", "secteur", "annee", "compte", "ligne", "base", "reel", "engage", "reste"])

    for l in s.lignes:
        writer.writerow([
            s.nom,
            s.secteur,
            s.annee_exercice,
            l.compte,
            l.libelle,
            f"{float(l.montant_base or 0):.2f}".replace(".", ","),
            f"{float(l.montant_reel or 0):.2f}".replace(".", ","),
            f"{float(l.engage or 0):.2f}".replace(".", ","),
            f"{float(l.reste or 0):.2f}".replace(".", ","),
        ])

    content = out.getvalue().encode("utf-8-sig")
    return Response(content, mimetype="text/csv", headers={
        "Content-Disposition": f"attachment; filename=subvention_{s.id}.csv"
    })


# --------- Bilan par subvention ---------
@bp.route("/subvention/<int:subvention_id>/bilan")
@login_required
@require_perm("subventions:view")
def subvention_bilan(subvention_id: int):
    """
    Vue détaillée pour un financeur / subvention.

    Affiche un récapitulatif des montants demandés/attribués/reçus et des lignes de budget,
    avec une représentation graphique proportionnelle par ligne (charges et produits).
    """
    sub = Subvention.query.get_or_404(subvention_id)
    # Vérification des droits : un responsable ne peut consulter que son secteur
    if not can_see_secteur(sub.secteur):
        abort(403)

    # Collecte des lignes de budget
    lignes: list[dict[str, float | str]] = []
    # Calcul du montant maximum utilisé pour la largeur des barres
    max_total = 0.0
    for l in sub.lignes:
        base = float(l.montant_base or 0)
        reel = float(l.montant_reel or 0)
        engage = float(l.engage or 0)
        reste = float(l.reste or 0)
        nature = getattr(l, "nature", "charge")
        total_for_max = reel + engage + reste
        if total_for_max > max_total:
            max_total = total_for_max
        lignes.append({
            "id": l.id,
            "compte": l.compte,
            "libelle": l.libelle,
            "base": base,
            "reel": reel,
            "engage": engage,
            "reste": reste,
            "nature": nature,
        })

    # Calcul des pourcentages pour chaque ligne (évite de surcharger le template)
    if max_total > 0:
        for d in lignes:
            d["p_reel"] = (d["reel"] / max_total) * 100.0 if d["reel"] else 0.0
            d["p_engage"] = (d["engage"] / max_total) * 100.0 if d["engage"] else 0.0
            d["p_reste"] = (d["reste"] / max_total) * 100.0 if d["reste"] else 0.0
    else:
        for d in lignes:
            d["p_reel"] = d["p_engage"] = d["p_reste"] = 0.0

    # Totaux synthétiques (charges et produits séparés)
    totals = {
        "demande": float(sub.montant_demande or 0),
        "attribue": float(sub.montant_attribue or 0),
        "recu": float(sub.montant_recu or 0),
        "base_charges": 0.0,
        "base_produits": 0.0,
        "reel_charges": 0.0,
        "reel_produits": 0.0,
        "engage": 0.0,
        "reste": 0.0,
    }
    for d in lignes:
        if d["nature"] == "produit":
            totals["base_produits"] += d["base"]
            totals["reel_produits"] += d["reel"]
        else:
            totals["base_charges"] += d["base"]
            totals["reel_charges"] += d["reel"]
            totals["engage"] += d["engage"]
            totals["reste"] += d["reste"]

    # Arrondis des totaux
    for k in totals:
        totals[k] = round(float(totals[k] or 0), 2)

    return render_template(
        "subvention_bilan.html",
        sub=sub,
        lignes=lignes,
        totals=totals,
        max_total=max_total,
    )



# ---------------------------------------------------------------------
# RBAC Test (diagnostic)
# ---------------------------------------------------------------------
@bp.route("/rbac-test")
@login_required
def rbac_test():
    """Page de diagnostic RBAC.
    Affiche les rôles/perms effectifs de l'utilisateur connecté (et le legacy role si présent).
    """
    from app.models import Permission  # import local pour éviter les cycles

    # Rôles RBAC (selon ton implémentation, role_codes peut être une méthode ou une propriété)
    role_codes_val = []
    if hasattr(current_user, "role_codes"):
        rc = getattr(current_user, "role_codes")
        try:
            role_codes_val = rc() if callable(rc) else list(rc or [])
        except Exception:
            try:
                role_codes_val = list(rc or [])
            except Exception:
                role_codes_val = []

    # legacy role string (compat)
    legacy_role = getattr(current_user, "role", None)

    perms = Permission.query.order_by(Permission.category.asc(), Permission.code.asc()).all()

    has_perm_fn = getattr(current_user, "has_perm", None)
    def _has(code: str) -> bool:
        try:
            return bool(has_perm_fn(code)) if callable(has_perm_fn) else False
        except Exception:
            return False

    perms_by_cat = {}
    for p in perms:
        cat = getattr(p, "category", None) or "Autre"
        perms_by_cat.setdefault(cat, []).append({
            "code": p.code,
            "label": getattr(p, "label", "") or p.code,
            "granted": _has(p.code),
        })

    # stats rapides
    total = sum(len(v) for v in perms_by_cat.values())
    granted = sum(1 for v in perms_by_cat.values() for x in v if x["granted"])

    user_info = {
        "id": getattr(current_user, "id", None),
        "email": getattr(current_user, "email", None),
        "nom": getattr(current_user, "nom", None),
        "secteur_assigne": getattr(current_user, "secteur_assigne", None),
        "legacy_role": legacy_role,
        "role_codes": role_codes_val,
        "perms_total": total,
        "perms_granted": granted,
    }

    return render_template(
        "rbac_test.html",
        user_info=user_info,
        perms_by_cat=perms_by_cat,
        total_perms=total,
        granted_perms=granted,
    )
