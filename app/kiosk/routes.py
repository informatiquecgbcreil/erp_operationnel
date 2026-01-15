import os
import base64
import secrets
from datetime import datetime, date

from flask import (
    render_template,
    request,
    redirect,
    url_for,
    abort,
    current_app,
    jsonify,
    flash,
)

from app.extensions import db
from app.models import (
    SessionActivite,
    AtelierActivite,
    Participant,
    PresenceActivite,
    Quartier,
    AtelierCapaciteMois,
)

from . import bp
from app.activite.services.docx_utils import generate_individuel_mensuel_docx


def _ensure_seed_quartiers():
    if Quartier.query.count() > 0:
        return
    seeds = [
        ("Creil", "Rouher", True),
        ("Creil", "Hauts de Creil", True),
        ("Creil", "Autre Creil", False),
        ("Creil", "Je ne sais pas", False),
    ]
    for ville, nom, is_qpv in seeds:
        db.session.add(Quartier(ville=ville, nom=nom, is_qpv=is_qpv))
    db.session.commit()


def _ensure_month_capacity(atelier: AtelierActivite, session: SessionActivite):
    if atelier.type_atelier != "INDIVIDUEL_MENSUEL":
        return
    if not session.rdv_date:
        return
    annee, mois = session.rdv_date.year, session.rdv_date.month
    cap = AtelierCapaciteMois.query.filter_by(atelier_id=atelier.id, annee=annee, mois=mois).first()
    if cap:
        return
    heures = float(atelier.heures_dispo_defaut_mois or 0.0)
    db.session.add(AtelierCapaciteMois(atelier_id=atelier.id, annee=annee, mois=mois, heures_dispo=heures, locked=False))
    db.session.commit()


def _get_open_session_by_pin(pin: str):
    if not pin:
        return None
    pin = pin.strip()
    if not pin:
        return None
    return (
        SessionActivite.query
        .filter_by(kiosk_open=True, kiosk_pin=pin)
        .filter(SessionActivite.is_deleted.is_(False))
        .order_by(SessionActivite.created_at.desc())
        .first()
    )


def _get_open_session_by_token(token: str):
    if not token:
        return None
    token = token.strip()
    if not token:
        return None
    return (
        SessionActivite.query
        .filter_by(kiosk_open=True, kiosk_token=token)
        .filter(SessionActivite.is_deleted.is_(False))
        .first()
    )


def _session_label(s: SessionActivite):
    atelier = AtelierActivite.query.get(s.atelier_id)
    secteur = s.secteur
    nom = atelier.nom if atelier else "Atelier"
    if s.session_type == "COLLECTIF":
        d = s.date_session.isoformat() if s.date_session else ""
        h = ""
        if s.heure_debut:
            h = s.heure_debut
            if s.heure_fin:
                h += f"-{s.heure_fin}"
        return f"{secteur} — {nom} — {d} {h}".strip()
    else:
        d = s.rdv_date.isoformat() if s.rdv_date else ""
        h = ""
        if s.rdv_debut:
            h = s.rdv_debut
            if s.rdv_fin:
                h += f"-{s.rdv_fin}"
        return f"{secteur} — {nom} — RDV {d} {h}".strip()


@bp.route("/", methods=["GET", "POST"])
def kiosk_home():
    """Page publique: saisie PIN + liste des sessions ouvertes."""
    if request.method == "POST":
        pin = (request.form.get("pin") or "").strip()
        s = _get_open_session_by_pin(pin)
        if not s:
            flash("Code invalide ou session fermée.", "danger")
            return redirect(url_for("kiosk.kiosk_home"))
        return redirect(url_for("kiosk.kiosk_session", token=s.kiosk_token))

    today = date.today()
    # sessions ouvertes du jour (collectif: date_session, individuel: rdv_date)
    sessions = (
        SessionActivite.query.filter_by(kiosk_open=True)
        .filter(SessionActivite.is_deleted.is_(False))
        .order_by(SessionActivite.created_at.desc())
        .limit(300)
        .all()
    )

    filtered = []
    for s in sessions:
        d = s.date_session if s.session_type == "COLLECTIF" else s.rdv_date
        if d == today:
            filtered.append(s)

    # on inverse pour afficher dans l'ordre horaire approximatif
    filtered = list(reversed(filtered))

    entries = []
    for s in filtered:
        atelier = AtelierActivite.query.get(s.atelier_id)
        entries.append({
            "token": s.kiosk_token,
            "pin": s.kiosk_pin,
            "label": _session_label(s),
            "secteur": s.secteur,
            "atelier": atelier.nom if atelier else "Atelier",
            "type": s.session_type,
            "date": (s.date_session or s.rdv_date),
            "debut": s.heure_debut or s.rdv_debut,
            "fin": s.heure_fin or s.rdv_fin,
        })

    return render_template("kiosk/index.html", sessions=entries)


@bp.route("/session/<token>/search")
def kiosk_search(token: str):
    """Recherche participants (protégée par token de session kiosque)."""
    s = _get_open_session_by_token(token)
    if not s:
        return jsonify({"results": []})
    q = (request.args.get("q") or "").strip()
    if len(q) < 2:
        return jsonify({"results": []})

    q_norm = q.replace("%", "").replace("_", "")

    # Recherche simple nom/prénom (LIKE). SQLite: case-insensitive sur ASCII, mais c'est ok.
    candidates = (
        Participant.query.filter(
            (Participant.nom.ilike(f"%{q_norm}%")) | (Participant.prenom.ilike(f"%{q_norm}%"))
        )
        .order_by(Participant.nom.asc(), Participant.prenom.asc())
        .limit(12)
        .all()
    )

    res = []
    for p in candidates:
        label = f"{p.nom} {p.prenom}"
        if p.ville:
            label += f" · {p.ville}"
        res.append({"id": p.id, "label": label})
    return jsonify({"results": res})


@bp.route("/session/<token>", methods=["GET", "POST"])
def kiosk_session(token: str):
    """Page publique d'émargement d'une session précise."""
    _ensure_seed_quartiers()
    s = _get_open_session_by_token(token)
    if not s:
        abort(404)

    atelier = AtelierActivite.query.get_or_404(s.atelier_id)
    motifs = atelier.motifs() or []
    quartiers = Quartier.query.filter_by(ville="Creil").order_by(Quartier.nom.asc()).all()

    message_ok = None

    if request.method == "POST":
        action = request.form.get("action")

        if action == "add_participant":
            nom = (request.form.get("nom") or "").strip()
            prenom = (request.form.get("prenom") or "").strip()
            ville = (request.form.get("ville") or "").strip() or None
            email = (request.form.get("email") or "").strip() or None
            telephone = (request.form.get("telephone") or "").strip() or None
            genre = (request.form.get("genre") or "").strip() or None
            date_naissance = request.form.get("date_naissance") or None
            quartier_id = request.form.get("quartier_id") or None

            if not nom or not prenom:
                flash("Nom et prénom obligatoires.", "danger")
                return redirect(url_for("kiosk.kiosk_session", token=token))

            qid = int(quartier_id) if quartier_id else None
            if (ville or "").strip().lower() != "creil":
                qid = None

            dn = None
            if date_naissance:
                try:
                    dn = datetime.strptime(date_naissance, "%Y-%m-%d").date()
                except Exception:
                    dn = None

            p = Participant(
                nom=nom,
                prenom=prenom,
                ville=ville,
                email=email,
                telephone=telephone,
                genre=genre,
                date_naissance=dn,
                quartier_id=qid,
            )
            db.session.add(p)
            db.session.commit()
            flash("Participant créé. Sélectionne-le ci-dessous puis signe.", "success")
            return redirect(url_for("kiosk.kiosk_session", token=token, highlight=p.id))

        if action == "emarger":
            participant_id = request.form.get("participant_id")
            motif = request.form.get("motif") or None
            motif_autre = (request.form.get("motif_autre") or "").strip() or None
            signature_data = request.form.get("signature_data")

            if not participant_id or str(participant_id).lower() in {"null", "undefined"}:
                flash("Choisis ton nom dans la liste.", "danger")
                return redirect(url_for("kiosk.kiosk_session", token=token))

            participant = Participant.query.get(int(participant_id))
            if not participant:
                flash("Participant introuvable.", "danger")
                return redirect(url_for("kiosk.kiosk_session", token=token))

            sig_path = None
            if signature_data and signature_data.startswith("data:image"):
                try:
                    header, b64data = signature_data.split(",", 1)
                    binary = base64.b64decode(b64data)
                    sig_dir = os.path.join(current_app.instance_path, "signatures_tmp")
                    os.makedirs(sig_dir, exist_ok=True)
                    sig_filename = f"sig_kiosk_s{s.id}_p{participant.id}_{int(datetime.utcnow().timestamp())}.png"
                    sig_path = os.path.join(sig_dir, sig_filename)
                    with open(sig_path, "wb") as f:
                        f.write(binary)
                except Exception:
                    sig_path = None

            try:
                pr = PresenceActivite(
                    session_id=s.id,
                    participant_id=participant.id,
                    motif=motif,
                    motif_autre=motif_autre,
                    signature_path=sig_path,
                )
                db.session.add(pr)
                db.session.commit()
            except Exception:
                db.session.rollback()
                flash("Tu es déjà émargé(e) sur cette session.", "warning")
                return redirect(url_for("kiosk.kiosk_session", token=token))

            # Actions post (individuel mensuel)
            if s.session_type == "INDIVIDUEL_MENSUEL":
                _ensure_month_capacity(atelier, s)
                generate_individuel_mensuel_docx(app=current_app, atelier=atelier, annee=s.rdv_date.year, mois=s.rdv_date.month)

            message_ok = "Merci, c’est bon !"

    highlight = request.args.get("highlight")
    highlight_label = None
    if highlight:
        try:
            hp = Participant.query.get(int(highlight))
            if hp:
                highlight_label = f"{hp.nom} {hp.prenom}" + (f" · {hp.ville}" if hp.ville else "")
        except Exception:
            highlight = None

    label = _session_label(s)

    return render_template(
        "kiosk/session.html",
        session=s,
        atelier=atelier,
        session_label=label,
        motifs=motifs,
        quartiers=quartiers,
        message_ok=message_ok,
        highlight=highlight,
        highlight_label=highlight_label,
        token=token,
    )
