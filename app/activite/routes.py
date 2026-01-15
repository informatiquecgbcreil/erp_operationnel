import os
import base64
import secrets
from datetime import datetime, date

from flask import render_template, request, redirect, url_for, flash, current_app, send_file, abort
from werkzeug.utils import secure_filename
from flask_login import login_required, current_user
from ..rbac import require_perm
from sqlalchemy import or_

from app.extensions import db
from app.models import (
    AtelierActivite,
    SessionActivite,
    Participant,
    PresenceActivite,
    Quartier,
    Referentiel,
    AtelierCapaciteMois,
    ArchiveEmargement,
    Evaluation,
    Objectif,
)

from . import bp
from .services.docx_utils import generate_collectif_docx_pdf, generate_individuel_mensuel_docx, finalize_individuel_mensuel_pdf
from .services.mail_utils import send_email_with_attachment


# ------------------ Helpers ------------------

def _is_admin_global():
    return current_user.is_authenticated and getattr(current_user, "has_role", lambda *_: False)("admin_tech")


def _user_secteur():
    if _is_admin_global():
        return (request.args.get("secteur") or current_user.secteur_assigne or "").strip() or "Numérique"
    # responsable_secteur = admin de son secteur
    return (current_user.secteur_assigne or "").strip() or "Numérique"


def _ensure_seed_quartiers():
    # Seed minimal Creil quartiers if none exist
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


def _load_referentiels():
    return Referentiel.query.order_by(Referentiel.nom.asc()).all()


def _ensure_seed_ateliers(secteur: str):
    """Seed minimal ateliers for a smoother IRL start.

    We keep this extremely conservative to avoid surprising colleagues.
    - Numérique: two INDIVIDUEL_MENSUEL ateliers requested by Antoine
    """
    if not secteur:
        return
    if secteur.strip().lower() != "numérique" and secteur.strip().lower() != "numerique":
        return

    # If the sector already has ateliers, we don't seed anything.
    if AtelierActivite.query.filter_by(secteur=secteur, is_deleted=False).count() > 0:
        return

    # Seed requested ateliers individuels mensuels
    seeds = [
        ("S.O(rdi).S", "INDIVIDUEL_MENSUEL"),
        ("Accès aux droits", "INDIVIDUEL_MENSUEL"),
    ]
    for nom, type_atelier in seeds:
        db.session.add(
            AtelierActivite(
                secteur=secteur,
                nom=nom,
                type_atelier=type_atelier,
                # heures_dispo_defaut_mois left to the sector referent
                heures_dispo_defaut_mois=None,
                # Provide a sensible default motifs list (editable)
                motifs_json=None,
            )
        )
    db.session.commit()


def _safe_unlink(path: str | None) -> None:
    """Supprime un fichier si possible, sans jamais faire planter la requête."""
    if not path:
        return
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


@bp.route("/")
@login_required
def index():
    _ensure_seed_quartiers()
    secteur = _user_secteur()
    _ensure_seed_ateliers(secteur)
    corbeille = (request.args.get("corbeille") == "1")
    q = AtelierActivite.query.filter_by(secteur=secteur)
    if corbeille:
        q = q.filter(AtelierActivite.is_deleted.is_(True))
    else:
        q = q.filter(AtelierActivite.is_deleted.is_(False))
    ateliers = q.order_by(AtelierActivite.nom.asc()).all()
    return render_template(
        "activite/index.html",
        secteur=secteur,
        ateliers=ateliers,
        is_admin_global=_is_admin_global(),
        corbeille=corbeille,
    )


# ------------------ Gestion Participants (par secteur) ------------------


@bp.route("/participants")
@login_required
def participants():
    """Liste des participants ayant au moins une présence dans le secteur."""
    secteur = _user_secteur()
    q = (request.args.get("q") or "").strip()

    # Participants présents dans ce secteur (via sessions)
    base = (
        db.session.query(Participant)
        .join(PresenceActivite, PresenceActivite.participant_id == Participant.id)
        .join(SessionActivite, SessionActivite.id == PresenceActivite.session_id)
        .filter(SessionActivite.secteur == secteur)
        .distinct()
    )
    if q:
        like = f"%{q.lower()}%"
        base = base.filter(
            or_(
                db.func.lower(Participant.nom).like(like),
                db.func.lower(Participant.prenom).like(like),
                db.func.lower(db.func.coalesce(Participant.email, "")).like(like),
                db.func.lower(db.func.coalesce(Participant.telephone, "")).like(like),
            )
        )

    participants_list = base.order_by(Participant.nom.asc(), Participant.prenom.asc()).limit(500).all()

    # Mini-stats par participant dans le secteur (visites + dernière venue)
    stats_map = {}
    rows = (
        db.session.query(
            PresenceActivite.participant_id,
            db.func.count(PresenceActivite.id).label("visites"),
            db.func.max(PresenceActivite.created_at).label("last_seen"),
        )
        .join(SessionActivite, SessionActivite.id == PresenceActivite.session_id)
        .filter(SessionActivite.secteur == secteur)
        .group_by(PresenceActivite.participant_id)
        .all()
    )
    for r in rows:
        stats_map[r.participant_id] = {"visites": int(r.visites or 0), "last_seen": r.last_seen}

    return render_template(
        "activite/participants.html",
        secteur=secteur,
        q=q,
        participants=participants_list,
        stats_map=stats_map,
        is_admin_global=_is_admin_global(),
    )


@bp.route("/participant/<int:participant_id>/edit", methods=["GET", "POST"])
@login_required
def participant_edit(participant_id: int):
    secteur = _user_secteur()
    p = Participant.query.get_or_404(participant_id)

    # Autorisation: doit être "dans" le secteur (au moins une présence) ou admin global
    if not _is_admin_global():
        in_secteur = (
            db.session.query(PresenceActivite.id)
            .join(SessionActivite, SessionActivite.id == PresenceActivite.session_id)
            .filter(PresenceActivite.participant_id == p.id)
            .filter(SessionActivite.secteur == secteur)
            .first()
            is not None
        )
        if not in_secteur:
            flash("Accès refusé.", "danger")
            return redirect(url_for("activite.participants"))

    if request.method == "POST":
        p.nom = (request.form.get("nom") or p.nom).strip()
        p.prenom = (request.form.get("prenom") or p.prenom).strip()
        p.adresse = (request.form.get("adresse") or "").strip() or None
        p.ville = (request.form.get("ville") or "").strip() or None
        p.email = (request.form.get("email") or "").strip() or None
        p.telephone = (request.form.get("telephone") or "").strip() or None
        p.genre = (request.form.get("genre") or "").strip() or None
        p.type_public = (request.form.get("type_public") or p.type_public or "H").strip()[:2]

        dn = (request.form.get("date_naissance") or "").strip()
        if dn:
            try:
                p.date_naissance = datetime.strptime(dn, "%Y-%m-%d").date()
            except Exception:
                flash("Date de naissance invalide.", "warning")
        else:
            p.date_naissance = None

        qid = (request.form.get("quartier_id") or "").strip()
        p.quartier_id = int(qid) if qid.isdigit() else None

        db.session.commit()
        flash("Participant mis à jour.", "success")
        return redirect(url_for("activite.participants"))

    quartiers = Quartier.query.order_by(Quartier.ville.asc(), Quartier.nom.asc()).all()
    return render_template("activite/participant_form.html", secteur=secteur, p=p, quartiers=quartiers)


@bp.route("/participant/<int:participant_id>/anonymize", methods=["POST"])
@login_required
def participant_anonymize(participant_id: int):
    """Anonymise un participant (conserve les stats mais supprime les identifiants)."""
    secteur = _user_secteur()
    p = Participant.query.get_or_404(participant_id)

    if not _is_admin_global():
        in_secteur = (
            db.session.query(PresenceActivite.id)
            .join(SessionActivite, SessionActivite.id == PresenceActivite.session_id)
            .filter(PresenceActivite.participant_id == p.id)
            .filter(SessionActivite.secteur == secteur)
            .first()
            is not None
        )
        if not in_secteur:
            flash("Accès refusé.", "danger")
            return redirect(url_for("activite.participants"))

    # Champs identifiants
    p.nom = "Anonyme"
    p.prenom = "Anonyme"
    p.adresse = None
    p.ville = None
    p.email = None
    p.telephone = None

    # Option : anonymisation stricte (efface aussi démographie)
    strict = (request.form.get("strict") == "1")
    if strict:
        p.genre = None
        p.date_naissance = None
        p.type_public = "H"
        p.quartier_id = None

    db.session.commit()
    flash("Participant anonymisé.", "success")
    return redirect(url_for("activite.participants"))


@bp.route("/participant/<int:participant_id>/delete", methods=["POST"])
@login_required
def participant_delete(participant_id: int):
    """Suppression définitive : uniquement si le participant n'existe pas dans d'autres secteurs.

    (Admin global : bypass.)
    """
    secteur = _user_secteur()
    p = Participant.query.get_or_404(participant_id)

    if not _is_admin_global():
        # Vérifie présence hors secteur
        other = (
            db.session.query(PresenceActivite.id)
            .join(SessionActivite, SessionActivite.id == PresenceActivite.session_id)
            .filter(PresenceActivite.participant_id == p.id)
            .filter(SessionActivite.secteur != secteur)
            .first()
        )
        if other is not None:
            flash("Suppression refusée : ce participant est utilisé dans d'autres secteurs. Utilise 'Anonymiser'.", "warning")
            return redirect(url_for("activite.participants"))

    # Nettoie signatures puis supprime présences
    presences = PresenceActivite.query.filter_by(participant_id=p.id).all()
    for pr in presences:
        if pr.signature_path:
            try:
                # signature_path est un chemin fichier
                if os.path.exists(pr.signature_path):
                    os.remove(pr.signature_path)
            except Exception:
                pass
        db.session.delete(pr)

    db.session.delete(p)
    db.session.commit()
    flash("Participant supprimé définitivement.", "success")
    return redirect(url_for("activite.participants"))


@bp.route("/atelier/new", methods=["GET", "POST"])
@login_required
def atelier_new():
    secteur = _user_secteur()
    if request.method == "POST":
        nom = (request.form.get("nom") or "").strip()
        if not nom:
            flash("Nom d'atelier obligatoire.", "danger")
            referentiels = _load_referentiels()
            return render_template(
                "activite/atelier_form.html",
                secteur=secteur,
                atelier=None,
                referentiels=referentiels,
                selected_competences=set(),
            )

        type_atelier = request.form.get("type_atelier") or "COLLECTIF"
        description = (request.form.get("description") or "").strip() or None
        duree_defaut_minutes = request.form.get("duree_defaut_minutes") or None

        capacite_defaut = request.form.get("capacite_defaut") or None
        heures_dispo_defaut_mois = request.form.get("heures_dispo_defaut_mois") or None

        motifs = [m.strip() for m in (request.form.get("motifs") or "").split(";") if m.strip()]
        motifs_json = None
        if motifs:
            import json

            motifs_json = json.dumps(motifs, ensure_ascii=False)

        a = AtelierActivite(
            secteur=secteur,
            nom=nom,
            description=description,
            type_atelier=type_atelier,
            capacite_defaut=int(capacite_defaut) if capacite_defaut else None,
            heures_dispo_defaut_mois=float(heures_dispo_defaut_mois) if heures_dispo_defaut_mois else None,
            duree_defaut_minutes=int(duree_defaut_minutes) if duree_defaut_minutes else None,
            motifs_json=motifs_json,
        )
        competence_ids = [int(cid) for cid in request.form.getlist("competence_ids") if cid.isdigit()]
        if competence_ids:
            a.competences = Competence.query.filter(Competence.id.in_(competence_ids)).all()
        db.session.add(a)
        db.session.commit()
        flash("Atelier créé.", "success")
        return redirect(url_for("activite.index"))

    referentiels = _load_referentiels()
    return render_template(
        "activite/atelier_form.html",
        secteur=secteur,
        atelier=None,
        referentiels=referentiels,
        selected_competences=set(),
    )




@bp.route("/atelier/<int:atelier_id>/edit", methods=["GET", "POST"])
@login_required
def atelier_edit(atelier_id: int):
    secteur = _user_secteur()
    atelier = AtelierActivite.query.get_or_404(atelier_id)
    if atelier.is_deleted:
        flash("Cet atelier est dans la corbeille. Restaure-le pour le modifier.", "warning")
        return redirect(url_for("activite.index", corbeille=1))
    if not _is_admin_global() and atelier.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    if request.method == "POST":
        atelier.nom = (request.form.get("nom") or atelier.nom).strip()
        atelier.description = (request.form.get("description") or "").strip() or None
        atelier.type_atelier = request.form.get("type_atelier") or atelier.type_atelier

        duree_defaut_minutes = request.form.get("duree_defaut_minutes") or None
        atelier.duree_defaut_minutes = int(duree_defaut_minutes) if duree_defaut_minutes else None

        capacite_defaut = request.form.get("capacite_defaut") or None
        atelier.capacite_defaut = int(capacite_defaut) if capacite_defaut else None

        heures_dispo_defaut_mois = request.form.get("heures_dispo_defaut_mois") or None
        atelier.heures_dispo_defaut_mois = float(heures_dispo_defaut_mois) if heures_dispo_defaut_mois else None

        motifs = [m.strip() for m in (request.form.get("motifs") or "").split(";") if m.strip()]
        if motifs:
            import json

            atelier.motifs_json = json.dumps(motifs, ensure_ascii=False)
        else:
            atelier.motifs_json = None

        competence_ids = [int(cid) for cid in request.form.getlist("competence_ids") if cid.isdigit()]
        if competence_ids:
            atelier.competences = Competence.query.filter(Competence.id.in_(competence_ids)).all()
        else:
            atelier.competences = []

        db.session.commit()
        flash("Atelier mis à jour.", "success")
        return redirect(url_for("activite.index"))

    motifs_str = "; ".join(atelier.motifs() or [])
    referentiels = _load_referentiels()
    selected_competences = {c.id for c in atelier.competences}
    return render_template(
        "activite/atelier_form.html",
        secteur=secteur,
        atelier=atelier,
        motifs_str=motifs_str,
        referentiels=referentiels,
        selected_competences=selected_competences,
    )


@bp.route("/atelier/<int:atelier_id>/sessions")
@login_required
def sessions(atelier_id: int):
    secteur = _user_secteur()
    atelier = AtelierActivite.query.get_or_404(atelier_id)
    corbeille = (request.args.get("corbeille") == "1")
    if atelier.is_deleted and not corbeille:
        flash("Cet atelier est dans la corbeille.", "warning")
        return redirect(url_for("activite.index", corbeille=1))
    if not _is_admin_global() and atelier.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    q = SessionActivite.query.filter_by(atelier_id=atelier.id)
    if corbeille:
        q = q.filter(SessionActivite.is_deleted.is_(True))
    else:
        q = q.filter(SessionActivite.is_deleted.is_(False))
    q = q.order_by(SessionActivite.created_at.desc())
    sessions_list = q.limit(200).all()

    # Fill rates
    session_stats = []
    for s in sessions_list:
        nb = len(s.presences)
        cap = s.capacite or 0
        taux = None
        if s.session_type == "COLLECTIF" and cap > 0:
            taux = round((nb / cap) * 100, 1)
        session_stats.append((s, nb, cap, taux))

    return render_template(
        "activite/sessions.html",
        secteur=secteur,
        atelier=atelier,
        sessions=session_stats,
        corbeille=corbeille,
        current_year=date.today().year,
        current_month=date.today().month,
    )


# ------------------ Suppression / Restauration (soft-delete) ------------------


@bp.route("/atelier/<int:atelier_id>/delete", methods=["POST"])
@login_required
@require_perm('activite:delete')
def atelier_delete(atelier_id: int):
    """Met un atelier (et ses sessions) en corbeille.

    On ne supprime jamais physiquement par défaut (utile en phase de test).
    """
    secteur = _user_secteur()
    atelier = AtelierActivite.query.get_or_404(atelier_id)
    if not _is_admin_global() and atelier.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    if atelier.is_deleted:
        flash("Atelier déjà dans la corbeille.", "info")
        return redirect(url_for("activite.index", corbeille=1))

    atelier.is_deleted = True
    atelier.deleted_at = datetime.utcnow()

    # Met aussi les sessions en corbeille + ferme les kiosques ouverts
    for s in SessionActivite.query.filter_by(atelier_id=atelier.id).all():
        s.is_deleted = True
        s.deleted_at = datetime.utcnow()
        s.kiosk_open = False
        s.kiosk_pin = None
        s.kiosk_token = None

    db.session.commit()
    flash("Atelier placé dans la corbeille (restaurable).", "success")
    return redirect(url_for("activite.index"))


@bp.route("/atelier/<int:atelier_id>/restore", methods=["POST"])
@login_required
def atelier_restore(atelier_id: int):
    """Restaure un atelier (et ses sessions)."""
    secteur = _user_secteur()
    atelier = AtelierActivite.query.get_or_404(atelier_id)
    if not _is_admin_global() and atelier.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    if not atelier.is_deleted:
        flash("Atelier déjà actif.", "info")
        return redirect(url_for("activite.index"))

    atelier.is_deleted = False
    atelier.deleted_at = None
    for s in SessionActivite.query.filter_by(atelier_id=atelier.id).all():
        s.is_deleted = False
        s.deleted_at = None
    db.session.commit()
    flash("Atelier restauré.", "success")
    return redirect(url_for("activite.index"))


@bp.route("/session/<int:session_id>/delete", methods=["POST"])
@login_required
@require_perm('activite:delete')
def session_delete(session_id: int):
    """Met une session/RDV en corbeille."""
    secteur = _user_secteur()
    s = SessionActivite.query.get_or_404(session_id)
    atelier = AtelierActivite.query.get_or_404(s.atelier_id)
    if not _is_admin_global() and s.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    if s.is_deleted:
        flash("Session déjà dans la corbeille.", "info")
        return redirect(url_for("activite.sessions", atelier_id=atelier.id, corbeille=1))

    s.is_deleted = True
    s.deleted_at = datetime.utcnow()
    s.kiosk_open = False
    s.kiosk_pin = None
    s.kiosk_token = None
    db.session.commit()

    flash("Session placée dans la corbeille (restaurable).", "success")
    return redirect(url_for("activite.sessions", atelier_id=atelier.id))


@bp.route("/session/<int:session_id>/restore", methods=["POST"])
@login_required
def session_restore(session_id: int):
    """Restaure une session/RDV."""
    secteur = _user_secteur()
    s = SessionActivite.query.get_or_404(session_id)
    atelier = AtelierActivite.query.get_or_404(s.atelier_id)
    if not _is_admin_global() and s.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))
    if atelier.is_deleted:
        flash("Restaure d'abord l'atelier.", "warning")
        return redirect(url_for("activite.index", corbeille=1))

    if not s.is_deleted:
        flash("Session déjà active.", "info")
        return redirect(url_for("activite.sessions", atelier_id=atelier.id))

    s.is_deleted = False
    s.deleted_at = None
    db.session.commit()
    flash("Session restaurée.", "success")
    return redirect(url_for("activite.sessions", atelier_id=atelier.id))


@bp.route("/session/<int:session_id>/purge", methods=["POST"])
@login_required
@require_perm('activite:purge')
def session_purge(session_id: int):
    """Suppression définitive d'une session.

    - Efface la session
    - Efface ses présences
    - Supprime les fichiers de signatures
    - Supprime les archives (docx/pdf/corrected) rattachées à la session

    Par sécurité, on attend que la session soit déjà dans la corbeille
    (is_deleted=True), sauf admin global.
    """
    secteur = _user_secteur()
    s = SessionActivite.query.get_or_404(session_id)
    atelier = AtelierActivite.query.get_or_404(s.atelier_id)

    if not _is_admin_global() and s.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    if not _is_admin_global() and not s.is_deleted:
        flash("Place d'abord la session dans la corbeille avant suppression définitive.", "warning")
        return redirect(url_for("activite.sessions", atelier_id=atelier.id))

    # 1) signatures + présences
    presences = PresenceActivite.query.filter_by(session_id=s.id).all()
    for pr in presences:
        _safe_unlink(pr.signature_path)
        db.session.delete(pr)

    # 2) archives liées à cette session
    archives = ArchiveEmargement.query.filter_by(session_id=s.id).all()
    for a in archives:
        _safe_unlink(a.docx_path)
        _safe_unlink(a.pdf_path)
        _safe_unlink(a.corrected_docx_path)
        _safe_unlink(a.corrected_pdf_path)
        db.session.delete(a)

    # 3) supprime la session
    db.session.delete(s)
    db.session.commit()
    flash("Session supprimée définitivement.", "success")
    return redirect(url_for("activite.sessions", atelier_id=atelier.id, corbeille=1))


@bp.route("/atelier/<int:atelier_id>/session/new", methods=["GET", "POST"])
@login_required
def session_new(atelier_id: int):
    secteur = _user_secteur()
    atelier = AtelierActivite.query.get_or_404(atelier_id)
    if atelier.is_deleted:
        flash("Cet atelier est dans la corbeille. Restaure-le pour créer une session.", "warning")
        return redirect(url_for("activite.index", corbeille=1))
    if not _is_admin_global() and atelier.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    if request.method == "POST":
        session_type = atelier.type_atelier
        if session_type == "INDIVIDUEL_MENSUEL":
            # RDV
            rdv_date = request.form.get("rdv_date")
            rdv_debut = (request.form.get("rdv_debut") or "").strip() or None
            rdv_fin = (request.form.get("rdv_fin") or "").strip() or None
            if not rdv_date:
                flash("Date RDV obligatoire.", "danger")
                referentiels = _load_referentiels()
                selected_competences = {c.id for c in atelier.competences}
                return render_template(
                    "activite/session_form.html",
                    secteur=secteur,
                    atelier=atelier,
                    session=None,
                    referentiels=referentiels,
                    selected_competences=selected_competences,
                )
            rdv_date_obj = datetime.strptime(rdv_date, "%Y-%m-%d").date()
            s = SessionActivite(
                atelier_id=atelier.id,
                secteur=atelier.secteur,
                session_type="INDIVIDUEL_MENSUEL",
                rdv_date=rdv_date_obj,
                rdv_debut=rdv_debut,
                rdv_fin=rdv_fin,
            )
        else:
            # collectif
            date_session = request.form.get("date_session")
            heure_debut = (request.form.get("heure_debut") or "").strip() or None
            heure_fin = (request.form.get("heure_fin") or "").strip() or None
            capacite = request.form.get("capacite") or atelier.capacite_defaut
            if not date_session:
                flash("Date de session obligatoire.", "danger")
                referentiels = _load_referentiels()
                selected_competences = {c.id for c in atelier.competences}
                return render_template(
                    "activite/session_form.html",
                    secteur=secteur,
                    atelier=atelier,
                    session=None,
                    referentiels=referentiels,
                    selected_competences=selected_competences,
                )
            date_obj = datetime.strptime(date_session, "%Y-%m-%d").date()
            s = SessionActivite(
                atelier_id=atelier.id,
                secteur=atelier.secteur,
                session_type="COLLECTIF",
                date_session=date_obj,
                heure_debut=heure_debut,
                heure_fin=heure_fin,
                capacite=int(capacite) if capacite else None,
            )

        competence_ids = [int(cid) for cid in request.form.getlist("competence_ids") if cid.isdigit()]
        if competence_ids:
            s.competences = Competence.query.filter(Competence.id.in_(competence_ids)).all()
        else:
            s.competences = []

        db.session.add(s)
        db.session.commit()
        flash("Session créée.", "success")
        return redirect(url_for("activite.emargement", session_id=s.id))

    referentiels = _load_referentiels()
    selected_competences = {c.id for c in atelier.competences}
    return render_template(
        "activite/session_form.html",
        secteur=secteur,
        atelier=atelier,
        session=None,
        referentiels=referentiels,
        selected_competences=selected_competences,
    )


@bp.route("/session/<int:session_id>/emargement", methods=["GET", "POST"])
@login_required
def emargement(session_id: int):
    secteur = _user_secteur()
    s = SessionActivite.query.get_or_404(session_id)
    atelier = AtelierActivite.query.get_or_404(s.atelier_id)
    if s.is_deleted or atelier.is_deleted:
        flash("Cette session/atelier est dans la corbeille.", "warning")
        return redirect(url_for("activite.sessions", atelier_id=atelier.id, corbeille=1))
    if not _is_admin_global() and s.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    _ensure_seed_quartiers()
    quartiers = Quartier.query.filter_by(ville="Creil").order_by(Quartier.nom.asc()).all()

    if request.method == "POST":
        action = request.form.get("action")
        if action == "save_evaluation":
            participant_id = request.form.get("participant_id")
            if not participant_id:
                flash("Participant manquant.", "danger")
                return redirect(url_for("activite.emargement", session_id=session_id))
            participant = Participant.query.get(int(participant_id))
            if not participant:
                flash("Participant introuvable.", "danger")
                return redirect(url_for("activite.emargement", session_id=session_id))

            eval_date = s.rdv_date or s.date_session or date.today()
            competence_ids = [int(cid) for cid in request.form.getlist("competence_ids") if cid.isdigit()]
            for comp_id in competence_ids:
                etat = request.form.get(f"etat_{comp_id}")
                if etat is None:
                    continue
                try:
                    etat_value = int(etat)
                except ValueError:
                    continue
                commentaire = (request.form.get(f"commentaire_{comp_id}") or "").strip() or None
                evaluation = Evaluation.query.filter_by(
                    participant_id=participant.id,
                    competence_id=comp_id,
                    session_id=s.id,
                ).first()
                if evaluation:
                    evaluation.etat = etat_value
                    evaluation.commentaire = commentaire
                    evaluation.user_id = current_user.id
                    evaluation.date_evaluation = eval_date
                else:
                    evaluation = Evaluation(
                        participant_id=participant.id,
                        competence_id=comp_id,
                        session_id=s.id,
                        user_id=current_user.id,
                        etat=etat_value,
                        date_evaluation=eval_date,
                        commentaire=commentaire,
                    )
                    db.session.add(evaluation)
            db.session.commit()
            flash("Évaluation enregistrée.", "success")
            return redirect(url_for("activite.emargement", session_id=session_id, highlight=participant.id))

        if action == "bulk_validate":
            eval_date = s.rdv_date or s.date_session or date.today()
            session_objectifs = Objectif.query.filter_by(session_id=s.id, type="operationnel").all()
            session_competences = {comp for obj in session_objectifs for comp in obj.competences}
            presences = PresenceActivite.query.filter_by(session_id=session_id).all()
            for pr in presences:
                for comp in session_competences:
                    evaluation = Evaluation.query.filter_by(
                        participant_id=pr.participant_id,
                        competence_id=comp.id,
                        session_id=s.id,
                    ).first()
                    if evaluation:
                        evaluation.etat = 2
                        evaluation.user_id = current_user.id
                        evaluation.date_evaluation = eval_date
                    else:
                        db.session.add(Evaluation(
                            participant_id=pr.participant_id,
                            competence_id=comp.id,
                            session_id=s.id,
                            user_id=current_user.id,
                            etat=2,
                            date_evaluation=eval_date,
                        ))
            db.session.commit()
            flash("Évaluation rapide appliquée.", "success")
            return redirect(url_for("activite.emargement", session_id=session_id))

        if action == "add_participant":
            nom = (request.form.get("nom") or "").strip()
            prenom = (request.form.get("prenom") or "").strip()
            ville = (request.form.get("ville") or "").strip() or None
            adresse = (request.form.get("adresse") or "").strip() or None
            email = (request.form.get("email") or "").strip() or None
            telephone = (request.form.get("telephone") or "").strip() or None
            genre = (request.form.get("genre") or "").strip() or None
            date_naissance = request.form.get("date_naissance") or None
            type_public = (request.form.get("type_public") or "H").strip().upper() or "H"
            quartier_id = request.form.get("quartier_id") or None

            if not nom or not prenom:
                flash("Nom et prénom obligatoires.", "danger")
                return redirect(url_for("activite.emargement", session_id=session_id))

            dn = None
            if date_naissance:
                try:
                    dn = datetime.strptime(date_naissance, "%Y-%m-%d").date()
                except Exception:
                    dn = None

            qid = int(quartier_id) if quartier_id else None
            if (ville or "").strip().lower() != "creil":
                qid = None

            p = Participant(
                nom=nom,
                prenom=prenom,
                ville=ville,
                adresse=adresse,
                email=email,
                telephone=telephone,
                genre=genre,
                date_naissance=dn,
                quartier_id=qid,
                type_public=type_public,
            )
            db.session.add(p)
            db.session.commit()
            flash("Participant créé.", "success")
            return redirect(url_for("activite.emargement", session_id=session_id, highlight=p.id))

        if action == "emarger":
            participant_id = request.form.get("participant_id")
            motif = request.form.get("motif") or None
            motif_autre = (request.form.get("motif_autre") or "").strip() or None
            signature_data = request.form.get("signature_data")

            if not participant_id:
                flash("Choisis un participant.", "danger")
                return redirect(url_for("activite.emargement", session_id=session_id))
            participant = Participant.query.get(int(participant_id))
            if not participant:
                flash("Participant introuvable.", "danger")
                return redirect(url_for("activite.emargement", session_id=session_id))

            sig_path = None
            if signature_data and signature_data.startswith("data:image"):
                try:
                    header, b64data = signature_data.split(",", 1)
                    binary = base64.b64decode(b64data)
                    sig_dir = os.path.join(current_app.instance_path, "signatures_tmp")
                    os.makedirs(sig_dir, exist_ok=True)
                    sig_filename = f"sig_s{session_id}_p{participant.id}_{int(datetime.utcnow().timestamp())}.png"
                    sig_path = os.path.join(sig_dir, sig_filename)
                    with open(sig_path, "wb") as f:
                        f.write(binary)
                except Exception:
                    sig_path = None

            try:
                pr = PresenceActivite.query.filter_by(session_id=session_id, participant_id=participant.id).first()
                if pr:
                    pr.motif = motif
                    pr.motif_autre = motif_autre
                    if sig_path:
                        pr.signature_path = sig_path
                else:
                    pr = PresenceActivite(
                        session_id=session_id,
                        participant_id=participant.id,
                        motif=motif,
                        motif_autre=motif_autre,
                        signature_path=sig_path,
                    )
                    db.session.add(pr)
                db.session.commit()
            except Exception:
                db.session.rollback()
                flash("Impossible d'enregistrer l'émargement (conflit ou erreur).", "danger")
                return redirect(url_for("activite.emargement", session_id=session_id))

            # Post actions: update monthly docx for individuel
            if s.session_type == "INDIVIDUEL_MENSUEL":
                _ensure_month_capacity(atelier, s)
                generate_individuel_mensuel_docx(app=current_app, atelier=atelier, annee=s.rdv_date.year, mois=s.rdv_date.month)

            flash("Émargement enregistré.", "success")
            return redirect(url_for("activite.emargement", session_id=session_id))

    # participants list for autocomplete
    participants = Participant.query.order_by(Participant.nom.asc(), Participant.prenom.asc()).limit(500).all()
    motifs = atelier.motifs() or []
    presences = PresenceActivite.query.filter_by(session_id=session_id).order_by(PresenceActivite.created_at.asc()).all()
    session_objectifs = Objectif.query.filter_by(session_id=s.id, type="operationnel").order_by(Objectif.created_at.asc()).all()
    objectifs_payload = []
    for obj in session_objectifs:
        competences = sorted(
            obj.competences,
            key=lambda c: ((c.code or "").lower(), (c.nom or "").lower()),
        )
        objectifs_payload.append({"objectif": obj, "competences": competences})
    session_competences = sorted(
        {comp for payload in objectifs_payload for comp in payload["competences"]},
        key=lambda c: ((c.code or "").lower(), (c.nom or "").lower()),
    )
    evaluations = Evaluation.query.filter_by(session_id=s.id).all()
    evaluation_map = {(e.participant_id, e.competence_id): e for e in evaluations}

    return render_template(
        "activite/emargement.html",
        secteur=secteur,
        atelier=atelier,
        session=s,
        participants=participants,
        presences=presences,
        motifs=motifs,
        quartiers=quartiers,
        session_competences=session_competences,
        objectifs_payload=objectifs_payload,
        evaluation_map=evaluation_map,
    )




@bp.route("/session/<int:session_id>/kiosk_open")
@login_required
def kiosk_open(session_id: int):
    """Ouvre le kiosque public pour une session et génère PIN + token."""
    secteur = _user_secteur()
    s = SessionActivite.query.get_or_404(session_id)
    atelier = AtelierActivite.query.get_or_404(s.atelier_id)
    if not _is_admin_global() and s.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    # Token long (imprévisible)
    token = secrets.token_urlsafe(24)
    # PIN court (4 chiffres), unique parmi les sessions ouvertes
    for _ in range(50):
        pin = f"{secrets.randbelow(10000):04d}"
        exists = SessionActivite.query.filter_by(kiosk_open=True, kiosk_pin=pin).first()
        if not exists:
            break
    else:
        pin = None

    s.kiosk_open = True
    s.kiosk_token = token
    s.kiosk_pin = pin
    s.kiosk_opened_at = datetime.utcnow()
    db.session.commit()

    flash(f"Kiosque ouvert (code: {pin}).", "success")
    return redirect(url_for("activite.emargement", session_id=session_id))


@bp.route("/session/<int:session_id>/kiosk_close")
@login_required
def kiosk_close(session_id: int):
    """Ferme le kiosque public pour une session (PIN/token expirent)."""
    secteur = _user_secteur()
    s = SessionActivite.query.get_or_404(session_id)
    if not _is_admin_global() and s.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    s.kiosk_open = False
    s.kiosk_pin = None
    s.kiosk_token = None
    db.session.commit()

    flash("Kiosque fermé.", "success")
    return redirect(url_for("activite.emargement", session_id=session_id))
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
    cap = AtelierCapaciteMois(atelier_id=atelier.id, annee=annee, mois=mois, heures_dispo=heures, locked=False)
    db.session.add(cap)
    db.session.commit()


@bp.route("/session/<int:session_id>/generate_collectif")
@login_required
def generate_collectif(session_id: int):
    secteur = _user_secteur()
    s = SessionActivite.query.get_or_404(session_id)
    atelier = AtelierActivite.query.get_or_404(s.atelier_id)
    if s.session_type != "COLLECTIF":
        flash("Uniquement pour les sessions collectives.", "warning")
        return redirect(url_for("activite.emargement", session_id=session_id))
    if not _is_admin_global() and s.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    out_docx, out_pdf = generate_collectif_docx_pdf(app=current_app, atelier=atelier, session=s)
    # register archive
    annee = (s.date_session.year if s.date_session else datetime.utcnow().year)
    mois = (s.date_session.month if s.date_session else datetime.utcnow().month)
    arch = ArchiveEmargement.query.filter_by(atelier_id=atelier.id, session_id=s.id, annee=annee, mois=mois).first()
    if not arch:
        arch = ArchiveEmargement(secteur=atelier.secteur, atelier_id=atelier.id, session_id=s.id, annee=annee, mois=mois)
        db.session.add(arch)
    arch.docx_path = out_docx
    arch.pdf_path = out_pdf
    arch.status = "locked"
    db.session.commit()

    if out_pdf and os.path.exists(out_pdf):
        return send_file(out_pdf, as_attachment=True)
    if out_docx and os.path.exists(out_docx):
        return send_file(out_docx, as_attachment=True)
    flash("Génération échouée.", "danger")
    return redirect(url_for("activite.emargement", session_id=session_id))


def _best_archive_path(arch: ArchiveEmargement, kind: str) -> str | None:
    """Return the best available file path for download/email.

    kind: 'docx' or 'pdf'
    Prefers corrected version if present.
    """
    if not arch:
        return None
    if kind == "pdf":
        return arch.corrected_pdf_path or arch.pdf_path
    return arch.corrected_docx_path or arch.docx_path


@bp.route("/session/<int:session_id>/archive/<string:kind>")
@login_required
def download_collectif_archive(session_id: int, kind: str):
    """Téléchargement : DOCX/PDF, en préférant la version corrigée si dispo."""
    if kind not in {"docx", "pdf"}:
        abort(404)
    secteur = _user_secteur()
    s = SessionActivite.query.get_or_404(session_id)
    atelier = AtelierActivite.query.get_or_404(s.atelier_id)
    if not _is_admin_global() and s.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    # Ensure archive exists (generate if missing)
    annee = (s.date_session.year if s.date_session else datetime.utcnow().year)
    mois = (s.date_session.month if s.date_session else datetime.utcnow().month)
    arch = ArchiveEmargement.query.filter_by(atelier_id=atelier.id, session_id=s.id, annee=annee, mois=mois).first()
    if not arch or not arch.docx_path:
        out_docx, out_pdf = generate_collectif_docx_pdf(app=current_app, atelier=atelier, session=s)
        if not arch:
            arch = ArchiveEmargement(secteur=atelier.secteur, atelier_id=atelier.id, session_id=s.id, annee=annee, mois=mois)
            db.session.add(arch)
        arch.docx_path = out_docx
        arch.pdf_path = out_pdf
        db.session.commit()

    path = _best_archive_path(arch, kind)
    if path and os.path.exists(path):
        return send_file(path, as_attachment=True)
    flash("Fichier introuvable (génère d'abord, ou LibreOffice manque pour le PDF).", "warning")
    return redirect(url_for("activite.emargement", session_id=session_id))


@bp.route("/session/<int:session_id>/archive/upload", methods=["POST"])
@login_required
def upload_collectif_corrected(session_id: int):
    """Upload d'une version corrigée (DOCX ou PDF) pour une session collective."""
    secteur = _user_secteur()
    s = SessionActivite.query.get_or_404(session_id)
    atelier = AtelierActivite.query.get_or_404(s.atelier_id)
    if not _is_admin_global() and s.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    f = request.files.get("file")
    if not f or not f.filename:
        flash("Aucun fichier.", "warning")
        return redirect(url_for("activite.emargement", session_id=session_id))

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in {".docx", ".pdf"}:
        flash("Uniquement .docx ou .pdf", "warning")
        return redirect(url_for("activite.emargement", session_id=session_id))

    annee = (s.date_session.year if s.date_session else datetime.utcnow().year)
    mois = (s.date_session.month if s.date_session else datetime.utcnow().month)
    arch = ArchiveEmargement.query.filter_by(atelier_id=atelier.id, session_id=s.id, annee=annee, mois=mois).first()
    if not arch:
        arch = ArchiveEmargement(secteur=atelier.secteur, atelier_id=atelier.id, session_id=s.id, annee=annee, mois=mois)
        db.session.add(arch)
        db.session.commit()

    # store next to generated file if possible, else in instance/archives_emargements
    base_dir = None
    base_doc = arch.docx_path or ""
    if base_doc and os.path.exists(base_doc):
        base_dir = os.path.dirname(base_doc)
    else:
        base_dir = os.path.join(current_app.instance_path, "archives_emargements")
        os.makedirs(base_dir, exist_ok=True)

    safe = secure_filename(os.path.splitext(f.filename)[0])
    out_name = f"CORRIGE__{safe}{ext}"
    out_path = os.path.join(base_dir, out_name)
    f.save(out_path)

    if ext == ".docx":
        arch.corrected_docx_path = out_path
    else:
        arch.corrected_pdf_path = out_path
    db.session.commit()

    flash("Version corrigée enregistrée.", "success")
    return redirect(url_for("activite.emargement", session_id=session_id))


@bp.route("/session/<int:session_id>/archive/email", methods=["POST"])
@login_required
def email_collectif_archive(session_id: int):
    """Envoie par mail le meilleur fichier (PDF si possible, sinon DOCX)."""
    secteur = _user_secteur()
    s = SessionActivite.query.get_or_404(session_id)
    atelier = AtelierActivite.query.get_or_404(s.atelier_id)
    if not _is_admin_global() and s.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    to = (request.form.get("to") or "").strip()
    if not to:
        flash("Email destinataire manquant.", "warning")
        return redirect(url_for("activite.emargement", session_id=session_id))

    annee = (s.date_session.year if s.date_session else datetime.utcnow().year)
    mois = (s.date_session.month if s.date_session else datetime.utcnow().month)
    arch = ArchiveEmargement.query.filter_by(atelier_id=atelier.id, session_id=s.id, annee=annee, mois=mois).first()
    if not arch or not arch.docx_path:
        out_docx, out_pdf = generate_collectif_docx_pdf(app=current_app, atelier=atelier, session=s)
        if not arch:
            arch = ArchiveEmargement(secteur=atelier.secteur, atelier_id=atelier.id, session_id=s.id, annee=annee, mois=mois)
            db.session.add(arch)
        arch.docx_path = out_docx
        arch.pdf_path = out_pdf
        db.session.commit()

    # prefer PDF, else DOCX
    attachment = _best_archive_path(arch, "pdf") or _best_archive_path(arch, "docx")
    if not attachment or not os.path.exists(attachment):
        flash("Aucun document à envoyer.", "warning")
        return redirect(url_for("activite.emargement", session_id=session_id))

    cfg = current_app.config
    if not cfg.get("MAIL_HOST") or not cfg.get("MAIL_SENDER"):
        flash("SMTP non configuré (MAIL_HOST/MAIL_SENDER).", "warning")
        return redirect(url_for("activite.emargement", session_id=session_id))

    subject = request.form.get("subject") or f"Émargement - {atelier.secteur} - {atelier.nom} - {annee}-{mois:02d}"
    body = request.form.get("body") or "Ci-joint le document d'émargement."

    try:
        send_email_with_attachment(
            host=cfg.get("MAIL_HOST"),
            port=int(cfg.get("MAIL_PORT", 587)),
            username=cfg.get("MAIL_USERNAME") or None,
            password=cfg.get("MAIL_PASSWORD") or None,
            use_tls=bool(cfg.get("MAIL_USE_TLS", True)),
            sender=cfg.get("MAIL_SENDER"),
            to=to,
            subject=subject,
            body=body,
            attachment_path=attachment,
        )
        arch.last_emailed_to = to
        arch.last_emailed_at = datetime.utcnow()
        db.session.commit()
        flash("Email envoyé.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Échec envoi mail : {e}", "danger")
    return redirect(url_for("activite.emargement", session_id=session_id))


@bp.route("/atelier/<int:atelier_id>/individuel/<int:annee>/<int:mois>/docx")
@login_required
def download_individuel_docx(atelier_id: int, annee: int, mois: int):
    # Compat backward: route historique => docx
    return download_individuel_archive(atelier_id=atelier_id, annee=annee, mois=mois, kind="docx")


@bp.route("/atelier/<int:atelier_id>/individuel/<int:annee>/<int:mois>/archive/<string:kind>")
@login_required
def download_individuel_archive(atelier_id: int, annee: int, mois: int, kind: str):
    """Téléchargement : DOCX/PDF, en préférant la version corrigée si dispo."""
    if kind not in {"docx", "pdf"}:
        abort(404)
    secteur = _user_secteur()
    atelier = AtelierActivite.query.get_or_404(atelier_id)
    if atelier.type_atelier != "INDIVIDUEL_MENSUEL":
        flash("Atelier non individuel mensuel.", "warning")
        return redirect(url_for("activite.index"))
    if not _is_admin_global() and atelier.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    arch = ArchiveEmargement.query.filter_by(atelier_id=atelier.id, session_id=None, annee=annee, mois=mois).first()
    if not arch or not arch.docx_path:
        out_docx = generate_individuel_mensuel_docx(app=current_app, atelier=atelier, annee=annee, mois=mois)
        out_pdf = None
        if out_docx:
            out_pdf = finalize_individuel_mensuel_pdf(app=current_app, atelier=atelier, annee=annee, mois=mois)
        if not arch:
            arch = ArchiveEmargement(secteur=atelier.secteur, atelier_id=atelier.id, session_id=None, annee=annee, mois=mois)
            db.session.add(arch)
        arch.docx_path = out_docx
        arch.pdf_path = out_pdf
        db.session.commit()

    path = _best_archive_path(arch, kind)
    if path and os.path.exists(path):
        return send_file(path, as_attachment=True)
    flash("Fichier introuvable (LibreOffice manquant pour le PDF ?).", "warning")
    return redirect(url_for("activite.sessions", atelier_id=atelier_id))


@bp.route("/atelier/<int:atelier_id>/individuel/<int:annee>/<int:mois>/archive/upload", methods=["POST"])
@login_required
def upload_individuel_corrected(atelier_id: int, annee: int, mois: int):
    """Upload d'une version corrigée (DOCX ou PDF) pour le mensuel individuel."""
    secteur = _user_secteur()
    atelier = AtelierActivite.query.get_or_404(atelier_id)
    if atelier.type_atelier != "INDIVIDUEL_MENSUEL":
        flash("Atelier non individuel mensuel.", "warning")
        return redirect(url_for("activite.index"))
    if not _is_admin_global() and atelier.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    f = request.files.get("file")
    if not f or not f.filename:
        flash("Aucun fichier.", "warning")
        return redirect(url_for("activite.sessions", atelier_id=atelier_id))
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in {".docx", ".pdf"}:
        flash("Uniquement .docx ou .pdf", "warning")
        return redirect(url_for("activite.sessions", atelier_id=atelier_id))

    arch = ArchiveEmargement.query.filter_by(atelier_id=atelier.id, session_id=None, annee=annee, mois=mois).first()
    if not arch:
        arch = ArchiveEmargement(secteur=atelier.secteur, atelier_id=atelier.id, session_id=None, annee=annee, mois=mois)
        db.session.add(arch)
        db.session.commit()

    base_dir = None
    base_doc = arch.docx_path or ""
    if base_doc and os.path.exists(base_doc):
        base_dir = os.path.dirname(base_doc)
    else:
        base_dir = os.path.join(current_app.instance_path, "archives_emargements")
        os.makedirs(base_dir, exist_ok=True)

    safe = secure_filename(os.path.splitext(f.filename)[0])
    out_name = f"CORRIGE__{safe}{ext}"
    out_path = os.path.join(base_dir, out_name)
    f.save(out_path)
    if ext == ".docx":
        arch.corrected_docx_path = out_path
    else:
        arch.corrected_pdf_path = out_path
    db.session.commit()
    flash("Version corrigée enregistrée.", "success")
    return redirect(url_for("activite.sessions", atelier_id=atelier_id))


@bp.route("/atelier/<int:atelier_id>/individuel/<int:annee>/<int:mois>/archive/email", methods=["POST"])
@login_required
def email_individuel_archive(atelier_id: int, annee: int, mois: int):
    """Envoie par mail le meilleur fichier (PDF si possible, sinon DOCX)."""
    secteur = _user_secteur()
    atelier = AtelierActivite.query.get_or_404(atelier_id)
    if atelier.type_atelier != "INDIVIDUEL_MENSUEL":
        flash("Atelier non individuel mensuel.", "warning")
        return redirect(url_for("activite.index"))
    if not _is_admin_global() and atelier.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    to = (request.form.get("to") or "").strip()
    if not to:
        flash("Email destinataire manquant.", "warning")
        return redirect(url_for("activite.sessions", atelier_id=atelier_id))

    arch = ArchiveEmargement.query.filter_by(atelier_id=atelier.id, session_id=None, annee=annee, mois=mois).first()
    if not arch or not arch.docx_path:
        out_docx = generate_individuel_mensuel_docx(app=current_app, atelier=atelier, annee=annee, mois=mois)
        out_pdf = finalize_individuel_mensuel_pdf(app=current_app, atelier=atelier, annee=annee, mois=mois)
        if not arch:
            arch = ArchiveEmargement(secteur=atelier.secteur, atelier_id=atelier.id, session_id=None, annee=annee, mois=mois)
            db.session.add(arch)
        arch.docx_path = out_docx
        arch.pdf_path = out_pdf
        db.session.commit()

    attachment = _best_archive_path(arch, "pdf") or _best_archive_path(arch, "docx")
    if not attachment or not os.path.exists(attachment):
        flash("Aucun document à envoyer.", "warning")
        return redirect(url_for("activite.sessions", atelier_id=atelier_id))

    cfg = current_app.config
    if not cfg.get("MAIL_HOST") or not cfg.get("MAIL_SENDER"):
        flash("SMTP non configuré (MAIL_HOST/MAIL_SENDER).", "warning")
        return redirect(url_for("activite.sessions", atelier_id=atelier_id))

    subject = request.form.get("subject") or f"Émargement - {atelier.secteur} - {atelier.nom} - {annee}-{mois:02d}"
    body = request.form.get("body") or "Ci-joint le document d'émargement."

    try:
        send_email_with_attachment(
            host=cfg.get("MAIL_HOST"),
            port=int(cfg.get("MAIL_PORT", 587)),
            username=cfg.get("MAIL_USERNAME") or None,
            password=cfg.get("MAIL_PASSWORD") or None,
            use_tls=bool(cfg.get("MAIL_USE_TLS", True)),
            sender=cfg.get("MAIL_SENDER"),
            to=to,
            subject=subject,
            body=body,
            attachment_path=attachment,
        )
        arch.last_emailed_to = to
        arch.last_emailed_at = datetime.utcnow()
        db.session.commit()
        flash("Email envoyé.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Échec envoi mail : {e}", "danger")
    return redirect(url_for("activite.sessions", atelier_id=atelier_id))


@bp.route("/atelier/<int:atelier_id>/individuel/<int:annee>/<int:mois>/finalize")
@login_required
def finalize_individuel(atelier_id: int, annee: int, mois: int):
    """Génère le PDF figé (à lancer le 1er du mois suivant par tâche planifiée)."""
    secteur = _user_secteur()
    atelier = AtelierActivite.query.get_or_404(atelier_id)
    if atelier.type_atelier != "INDIVIDUEL_MENSUEL":
        flash("Atelier non individuel mensuel.", "warning")
        return redirect(url_for("activite.index"))
    if not _is_admin_global() and atelier.secteur != secteur:
        flash("Accès refusé.", "danger")
        return redirect(url_for("activite.index"))

    out_docx = generate_individuel_mensuel_docx(app=current_app, atelier=atelier, annee=annee, mois=mois)
    out_pdf = finalize_individuel_mensuel_pdf(app=current_app, atelier=atelier, annee=annee, mois=mois)

    cap = AtelierCapaciteMois.query.filter_by(atelier_id=atelier.id, annee=annee, mois=mois).first()
    if cap:
        cap.locked = True

    arch = ArchiveEmargement.query.filter_by(atelier_id=atelier.id, session_id=None, annee=annee, mois=mois).first()
    if not arch:
        arch = ArchiveEmargement(secteur=atelier.secteur, atelier_id=atelier.id, session_id=None, annee=annee, mois=mois)
        db.session.add(arch)
    arch.docx_path = out_docx
    arch.pdf_path = out_pdf
    arch.status = "locked" if out_pdf else "open"
    db.session.commit()

    if out_pdf and os.path.exists(out_pdf):
        return send_file(out_pdf, as_attachment=True)
    if out_docx and os.path.exists(out_docx):
        flash("PDF non généré (LibreOffice manquant ?). Téléchargement du DOCX.", "warning")
        return send_file(out_docx, as_attachment=True)
    flash("Finalisation échouée.", "danger")
    return redirect(url_for("activite.sessions", atelier_id=atelier_id))