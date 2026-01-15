from __future__ import annotations

from datetime import datetime

from flask import Blueprint, render_template, request, redirect, url_for, flash, abort
from flask_login import login_required, current_user
from ..rbac import require_perm, can
from ..rbac import can_access_secteur

from app.extensions import db
from app.models import Participant, PresenceActivite, SessionActivite, Evaluation


bp = Blueprint("participants", __name__, url_prefix="/participants")


def _current_secteur() -> str:
    return (getattr(current_user, "secteur_assigne", "") or "").strip()


def _is_global_role() -> bool:
    return current_user.has_perm("participants:view_all") or current_user.has_perm("scope:all_secteurs")


def _can_read_participant(p: Participant) -> bool:
    return bool(can('participants:view') or can('participants:edit') or can('participants:delete'))


def _can_edit_participant(p: Participant) -> bool:
    return bool(can('participants:edit'))


def _can_see_participant(p: Participant) -> bool:
    # Visible (ancienne logique) : créé par secteur OU déjà présent dans secteur
    # Utilisé uniquement pour les listings "dans mon secteur", PAS pour l'édition.
    if _is_global_role():
        return True
    if not current_user.has_perm("participants:view_all"):
        sec = _current_secteur()
        if not sec:
            return False
        if (p.created_secteur or "") == sec:
            return True
        has_presence = (
            db.session.query(PresenceActivite.id)
            .join(SessionActivite, SessionActivite.id == PresenceActivite.session_id)
            .filter(PresenceActivite.participant_id == p.id)
            .filter(SessionActivite.secteur == sec)
            .first()
        )
        return bool(has_presence)
    return False


@bp.route("/")
@login_required
def list_participants():
    if False:
        abort(403)

    q = (request.args.get("q") or "").strip()
    scope = (request.args.get("scope") or "").strip()  # ""/secteur, created, annuaire

    participants_q = Participant.query

    # Annuaire global : uniquement si recherche (>=2), sinon on retombe en sectoriel
    if scope == "annuaire" and (not q or len(q) < 2):
        scope = "secteur"

    if not current_user.has_perm("participants:view_all"):
        sec = _current_secteur()
        if not sec:
            abort(403)

        # En mode annuaire (avec recherche), pas de restriction sectorielle
        if not (scope == "annuaire" and q and len(q) >= 2):
            if scope == "created":
                participants_q = participants_q.filter(Participant.created_secteur == sec)
            else:
                # secteur = (créé par secteur) OU (a une présence dans secteur)
                subq_presence_ids = (
                    db.session.query(PresenceActivite.participant_id)
                    .join(SessionActivite, SessionActivite.id == PresenceActivite.session_id)
                    .filter(SessionActivite.secteur == sec)
                    .distinct()
                )
                participants_q = participants_q.filter(
                    (Participant.created_secteur == sec) | (Participant.id.in_(subq_presence_ids))
                )
    else:
        # finance/directrice : option filtre secteur
        if scope == "secteur":
            sec = (request.args.get("secteur") or "").strip()
            if sec:
                participants_q = participants_q.filter(Participant.created_secteur == sec)

    # filtre recherche (tous rôles)
    if q:
        like = f"%{q.lower()}%"
        participants_q = participants_q.filter(
            db.or_(
                db.func.lower(Participant.nom).like(like),
                db.func.lower(Participant.prenom).like(like),
                db.func.lower(db.func.coalesce(Participant.email, "")).like(like),
                db.func.lower(db.func.coalesce(Participant.telephone, "")).like(like),
            )
        )

    items = participants_q.order_by(Participant.nom.asc(), Participant.prenom.asc()).limit(1000).all()
    return render_template(
        "participants/list.html",
        items=items,
        q=q,
        scope=scope,
        secteur=_current_secteur(),
    )


@bp.route("/search")
@login_required
def search_participants():
    """Annuaire global (lecture seule) pour l'auto-complétion côté émargement."""
    if False:
        abort(403)

    q = (request.args.get("q") or "").strip()
    if not q or len(q) < 2:
        return {"items": []}

    like = f"%{q.lower()}%"
    participants_q = Participant.query.filter(
        db.or_(
            db.func.lower(Participant.nom).like(like),
            db.func.lower(Participant.prenom).like(like),
            db.func.lower(db.func.coalesce(Participant.email, "")).like(like),
            db.func.lower(db.func.coalesce(Participant.telephone, "")).like(like),
        )
    )

    items = (
        participants_q.order_by(Participant.nom.asc(), Participant.prenom.asc())
        .limit(30)
        .all()
    )

    def _year(d):
        try:
            return d.year if d else None
        except Exception:
            return None

    return {
        "items": [
            {
                "id": p.id,
                "nom": p.nom,
                "prenom": p.prenom,
                "annee_naissance": _year(getattr(p, "date_naissance", None)),
                "ville": getattr(p, "ville", None),
                "created_secteur": getattr(p, "created_secteur", None),
            }
            for p in items
        ]
    }


@bp.route("/new", methods=["GET", "POST"])
@login_required
def new_participant():
    if False:
        abort(403)

    if request.method == "POST":
        nom = (request.form.get("nom") or "").strip()
        prenom = (request.form.get("prenom") or "").strip()
        if not nom or not prenom:
            flash("Nom et prénom obligatoires.", "err")
            return redirect(url_for("participants.new_participant"))

        p = Participant(
            nom=nom,
            prenom=prenom,
            adresse=(request.form.get("adresse") or "").strip() or None,
            ville=(request.form.get("ville") or "").strip() or None,
            email=(request.form.get("email") or "").strip() or None,
            telephone=(request.form.get("telephone") or "").strip() or None,
            genre=(request.form.get("genre") or "").strip() or None,
            type_public=(request.form.get("type_public") or "H").strip() or "H",
            created_by_user_id=getattr(current_user, "id", None),
            created_secteur=(
                _current_secteur()
                if not current_user.has_perm("participants:view_all")
                else (request.form.get("created_secteur") or "").strip() or None
            ),
        )

        d = (request.form.get("date_naissance") or "").strip()
        if d:
            try:
                p.date_naissance = datetime.strptime(d, "%Y-%m-%d").date()
            except Exception:
                pass

        db.session.add(p)
        db.session.commit()
        flash("Participant créé.", "ok")
        return redirect(url_for("participants.edit_participant", participant_id=p.id))

    return render_template("participants/form.html", item=None, secteur=_current_secteur(), is_editable=True)


@bp.route("/<int:participant_id>/edit", methods=["GET", "POST"])
@login_required
def edit_participant(participant_id: int):
    if False:
        abort(403)

    p = Participant.query.get_or_404(participant_id)

    # Lecture globale autorisée (annuaire), mais édition verrouillée
    if not _can_read_participant(p):
        abort(403)

    is_editable = _can_edit_participant(p)

    if request.method == "POST":
        if not is_editable:
            abort(403)

        p.nom = (request.form.get("nom") or "").strip() or p.nom
        p.prenom = (request.form.get("prenom") or "").strip() or p.prenom
        p.adresse = (request.form.get("adresse") or "").strip() or None
        p.ville = (request.form.get("ville") or "").strip() or None
        p.email = (request.form.get("email") or "").strip() or None
        p.telephone = (request.form.get("telephone") or "").strip() or None
        p.genre = (request.form.get("genre") or "").strip() or None
        p.type_public = (request.form.get("type_public") or p.type_public or "H").strip() or "H"

        d = (request.form.get("date_naissance") or "").strip()
        if d:
            try:
                p.date_naissance = datetime.strptime(d, "%Y-%m-%d").date()
            except Exception:
                pass
        else:
            p.date_naissance = None

        # finance/directrice peuvent requalifier created_secteur
        if _is_global_role():
            p.created_secteur = (request.form.get("created_secteur") or "").strip() or None

        db.session.commit()
        flash("Participant mis à jour.", "ok")
        return redirect(url_for("participants.edit_participant", participant_id=p.id))

    return render_template("participants/form.html", item=p, secteur=_current_secteur(), is_editable=is_editable)


@bp.route("/<int:participant_id>/anonymize", methods=["POST"])
@login_required
def anonymize_participant(participant_id: int):
    if False:
        abort(403)

    p = Participant.query.get_or_404(participant_id)
    if not _can_edit_participant(p):
        abort(403)

    p.nom = "ANONYME"
    p.prenom = f"P{p.id}"
    p.adresse = None
    p.ville = None
    p.email = None
    p.telephone = None

    strict = (request.form.get("strict") or "").strip() == "1"
    if strict and _is_global_role():
        p.genre = None
        p.date_naissance = None
        p.quartier_id = None
        p.type_public = "H"

    db.session.commit()
    flash("Participant anonymisé (les stats sont conservées).", "ok")
    return redirect(url_for("participants.edit_participant", participant_id=p.id))


@bp.route("/<int:participant_id>/delete", methods=["POST"])
@login_required
def delete_participant(participant_id: int):
    if False:
        abort(403)
    if not can('participants:delete'):
        abort(403)

    p = Participant.query.get_or_404(participant_id)
    if not _can_edit_participant(p):
        abort(403)

    # garde-fou : un responsable secteur ne supprime pas si le participant existe ailleurs
    if not current_user.has_perm("participants:view_all"):
        sec = _current_secteur()
        other = (
            db.session.query(PresenceActivite.id)
            .join(SessionActivite, SessionActivite.id == PresenceActivite.session_id)
            .filter(PresenceActivite.participant_id == p.id)
            .filter(SessionActivite.secteur != sec)
            .first()
        )
        if other:
            flash("Suppression refusée : participant présent dans d'autres secteurs. Utiliser 'Anonymiser'.", "err")
            return redirect(url_for("participants.edit_participant", participant_id=p.id))

    db.session.query(PresenceActivite).filter(PresenceActivite.participant_id == p.id).delete(synchronize_session=False)
    db.session.query(Evaluation).filter(Evaluation.participant_id == p.id).delete(synchronize_session=False)
    db.session.delete(p)
    db.session.commit()
    flash("Participant supprimé définitivement.", "warning")
    return redirect(url_for("participants.list_participants"))
