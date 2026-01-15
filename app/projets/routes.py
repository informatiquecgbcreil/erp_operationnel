import os
import json
from datetime import datetime
from flask import Blueprint, render_template, request, redirect, url_for, flash, abort, current_app, send_from_directory
from flask_login import login_required, current_user
from app.rbac import require_perm, can, can_access_secteur
from app.secteurs import get_secteur_labels
from werkzeug.utils import secure_filename

from app.extensions import db
from app.models import (
    Projet,
    ChargeProjet,
    ProduitProjet,
    VentilationProjet,
    Depense,
    Subvention,
    SubventionProjet,
    AtelierActivite,
    ProjetAtelier,
    ProjetIndicateur,
    Competence,
    Objectif,
    objectif_competence,
    projet_competence,
    Referentiel,
)

bp = Blueprint("projets", __name__)

ALLOWED_CR = {"pdf", "doc", "docx", "odt"}


# ---------------------------------------------------------------------
# Helpers Budget AAP (UX)
# ---------------------------------------------------------------------

def _budget_stats(projet_id: int) -> dict:
    """Stats simples pour l'UX du budget AAP.

    Objectif : afficher en haut des pages un résumé (totaux + "reste à...")
    et permettre d'afficher des alertes rapides.
    """
    charges = ChargeProjet.query.filter_by(projet_id=projet_id).all()
    produits = ProduitProjet.query.filter_by(projet_id=projet_id).all()

    total_charges = float(sum((c.montant_previsionnel or 0) for c in charges))
    total_charges_ventile = float(sum((c.ventile or 0) for c in charges))
    total_charges_reste = max(0.0, total_charges - total_charges_ventile)

    total_demande = float(sum((p.montant_demande or 0) for p in produits))
    total_accorde = float(sum((p.montant_accorde or 0) for p in produits))
    total_recu = float(sum((p.montant_recu or 0) for p in produits))

    total_produits_ventile = float(sum((p.ventile or 0) for p in produits))
    total_produits_reste = float(sum((p.reste_a_ventiler or 0) for p in produits))

    # Base "produits" utilisée pour donner du contexte (reçu > accordé > demandé)
    base_produits = total_recu if total_recu > 0 else (total_accorde if total_accorde > 0 else total_demande)
    base_label = "reçu" if total_recu > 0 else ("accordé" if total_accorde > 0 else "demandé")

    def pct(a: float, b: float) -> int:
        if b <= 0:
            return 0
        v = int(round((a / b) * 100))
        return max(0, min(100, v))

    return {
        "total_charges": total_charges,
        "total_charges_ventile": total_charges_ventile,
        "total_charges_reste": total_charges_reste,
        "total_demande": total_demande,
        "total_accorde": total_accorde,
        "total_recu": total_recu,
        "base_produits": base_produits,
        "base_label": base_label,
        "total_produits_ventile": total_produits_ventile,
        "total_produits_reste": total_produits_reste,
        "pct_charges_financees": pct(total_charges_ventile, total_charges),
        "pct_produits_ventiles": pct(total_produits_ventile, base_produits if base_produits > 0 else total_charges),
    }


INDICATOR_TEMPLATES = {
    'participants_uniques': 'Participants uniques',
    'presences_totales': 'Présences totales',
    'sessions_totales': 'Sessions réalisées',
    'recurrence_2plus': 'Participants récurrents (≥2 séances)',
    'depenses_totales': 'Dépenses totales (charges)',
    'recettes_totales': 'Recettes totales (produits)',
    'cout_par_participant': 'Coût par participant',
    'cout_par_presence': 'Coût par présence',
}


INDICATOR_PACKS = {
    "caf_base": {
        "label": "Pack CAF (base)",
        "codes": ["participants_uniques", "presences_totales", "sessions_totales", "recurrence_2plus"],
    },
    "financier": {
        "label": "Pack Financier",
        "codes": ["depenses_totales", "recettes_totales", "cout_par_participant", "cout_par_presence"],
    },
    "jeunesse": {
        "label": "Pack Jeunesse (simple)",
        "codes": ["participants_uniques", "recurrence_2plus"],
    },
}

PERIOD_CHOICES = {
    "context": "Période sélectionnée (défaut)",
    "year": "Année sélectionnée",
    "custom": "Personnalisée (dates)",
}

TARGET_OP_CHOICES = {
    "ge": "Atteindre au moins (≥)",
    "le": "Ne pas dépasser (≤)",
}


def can_see_secteur(secteur: str) -> bool:
    return can_access_secteur(secteur)

def ensure_projets_folder():
    folder = os.path.join(current_app.root_path, "..", "static", "uploads", "projets")
    folder = os.path.abspath(folder)
    os.makedirs(folder, exist_ok=True)
    return folder

def allowed_cr(filename: str) -> bool:
    if not filename or "." not in filename:
        return False
    ext = filename.rsplit(".", 1)[1].lower()
    return ext in ALLOWED_CR

@bp.route("/projets")
@login_required
@require_perm("projets:view")
def projets_list():
    q = Projet.query
    if not can("scope:all_secteurs"):
        q = q.filter(Projet.secteur == current_user.secteur_assigne)

    projets = q.order_by(Projet.created_at.desc()).all()
    # Source unique: DB (fallback config) via get_secteur_labels
    secteurs = get_secteur_labels(active_only=False)
    return render_template("projets_list.html", projets=projets, secteurs=secteurs)

@bp.route("/projets/new", methods=["GET", "POST"])
@login_required
@require_perm("projets:edit")
def projets_new():
    secteurs = get_secteur_labels(active_only=True)

    if request.method == "POST":
        nom = (request.form.get("nom") or "").strip()
        secteur = (request.form.get("secteur") or "").strip()
        description = (request.form.get("description") or "").strip()

        if not can("scope:all_secteurs"):
            secteur = current_user.secteur_assigne

        if not nom or not secteur:
            flash("Nom + secteur obligatoires.", "danger")
            return redirect(url_for("projets.projets_new"))

        if not can_see_secteur(secteur):
            abort(403)

        p = Projet(nom=nom, secteur=secteur, description=description)
        db.session.add(p)
        db.session.commit()

        flash("Projet créé.", "success")
        return redirect(url_for("projets.projets_edit", projet_id=p.id))

    return render_template("projets_new.html", secteurs=secteurs)


@bp.route("/projets/<int:projet_id>", methods=["GET", "POST"])
@login_required
@require_perm("projets:view")
def projets_edit(projet_id):
    p = Projet.query.get_or_404(projet_id)
    if not can_see_secteur(p.secteur):
        abort(403)

    if request.method == "POST":
        if not can("projets:edit"):
            abort(403)
        action = request.form.get("action") or ""

        if action == "update":
            p.nom = (request.form.get("nom") or "").strip()
            p.description = (request.form.get("description") or "").strip()

            if not p.nom:
                flash("Nom obligatoire.", "danger")
                return redirect(url_for("projets.projets_edit", projet_id=p.id))

            db.session.commit()
            flash("Projet modifié.", "success")
            return redirect(url_for("projets.projets_edit", projet_id=p.id))

        if action == "update_competences":
            competence_ids = [int(cid) for cid in request.form.getlist("competence_ids") if cid.isdigit()]
            if competence_ids:
                p.competences = Competence.query.filter(Competence.id.in_(competence_ids)).all()
            else:
                p.competences = []
            db.session.commit()
            flash("Compétences du projet mises à jour.", "success")
            return redirect(url_for("projets.projets_edit", projet_id=p.id))

        if action == "upload_cr":
            if not can("projets:files"):
                abort(403)
            file = request.files.get("cr_file")
            if not file or not file.filename:
                flash("Aucun fichier.", "danger")
                return redirect(url_for("projets.projets_edit", projet_id=p.id))

            if not allowed_cr(file.filename):
                flash("Type autorisé : pdf/doc/docx/odt", "danger")
                return redirect(url_for("projets.projets_edit", projet_id=p.id))

            folder = ensure_projets_folder()
            safe_original = secure_filename(file.filename)
            ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
            stored = secure_filename(f"P{p.id}_{ts}_{safe_original}")
            file.save(os.path.join(folder, stored))

            p.cr_filename = stored
            p.cr_original_name = safe_original
            db.session.commit()

            flash("Compte-rendu uploadé.", "success")
            return redirect(url_for("projets.projets_edit", projet_id=p.id))

        if action == "toggle_subvention":
            if not can("subventions:link"):
                abort(403)
            sub_id = int(request.form.get("subvention_id") or 0)
            s = Subvention.query.get_or_404(sub_id)

            if s.secteur != p.secteur:
                abort(400)

            link = SubventionProjet.query.filter_by(projet_id=p.id, subvention_id=s.id).first()
            if link:
                db.session.delete(link)
                db.session.commit()
                flash("Subvention retirée du projet.", "warning")
            else:
                db.session.add(SubventionProjet(projet_id=p.id, subvention_id=s.id))
                db.session.commit()
                flash("Subvention ajoutée au projet.", "success")

            return redirect(url_for("projets.projets_edit", projet_id=p.id))

        # ---- Liens projet <-> ateliers ----
        if action == "toggle_atelier":
            if not can("ateliers:edit"):
                abort(403)
            atelier_id = int(request.form.get("atelier_id") or 0)
            a = AtelierActivite.query.get_or_404(atelier_id)
            if a.secteur != p.secteur or a.is_deleted:
                abort(400)

            link = ProjetAtelier.query.filter_by(projet_id=p.id, atelier_id=a.id).first()
            if link:
                db.session.delete(link)
                db.session.commit()
                flash("Atelier délié du projet.", "warning")
            else:
                db.session.add(ProjetAtelier(projet_id=p.id, atelier_id=a.id))
                db.session.commit()
                flash("Atelier lié au projet.", "success")

            return redirect(url_for("projets.projets_edit", projet_id=p.id))

        # ---- Indicateurs projet ----

        if action == "add_pack":
            pack = (request.form.get("pack") or "").strip()
            cfg = INDICATOR_PACKS.get(pack)
            if not cfg:
                flash("Pack invalide.", "danger")
                return redirect(url_for("projets.projets_edit", projet_id=p.id))

            added = 0
            for code in cfg["codes"]:
                if code not in INDICATOR_TEMPLATES:
                    continue
                exists = ProjetIndicateur.query.filter_by(projet_id=p.id, code=code).first()
                if exists:
                    continue
                db.session.add(ProjetIndicateur(
                    projet_id=p.id,
                    code=code,
                    label=INDICATOR_TEMPLATES.get(code, code),
                    is_active=True,
                    params_json=None,
                ))
                added += 1
            db.session.commit()
            flash(f"Pack ajouté ({added} indicateur(s)).", "success")
            return redirect(url_for("projets.projets_edit", projet_id=p.id))

        if action == "add_indicateur":
            code = (request.form.get("code") or "").strip()
            label = (request.form.get("label") or "").strip()
            if code not in INDICATOR_TEMPLATES:
                flash("Indicateur invalide.", "danger")
                return redirect(url_for("projets.projets_edit", projet_id=p.id))
            if not label:
                label = INDICATOR_TEMPLATES[code]

            exists = ProjetIndicateur.query.filter_by(projet_id=p.id, code=code).first()
            if exists:
                flash("Indicateur déjà présent pour ce projet.", "warning")
                return redirect(url_for("projets.projets_edit", projet_id=p.id))

            db.session.add(ProjetIndicateur(projet_id=p.id, code=code, label=label, is_active=True))
            db.session.commit()
            flash("Indicateur ajouté.", "success")
            return redirect(url_for("projets.projets_edit", projet_id=p.id))

        if action == "toggle_indicateur":
            indic_id = int(request.form.get("indicateur_id") or 0)
            ind = ProjetIndicateur.query.get_or_404(indic_id)
            if ind.projet_id != p.id:
                abort(400)
            ind.is_active = not bool(ind.is_active)
            db.session.commit()
            flash("Indicateur mis à jour.", "success")
            return redirect(url_for("projets.projets_edit", projet_id=p.id))

        if action == "save_indicateur":
            indic_id = int(request.form.get("indicateur_id") or 0)
            ind = ProjetIndicateur.query.get_or_404(indic_id)
            if ind.projet_id != p.id:
                abort(400)

            # label editable (optionnel)
            label = (request.form.get("label") or "").strip()
            if label:
                ind.label = label

            period = (request.form.get("period") or "context").strip()
            if period not in PERIOD_CHOICES:
                period = "context"

            target_raw = (request.form.get("target") or "").strip().replace(",", ".")
            target = None
            if target_raw:
                try:
                    target = float(target_raw)
                except ValueError:
                    target = None

            target_op = (request.form.get("target_op") or "ge").strip()
            if target_op not in TARGET_OP_CHOICES:
                target_op = "ge"

            atelier_id_raw = (request.form.get("atelier_id") or "").strip()
            atelier_id = None
            if atelier_id_raw:
                try:
                    atelier_id = int(atelier_id_raw)
                except ValueError:
                    atelier_id = None

            # bornes custom
            start = (request.form.get("start") or "").strip()
            end = (request.form.get("end") or "").strip()

            params = ind.params()
            params.update({
                "period": period,
                "target": target,
                "target_op": target_op,
                "atelier_id": atelier_id,
                "start": start if period == "custom" else None,
                "end": end if period == "custom" else None,
            })
            # nettoyage
            if params.get("atelier_id") is None:
                params.pop("atelier_id", None)
            if params.get("target") is None:
                params.pop("target", None)
            if period != "custom":
                params.pop("start", None)
                params.pop("end", None)

            ind.params_json = json.dumps(params, ensure_ascii=False)
            db.session.commit()
            flash("Paramètres de l'indicateur enregistrés.", "success")
            return redirect(url_for("projets.projets_edit", projet_id=p.id))


        if action == "delete_indicateur":
            indic_id = int(request.form.get("indicateur_id") or 0)
            ind = ProjetIndicateur.query.get_or_404(indic_id)
            if ind.projet_id != p.id:
                abort(400)
            db.session.delete(ind)
            db.session.commit()
            flash("Indicateur supprimé.", "warning")
            return redirect(url_for("projets.projets_edit", projet_id=p.id))

        abort(400)

    # ----- GET (lists) -----
    subs_q = Subvention.query.filter_by(est_archive=False).filter(Subvention.secteur == p.secteur)
    subs = subs_q.order_by(Subvention.annee_exercice.desc(), Subvention.nom.asc()).all()
    linked_subs = set(sp.subvention_id for sp in p.subventions)

    ateliers = AtelierActivite.query.filter_by(secteur=p.secteur, is_deleted=False).order_by(AtelierActivite.nom.asc()).all()
    linked_ateliers = set(link.atelier_id for link in ProjetAtelier.query.filter_by(projet_id=p.id).all())

    indicateurs = ProjetIndicateur.query.filter_by(projet_id=p.id).order_by(ProjetIndicateur.created_at.asc()).all()
    referentiels = Referentiel.query.order_by(Referentiel.nom.asc()).all()
    selected_competences = {c.id for c in p.competences}

    return render_template(
        "projets_edit.html",
        projet=p,
        subs=subs,
        linked=linked_subs,
        ateliers=ateliers,
        linked_ateliers=linked_ateliers,
        indicateurs=indicateurs,
        indicator_templates=INDICATOR_TEMPLATES,
        indicator_packs=INDICATOR_PACKS,
        period_choices=PERIOD_CHOICES,
        target_op_choices=TARGET_OP_CHOICES,
        referentiels=referentiels,
        selected_competences=selected_competences,
    )


@bp.route("/projets/<int:projet_id>/delete", methods=["POST"])
@login_required
@require_perm('projets:delete')
def projets_delete(projet_id: int):
    """Suppression définitive d'un projet (et de ses liens).

    On fait du *delete applicatif* (compatible SQLite/Postgres) :
    - objectifs + table pivot objectif_competence
    - liens projet_atelier
    - liens projet_competence (table pivot)
    - indicateurs
    - liens projet<->subvention (SubventionProjet)
    Puis on supprime le projet.

    ⚠️ Non réversible. (On pourrait faire un soft-delete plus tard si tu veux.)
    """
    projet = Projet.query.get_or_404(projet_id)

    if not can_access_secteur(getattr(projet, "secteur", None)):
        flash("Accès refusé.", "danger")
        return redirect(url_for("projets.projets_list"))

    try:
        # 1) Objectifs liés au projet + pivot objectif_competence
        objectif_ids = [o.id for o in Objectif.query.filter_by(projet_id=projet.id).all()]
        if objectif_ids:
            db.session.execute(
                objectif_competence.delete().where(objectif_competence.c.objectif_id.in_(objectif_ids))
            )
            # suppression des objectifs (les enfants sont aussi dans la liste via projet_id)
            Objectif.query.filter(Objectif.id.in_(objectif_ids)).delete(synchronize_session=False)

        # 2) Liens projet/ateliers
        ProjetAtelier.query.filter_by(projet_id=projet.id).delete(synchronize_session=False)

        # 3) Liens projet/compétences (table pivot)
        db.session.execute(
            projet_competence.delete().where(projet_competence.c.projet_id == projet.id)
        )

        # 4) Indicateurs
        ProjetIndicateur.query.filter_by(projet_id=projet.id).delete(synchronize_session=False)

        # 5) Liens projet/subventions
        SubventionProjet.query.filter_by(projet_id=projet.id).delete(synchronize_session=False)

        # 6) Fichier CR (si présent)
        try:
            if projet.cr_filename:
                folder = os.path.join(current_app.instance_path, "cr_projets")
                fpath = os.path.join(folder, projet.cr_filename)
                if os.path.exists(fpath):
                    os.remove(fpath)
        except Exception:
            pass

        db.session.delete(projet)
        db.session.commit()
        flash("Projet supprimé définitivement.", "warning")
    except Exception as e:
        db.session.rollback()
        flash(f"Suppression impossible : {e}", "danger")

    return redirect(url_for("projets.projets_list"))


@bp.route("/projets/cr/<int:projet_id>/download")
@login_required
@require_perm("projets:files")
def projets_cr_download(projet_id):
    p = Projet.query.get_or_404(projet_id)
    if not can_see_secteur(p.secteur):
        abort(403)

    if not p.cr_filename:
        abort(404)

    folder = ensure_projets_folder()
    return send_from_directory(folder, p.cr_filename, as_attachment=True, download_name=(p.cr_original_name or p.cr_filename))

# ---------------------------------------------------------------------
# Budget AAP par projet : Charges / Produits / Ventilation / Synthèse
# ---------------------------------------------------------------------

@bp.route("/projets/<int:projet_id>/budget")
@login_required
@require_perm("aap:view")
def projet_budget_home(projet_id):
    projet = Projet.query.get_or_404(projet_id)
    if not can_access_secteur(projet.secteur):
        abort(403)
    return redirect(url_for("projets.projet_budget_charges", projet_id=projet_id))


@bp.route("/projets/<int:projet_id>/budget/charges", methods=["GET", "POST"])
@login_required
@require_perm("aap:charges_view")
def projet_budget_charges(projet_id):
    projet = Projet.query.get_or_404(projet_id)
    if not can_access_secteur(projet.secteur):
        abort(403)

    if request.method == "POST":
        if not can("aap:charges_edit"):
            abort(403)
        libelle = (request.form.get("libelle") or "").strip()
        bloc = (request.form.get("bloc") or "directe").strip()
        code_plan = (request.form.get("code_plan") or "60").strip()
        montant = float(request.form.get("montant_previsionnel") or 0)

        if not libelle:
            flash("Libellé obligatoire.", "warning")
        else:
            c = ChargeProjet(
                projet_id=projet.id,
                libelle=libelle,
                bloc=bloc,
                code_plan=code_plan,
                montant_previsionnel=montant,
                montant_reel=float(request.form.get("montant_reel") or 0),
                commentaire=(request.form.get("commentaire") or "").strip() or None,
            )
            db.session.add(c)
            db.session.commit()
            flash("Charge ajoutée.", "success")
            return redirect(url_for("projets.projet_budget_charges", projet_id=projet.id))

    charges = ChargeProjet.query.filter_by(projet_id=projet.id).order_by(ChargeProjet.bloc.asc(), ChargeProjet.code_plan.asc(), ChargeProjet.id.asc()).all()
    return render_template(
        "projets_budget_charges.html",
        projet=projet,
        charges=charges,
        budget_stats=_budget_stats(projet.id),
        active_tab="charges",
    )


@bp.route("/projets/<int:projet_id>/budget/charges/<int:charge_id>/edit", methods=["GET", "POST"])
@login_required
@require_perm("aap:charges_edit")
def projet_budget_charge_edit(projet_id, charge_id):
    projet = Projet.query.get_or_404(projet_id)
    if not can_access_secteur(projet.secteur):
        abort(403)
    charge = ChargeProjet.query.filter_by(id=charge_id, projet_id=projet.id).first_or_404()

    if request.method == "POST":
        charge.libelle = (request.form.get("libelle") or "").strip()
        charge.bloc = (request.form.get("bloc") or "directe").strip()
        charge.code_plan = (request.form.get("code_plan") or "60").strip()
        charge.montant_previsionnel = float(request.form.get("montant_previsionnel") or 0)
        charge.montant_reel = float(request.form.get("montant_reel") or 0)
        charge.commentaire = (request.form.get("commentaire") or "").strip() or None
        db.session.commit()
        flash("Charge mise à jour.", "success")
        return redirect(url_for("projets.projet_budget_charges", projet_id=projet.id))

    return render_template(
        "projets_budget_charge_edit.html",
        projet=projet,
        charge=charge,
        budget_stats=_budget_stats(projet.id),
        active_tab="charges",
    )


@bp.route("/projets/<int:projet_id>/budget/charges/<int:charge_id>/delete", methods=["POST"])
@login_required
@require_perm("aap:charges_edit")
def projet_budget_charge_delete(projet_id, charge_id):
    projet = Projet.query.get_or_404(projet_id)
    if not can_access_secteur(projet.secteur):
        abort(403)
    charge = ChargeProjet.query.filter_by(id=charge_id, projet_id=projet.id).first_or_404()
    db.session.delete(charge)
    db.session.commit()
    flash("Charge supprimée.", "success")
    return redirect(url_for("projets.projet_budget_charges", projet_id=projet.id))


@bp.route("/projets/<int:projet_id>/budget/produits", methods=["GET", "POST"])
@login_required
@require_perm("aap:produits_view")
def projet_budget_produits(projet_id):
    projet = Projet.query.get_or_404(projet_id)
    if not can_access_secteur(projet.secteur):
        abort(403)

    if request.method == "POST":
        if not can("aap:produits_edit"):
            abort(403)
        financeur = (request.form.get("financeur") or "").strip()
        categorie = (request.form.get("categorie") or "autre").strip()
        statut = (request.form.get("statut") or "prevu").strip()
        demande = float(request.form.get("montant_demande") or 0)
        accorde = float(request.form.get("montant_accorde") or 0)
        recu = float(request.form.get("montant_recu") or 0)

        if not financeur:
            flash("Financeur obligatoire.", "warning")
        else:
            p = ProduitProjet(
                projet_id=projet.id,
                financeur=financeur,
                categorie=categorie,
                statut=statut,
                montant_demande=demande,
                montant_accorde=accorde,
                montant_recu=recu,
                reference_dossier=(request.form.get("reference_dossier") or "").strip() or None,
                commentaire=(request.form.get("commentaire") or "").strip() or None,
            )
            db.session.add(p)
            db.session.commit()
            flash("Produit/financeur ajouté.", "success")
            return redirect(url_for("projets.projet_budget_produits", projet_id=projet.id))

    produits = ProduitProjet.query.filter_by(projet_id=projet.id).order_by(ProduitProjet.categorie.asc(), ProduitProjet.financeur.asc()).all()
    return render_template(
        "projets_budget_produits.html",
        projet=projet,
        produits=produits,
        budget_stats=_budget_stats(projet.id),
        active_tab="produits",
    )


@bp.route("/projets/<int:projet_id>/budget/produits/<int:produit_id>/edit", methods=["GET", "POST"])
@login_required
@require_perm("aap:produits_edit")
def projet_budget_produit_edit(projet_id, produit_id):
    projet = Projet.query.get_or_404(projet_id)
    if not can_access_secteur(projet.secteur):
        abort(403)
    produit = ProduitProjet.query.filter_by(id=produit_id, projet_id=projet.id).first_or_404()

    if request.method == "POST":
        produit.financeur = (request.form.get("financeur") or "").strip()
        produit.categorie = (request.form.get("categorie") or "autre").strip()
        produit.statut = (request.form.get("statut") or "prevu").strip()
        produit.montant_demande = float(request.form.get("montant_demande") or 0)
        produit.montant_accorde = float(request.form.get("montant_accorde") or 0)
        produit.montant_recu = float(request.form.get("montant_recu") or 0)
        produit.reference_dossier = (request.form.get("reference_dossier") or "").strip() or None
        produit.commentaire = (request.form.get("commentaire") or "").strip() or None
        db.session.commit()
        flash("Produit/financeur mis à jour.", "success")
        return redirect(url_for("projets.projet_budget_produits", projet_id=projet.id))

    return render_template(
        "projets_budget_produit_edit.html",
        projet=projet,
        produit=produit,
        budget_stats=_budget_stats(projet.id),
        active_tab="produits",
    )


@bp.route("/projets/<int:projet_id>/budget/produits/<int:produit_id>/delete", methods=["POST"])
@login_required
@require_perm("aap:produits_edit")
def projet_budget_produit_delete(projet_id, produit_id):
    projet = Projet.query.get_or_404(projet_id)
    if not can_access_secteur(projet.secteur):
        abort(403)
    produit = ProduitProjet.query.filter_by(id=produit_id, projet_id=projet.id).first_or_404()
    db.session.delete(produit)
    db.session.commit()
    flash("Produit/financeur supprimé.", "success")
    return redirect(url_for("projets.projet_budget_produits", projet_id=projet.id))


@bp.route("/projets/<int:projet_id>/budget/ventilation", methods=["GET", "POST"])
@login_required
@require_perm("aap:ventilation_view")
def projet_budget_ventilation(projet_id):
    projet = Projet.query.get_or_404(projet_id)
    if not can_access_secteur(projet.secteur):
        abort(403)

    charges = ChargeProjet.query.filter_by(projet_id=projet.id).order_by(ChargeProjet.bloc.asc(), ChargeProjet.code_plan.asc(), ChargeProjet.id.asc()).all()
    produits = ProduitProjet.query.filter_by(projet_id=projet.id).order_by(ProduitProjet.categorie.asc(), ProduitProjet.financeur.asc()).all()

    # index existing ventilations
    existing = VentilationProjet.query.join(ChargeProjet).filter(ChargeProjet.projet_id == projet.id).all()
    vmap = {(v.charge_id, v.produit_id): v for v in existing}

    if request.method == "POST":
        if not can("aap:ventilation_edit"):
            abort(403)
        # matrice : v_<charge>_<produit> = montant
        # UX/sécurité : on refuse les ventilations incohérentes (somme > charge, somme > financeur)

        # 1) Lire les valeurs du formulaire en mémoire
        new_vals: dict[tuple[int, int], float] = {}
        for c in charges:
            for p in produits:
                key = f"v_{c.id}_{p.id}"
                if key not in request.form:
                    continue
                raw = (request.form.get(key) or "").strip().replace(",", ".")
                try:
                    val = float(raw) if raw else 0.0
                except ValueError:
                    val = 0.0
                if val < 0:
                    val = 0.0
                new_vals[(c.id, p.id)] = val

        # 2) Contrôles par charge
        errors = []
        for c in charges:
            s = sum((new_vals.get((c.id, p.id), float(vmap.get((c.id, p.id)).montant_ventile or 0)) if vmap.get((c.id, p.id)) else 0.0) for p in produits)
            max_c = float(c.montant_previsionnel or 0)
            # tolérance 1 centime
            if max_c > 0 and s - max_c > 0.01:
                errors.append(f"Charge '{c.libelle}' : {s:.2f}€ ventilés pour {max_c:.2f}€ prévus")

        # 3) Contrôles par produit/financeur
        for p in produits:
            s = sum((new_vals.get((c.id, p.id), float(vmap.get((c.id, p.id)).montant_ventile or 0)) if vmap.get((c.id, p.id)) else 0.0) for c in charges)
            # base de comparaison : reçu > accordé > demandé
            base = float(p.montant_recu or 0) if float(p.montant_recu or 0) > 0 else (float(p.montant_accorde or 0) if float(p.montant_accorde or 0) > 0 else float(p.montant_demande or 0))
            if base > 0 and s - base > 0.01:
                errors.append(f"Financeur '{p.financeur}' : {s:.2f}€ ventilés pour {base:.2f}€ ({'reçu' if float(p.montant_recu or 0) > 0 else ('accordé' if float(p.montant_accorde or 0) > 0 else 'demandé')})")

        if errors:
            flash("Ventilation refusée : incohérence détectée.", "danger")
            for e in errors[:8]:
                flash("• " + e, "warning")
            if len(errors) > 8:
                flash(f"(+{len(errors) - 8} autres incohérences)", "warning")
            return redirect(url_for("projets.projet_budget_ventilation", projet_id=projet.id))

        # 4) Appliquer en base (diff minimal)
        changed = 0
        for c in charges:
            for p in produits:
                val = new_vals.get((c.id, p.id), None)
                if val is None:
                    continue
                cur = vmap.get((c.id, p.id))

                if val <= 0:
                    if cur:
                        db.session.delete(cur)
                        changed += 1
                    continue

                if cur:
                    if float(cur.montant_ventile or 0) != val:
                        cur.montant_ventile = val
                        changed += 1
                else:
                    nv = VentilationProjet(charge_id=c.id, produit_id=p.id, montant_ventile=val)
                    db.session.add(nv)
                    vmap[(c.id, p.id)] = nv
                    changed += 1

        db.session.commit()
        flash(f"Ventilation enregistrée ({changed} modif).", "success")
        return redirect(url_for("projets.projet_budget_ventilation", projet_id=projet.id))

    # rebuild vmap with numbers for template
    vvals = {(cid, pid): float(v.montant_ventile or 0) for (cid, pid), v in vmap.items()}
    return render_template(
        "projets_budget_ventilation.html",
        projet=projet,
        charges=charges,
        produits=produits,
        vvals=vvals,
        budget_stats=_budget_stats(projet.id),
        active_tab="ventilation",
    )


@bp.route("/projets/<int:projet_id>/budget/synthese")
@login_required
@require_perm("aap:synthese_view")
def projet_budget_synthese(projet_id):
    projet = Projet.query.get_or_404(projet_id)
    if not can_access_secteur(projet.secteur):
        abort(403)
    charges = ChargeProjet.query.filter_by(projet_id=projet.id).all()
    produits = ProduitProjet.query.filter_by(projet_id=projet.id).all()

    alertes = []
    # charges non financées
    for c in charges:
        if float(c.reste_a_financer or 0) > 0.01:
            alertes.append(f"Charge non financée : {c.libelle} (reste {c.reste_a_financer:.2f}€)")
    # produits non ventilés
    for p in produits:
        if float(p.reste_a_ventiler or 0) > 0.01:
            alertes.append(f"Produit non ventilé : {p.financeur} (reste {p.reste_a_ventiler:.2f}€)")

    return render_template(
        "projets_budget_synthese.html",
        projet=projet,
        charges=charges,
        produits=produits,
        alertes=alertes,
        budget_stats=_budget_stats(projet.id),
        active_tab="synthese",
    )
