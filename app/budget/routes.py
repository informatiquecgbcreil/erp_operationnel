import os
from datetime import datetime
from flask import Blueprint, render_template, request, redirect, url_for, flash, abort, current_app, send_from_directory
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename

from app.extensions import db
from app.models import Subvention, LigneBudget, Depense, DepenseDocument
from app.rbac import require_perm, can_access_secteur

bp = Blueprint("budget", __name__)

ALLOWED_EXT = {"pdf", "png", "jpg", "jpeg", "webp", "doc", "docx", "xls", "xlsx"}

def can_see_secteur(secteur: str) -> bool:
    return can_access_secteur(secteur)

def depense_visible(dep: Depense) -> bool:
    l = dep.budget_source
    s = l.source_sub
    return can_see_secteur(s.secteur)

def allowed_file(filename: str) -> bool:
    if not filename or "." not in filename:
        return False
    ext = filename.rsplit(".", 1)[1].lower()
    return ext in ALLOWED_EXT

def ensure_justifs_folder():
    folder = os.path.join(current_app.root_path, "..", "static", "uploads", "justifs")
    folder = os.path.abspath(folder)
    os.makedirs(folder, exist_ok=True)
    return folder

@bp.route("/depense/nouvelle", methods=["GET", "POST"])
@login_required
@require_perm("depenses:create")
def depense_new():
    subs_q = Subvention.query.filter_by(est_archive=False)
    if not current_user.has_perm("scope:all_secteurs"):
        subs_q = subs_q.filter(Subvention.secteur == current_user.secteur_assigne)
    subs = subs_q.order_by(Subvention.annee_exercice.desc(), Subvention.nom.asc()).all()

    if request.method == "POST":
        sub_id = int(request.form.get("subvention_id") or 0)
        compte = (request.form.get("compte") or "").strip()
        ligne_id = int(request.form.get("ligne_budget_id") or 0)

        sub = Subvention.query.get_or_404(sub_id)
        if not can_see_secteur(sub.secteur):
            abort(403)

        ligne = LigneBudget.query.get_or_404(ligne_id)
        if ligne.subvention_id != sub.id:
            abort(400)

        if getattr(ligne, "nature", "charge") != "charge":
            flash("Impossible : une dépense doit être rattachée à une ligne de CHARGE (pas à un produit).", "danger")
            return redirect(url_for("budget.depense_new"))


        libelle = (request.form.get("libelle") or "").strip()
        montant = float(request.form.get("montant") or 0)
        type_depense = (request.form.get("type_depense") or "Fonctionnement").strip()

        date_str = (request.form.get("date_paiement") or "").strip()
        date_p = datetime.strptime(date_str, "%Y-%m-%d").date() if date_str else None

        if not libelle:
            flash("Libellé obligatoire.", "danger")
            return redirect(url_for("budget.depense_new"))

        dep = Depense(
            ligne_budget_id=ligne.id,
            libelle=libelle,
            montant=montant,
            date_paiement=date_p,
            type_depense=type_depense,
            fournisseur=(request.form.get("fournisseur") or "").strip() or None,
            # Depense model uses reference_piece (not 'reference')
            reference_piece=(request.form.get("reference") or "").strip() or None,
            mode_paiement=(request.form.get("mode_paiement") or "").strip() or None,
        )
        db.session.add(dep)
        db.session.commit()

        # Option ergonomique : créer en même temps une entrée inventaire (si demandé)
        if (request.form.get("create_inventory") or "").strip() == "1":
            try:
                from app.models import InventaireItem
                from app.inventaire_materiel.routes import _next_id_interne

                date_ref = date_p or datetime.utcnow().date()
                id_interne = _next_id_interne(sub.secteur, date_ref)

                inv = InventaireItem(
                    secteur=sub.secteur,
                    id_interne=id_interne,
                    categorie=(request.form.get("inv_categorie") or "").strip() or None,
                    designation=(request.form.get("inv_designation") or libelle).strip() or libelle,
                    quantite=int(request.form.get("inv_quantite") or 1),
                    valeur_unitaire=(float(request.form.get("inv_valeur_unitaire")) if (request.form.get("inv_valeur_unitaire") or "").strip() else None),
                    date_entree=date_ref,
                    localisation=(request.form.get("inv_localisation") or "").strip() or None,
                    etat=(request.form.get("inv_etat") or "OK").strip() or "OK",
                    numero_serie=(request.form.get("inv_numero_serie") or "").strip() or None,
                    notes=(request.form.get("inv_notes") or "").strip() or None,
                    depense_id=dep.id,
                    created_by=getattr(current_user, "id", None),
                )
                db.session.add(inv)
                db.session.commit()
                flash("Entrée inventaire créée et liée à la dépense.", "ok")
            except Exception as e:
                db.session.rollback()
                flash(f"Dépense enregistrée, mais inventaire non créé ({e}).", "warning")

        flash("Dépense enregistrée. (Justificatif conseillé 👀)", "warning")
        return redirect(url_for("budget.depense_edit", depense_id=dep.id))

    return render_template("depense_new.html", subs=subs)

@bp.route("/depense/<int:depense_id>/edit", methods=["GET", "POST"])
@login_required
@require_perm("depenses:create")
def depense_edit(depense_id):
    dep = Depense.query.get_or_404(depense_id)
    if not depense_visible(dep):
        abort(403)

    ligne = dep.budget_source
    sub = ligne.source_sub

    if request.method == "POST":
        action = request.form.get("action") or ""

        if action == "update":
            dep.libelle = (request.form.get("libelle") or "").strip()
            dep.montant = float(request.form.get("montant") or 0)
            dep.type_depense = (request.form.get("type_depense") or "Fonctionnement").strip()

            date_str = (request.form.get("date_paiement") or "").strip()
            dep.date_paiement = datetime.strptime(date_str, "%Y-%m-%d").date() if date_str else None

            db.session.commit()
            flash("Dépense modifiée.", "success")
            return redirect(url_for("budget.depense_edit", depense_id=dep.id))

        if action == "upload_doc":
            file = request.files.get("document")
            if not file or not file.filename:
                flash("Aucun fichier.", "danger")
                return redirect(url_for("budget.depense_edit", depense_id=dep.id))

            if not allowed_file(file.filename):
                flash("Type non autorisé (pdf, images, office).", "danger")
                return redirect(url_for("budget.depense_edit", depense_id=dep.id))

            folder = ensure_justifs_folder()
            safe_original = secure_filename(file.filename)

            ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
            stored = secure_filename(f"D{dep.id}_{ts}_{safe_original}")
            file.save(os.path.join(folder, stored))

            doc = DepenseDocument(depense_id=dep.id, filename=stored, original_name=safe_original)
            db.session.add(doc)
            db.session.commit()

            flash("Justificatif ajouté.", "success")
            return redirect(url_for("budget.depense_edit", depense_id=dep.id))

        abort(400)

    alloue = float(ligne.montant_reel or 0)
    engage = float(ligne.engage or 0)
    reste = float(ligne.reste or 0)
    existing_inv = list(getattr(dep, "inventaire_items", []) or [])

    return render_template(
        "depense_edit.html",
        dep=dep,
        ligne=ligne,
        sub=sub,
        alloue=alloue,
        engage=engage,
        reste=reste,
        inventaire_items=existing_inv,
        inventaire_secteur=sub.secteur,
    )

# ✅ NOUVEL ENDPOINT : suppression dépense (fiable)
@bp.route("/depense/<int:depense_id>/delete", methods=["POST"])
@login_required
@require_perm("depenses:delete")
def depense_delete(depense_id):
    dep = Depense.query.get_or_404(depense_id)
    if not depense_visible(dep):
        abort(403)

    # pour retour vers le pilotage
    ligne = dep.budget_source
    sub = ligne.source_sub

    # supprimer les docs + fichiers
    folder = ensure_justifs_folder()
    for doc in dep.documents:
        try:
            path = os.path.join(folder, doc.filename)
            if os.path.exists(path):
                os.remove(path)
        except:
            pass

    db.session.delete(dep)
    db.session.commit()

    flash("Dépense supprimée.", "warning")
    return redirect(url_for("main.subvention_pilotage", subvention_id=sub.id))

@bp.route("/depense/doc/<int:doc_id>/download")
@login_required
@require_perm("depenses:view")
def depense_doc_download(doc_id):
    doc = DepenseDocument.query.get_or_404(doc_id)
    dep = doc.depense

    if not depense_visible(dep):
        abort(403)

    folder = ensure_justifs_folder()
    return send_from_directory(folder, doc.filename, as_attachment=True, download_name=doc.original_name)

@bp.route("/depense/doc/<int:doc_id>/delete", methods=["POST"])
@login_required
@require_perm("depenses:delete")
def depense_doc_delete(doc_id):
    doc = DepenseDocument.query.get_or_404(doc_id)
    dep = doc.depense

    if not depense_visible(dep):
        abort(403)

    folder = ensure_justifs_folder()
    path = os.path.join(folder, doc.filename)

    db.session.delete(doc)
    db.session.commit()

    try:
        if os.path.exists(path):
            os.remove(path)
    except:
        pass

    flash("Justificatif supprimé.", "warning")
    return redirect(url_for("budget.depense_edit", depense_id=dep.id))

@bp.route("/depenses")
@login_required
@require_perm("depenses:view")
def depenses_list():
    # liste des subventions visibles
    subs_q = Subvention.query.filter_by(est_archive=False)
    if not current_user.has_perm("scope:all_secteurs"):
        subs_q = subs_q.filter(Subvention.secteur == current_user.secteur_assigne)
    subs = subs_q.order_by(Subvention.annee_exercice.desc(), Subvention.nom.asc()).all()

    # filtres
    sub_id = request.args.get("subvention_id", type=int)
    ligne_id = request.args.get("ligne_budget_id", type=int)

    dep_q = Depense.query.join(LigneBudget).join(Subvention)

    if not current_user.has_perm("scope:all_secteurs"):
        dep_q = dep_q.filter(Subvention.secteur == current_user.secteur_assigne)

    if sub_id:
        dep_q = dep_q.filter(LigneBudget.subvention_id == sub_id)

    if ligne_id:
        dep_q = dep_q.filter(Depense.ligne_budget_id == ligne_id)

    deps = dep_q.order_by(Depense.date_paiement.desc().nullslast(), Depense.id.desc()).all()

    # lignes possibles si une subvention est sélectionnée
    lignes = []
    if sub_id:
        lignes = LigneBudget.query.filter_by(subvention_id=sub_id).order_by(LigneBudget.compte.asc(), LigneBudget.libelle.asc()).all()

    return render_template(
        "depenses_list.html",
        subs=subs,
        lignes=lignes,
        selected_sub_id=sub_id,
        selected_ligne_id=ligne_id,
        deps=deps,
    )

