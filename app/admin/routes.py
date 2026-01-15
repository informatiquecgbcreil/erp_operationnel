from flask import Blueprint, render_template, request, redirect, url_for, flash, abort, current_app, jsonify
from flask_login import login_required, current_user

from app.extensions import db
from app.models import User, Role, Permission
from app.rbac import require_perm

bp = Blueprint("admin", __name__, url_prefix="/admin")



def _get_single_role_code_from_form() -> str | None:
    """Accept both old multi-select ('role_codes') and new single-select ('role_code' or 'roles')."""
    # Old UI: <select multiple name="role_codes">
    role_codes = request.form.getlist("role_codes")
    if role_codes:
        # take the first one deterministically
        return (role_codes[0] or "").strip() or None

    # New UI possibilities
    for key in ("role_code", "roles", "role"):
        v = (request.form.get(key) or "").strip()
        if v:
            return v
    return None


@bp.route("/users", methods=["GET", "POST"])
@login_required
@require_perm("admin:users")
def users():
    from app.models import User, Role
    from app.secteurs import get_secteur_labels
    from app.extensions import db

    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        nom = request.form.get("nom", "").strip()
        password = request.form.get("password", "")
        role_code = request.form.get("role")
        secteur = request.form.get("secteur_assigne") or None

        if not email or not password or not role_code:
            flash("Champs obligatoires manquants.", "danger")
            return redirect(url_for("admin.users"))

        if User.query.filter_by(email=email).first():
            flash("Un utilisateur avec cet email existe déjà.", "danger")
            return redirect(url_for("admin.users"))

        u = User(email=email, nom=nom or "Utilisateur")
        u.set_password(password)
        u.secteur_assigne = secteur

        role = Role.query.filter_by(code=role_code).first()
        if role:
            u.roles.append(role)

        db.session.add(u)
        db.session.commit()

        flash("Utilisateur créé.", "success")
        return redirect(url_for("admin.users"))

    users = User.query.order_by(User.nom).all()
    roles = Role.query.order_by(Role.code).all()
    secteurs = get_secteur_labels(active_only=True)

    return render_template(
        "admin_users.html",
        users=users,
        roles=roles,
        secteurs=secteurs,
    )


@bp.route("/delete/<int:user_id>", methods=["POST"])
@login_required
@require_perm("admin:users")
def delete_user(user_id):
    if current_user.id == user_id:
        flash("Tu peux pas te supprimer toi-même.", "danger")
        return redirect(url_for("admin.users"))

    u = User.query.get_or_404(user_id)
    db.session.delete(u)
    db.session.commit()

    flash("Utilisateur supprimé.", "warning")
    return redirect(url_for("admin.users"))


@bp.route("/droits", methods=["GET", "POST"])
@login_required
@require_perm("admin:rbac")
@require_perm("admin:rbac")
def droits():
    """UI RBAC: attribuer des rôles à un user, et éditer les perms d'un rôle."""

    roles = Role.query.order_by(Role.code.asc()).all()
    perms = Permission.query.order_by(Permission.category.asc(), Permission.code.asc()).all()
    users_list = User.query.order_by(User.nom.asc()).all()

    if request.method == "POST":
        action = request.form.get("action")

        # --- Affecter un rôle (unique) à un user ---
        if action == "set_user_roles":
            user_id = int(request.form.get("user_id") or 0)
            u = User.query.get_or_404(user_id)

            role_code = _get_single_role_code_from_form()

            # Force: 1 seul rôle RBAC
            u.roles = []
            if role_code:
                r = Role.query.filter_by(code=role_code).first()
                if r:
                    u.roles.append(r)
                    # legacy sync

            db.session.commit()
            flash("Rôles utilisateur mis à jour.", "success")
            return redirect(url_for("admin.droits"))

        # --- Affecter des permissions à un rôle ---
        if action == "set_role_perms":
            role_code = (request.form.get("role_code") or "").strip()
            perm_codes = set(request.form.getlist("perm_codes"))

            role = Role.query.filter_by(code=role_code).first_or_404()
            role.permissions = []
            for pcode in perm_codes:
                p = Permission.query.filter_by(code=pcode).first()
                if p:
                    role.permissions.append(p)

            db.session.commit()
            flash("Permissions du rôle mises à jour.", "success")
            return redirect(url_for("admin.droits"))

    return render_template(
        "admin_droits.html",
        roles=roles,
        perms=perms,
        users_list=users_list,
        users=users_list,
    )


# ---------------------------------------------------------------------------
# RBAC: endpoints attendus par le template admin_droits.html
# ---------------------------------------------------------------------------

@bp.route("/set_user_roles", methods=["POST"])
@login_required
@require_perm("admin:users")
@require_perm("admin:rbac")
def set_user_roles():
    user_id = int(request.form.get("user_id") or 0)
    u = User.query.get_or_404(user_id)

    role_code = _get_single_role_code_from_form()

    # Force: 1 seul rôle RBAC
    u.roles = []
    if role_code:
        r = Role.query.filter_by(code=role_code).first()
        if r:
            u.roles.append(r)
            # legacy sync

    db.session.commit()
    flash("Rôles utilisateur mis à jour.", "success")
    return redirect(url_for("admin.droits"))


@bp.route("/save_role_perms", methods=["POST"])
@login_required
@require_perm("admin:rbac")
@require_perm("admin:rbac")
def save_role_perms():
    role_code = (request.form.get("role_code") or "").strip()
    perm_codes = set(request.form.getlist("perm_codes"))

    role = Role.query.filter_by(code=role_code).first_or_404()
    role.permissions = []
    for pcode in perm_codes:
        p = Permission.query.filter_by(code=pcode).first()
        if p:
            role.permissions.append(p)

    db.session.commit()
    flash("Permissions du rôle mises à jour.", "success")
    return redirect(url_for("admin.droits"))


@bp.route("/create_role", methods=["POST"])
@login_required
@require_perm("admin:rbac")
@require_perm("admin:rbac")
def create_role():
    code = (request.form.get("code") or "").strip()
    label = (request.form.get("label") or "").strip() or None

    if not code:
        flash("Code de rôle obligatoire.", "danger")
        return redirect(url_for("admin.droits"))

    if Role.query.filter_by(code=code).first():
        flash("Ce rôle existe déjà.", "warning")
        return redirect(url_for("admin.droits"))

    r = Role(code=code, label=label)
    db.session.add(r)
    db.session.commit()

    flash("Rôle créé.", "success")
    return redirect(url_for("admin.droits"))


@bp.route("/delete_role", methods=["POST"])
@login_required
@require_perm("admin:rbac")
@require_perm("admin:rbac")
def delete_role():
    role_code = (request.form.get("role_code") or "").strip()
    if not role_code:
        flash("Choisis un rôle à supprimer.", "danger")
        return redirect(url_for("admin.droits"))

    r = Role.query.filter_by(code=role_code).first()
    if not r:
        flash("Rôle introuvable.", "warning")
        return redirect(url_for("admin.droits"))

    # Détache users + perms (évite erreurs tables d'association)
    for u in User.query.all():
        if hasattr(u, "roles") and r in u.roles:
            u.roles.remove(r)
    r.permissions = []

    db.session.delete(r)
    db.session.commit()

    flash("Rôle supprimé.", "warning")
    return redirect(url_for("admin.droits"))


@bp.route("/get_role_perms/<role_code>", methods=["GET"])
@login_required
@require_perm("admin:rbac")
@require_perm("admin:rbac")
def get_role_perms(role_code):
    r = Role.query.filter_by(code=role_code).first()
    if not r:
        return jsonify({"role": role_code, "perms": []})

    perms = [p.code for p in (r.permissions or [])]
    return jsonify({"role": r.code, "perms": perms})


@bp.route('/debug_rbac', methods=['GET'])
@login_required
@require_perm("admin:rbac")
@require_perm('admin:rbac')
def debug_rbac():
    # Liste des permissions effectives pour l'utilisateur courant (debug)
    perms=set()
    for r in getattr(current_user,'roles',[]) or []:
        for p in getattr(r,'permissions',[]) or []:
            perms.add(p.code)
    return render_template('admin_debug_rbac.html', perms=sorted(perms))
