from app.rbac import require_perm
import datetime

from flask import Blueprint, render_template, request, abort
from flask_login import login_required, current_user

from app.bilans.services import (
    compute_alertes,
    compute_depenses_mensuelles,
    compute_depenses_par_secteur,
    compute_kpis,
    compute_bilan_secteur,
    compute_bilan_subvention,
    compute_qualite_gestion,
    compute_stats_inventaire,
    compute_bilans_lourds,
    list_secteurs,
    list_subventions,
    list_exercice_years,
    scope_for_user,
)


bp = Blueprint("bilans", __name__, url_prefix="")


@bp.route("/bilans")
@login_required
@require_perm("bilans:view")
def dashboard():
    scope = scope_for_user(current_user)

    years = list_exercice_years(scope)
    # année sélectionnée : param ?year=YYYY, sinon la plus récente dispo
    year_param = request.args.get("year")
    try:
        year = int(year_param) if year_param else years[0]
    except (TypeError, ValueError):
        year = years[0]
    if year not in years:
        # évite de forcer un year arbitraire via l'URL
        abort(403)

    kpis = compute_kpis(year, scope)
    series = compute_depenses_mensuelles(year, scope)
    par_secteur = compute_depenses_par_secteur(year, scope)
    alertes = compute_alertes(year, scope)

    multi_secteurs = scope.secteurs is None

    return render_template(
        "bilans_dashboard.html",
        year=year,
        years=years,
        kpis=kpis,
        series=series,
        par_secteur=par_secteur,
        alertes=alertes,
        multi_secteurs=multi_secteurs,
        scope=scope,
    )



@bp.route("/bilans/lourds")
@login_required
@require_perm("bilans:view")
def bilans_lourds():
    year = request.args.get("year", type=int) or datetime.date.today().year
    scope = scope_for_user(current_user)
    stats = compute_bilans_lourds(year, scope)
    return render_template("bilans_lourds.html", year=year, stats=stats)


@bp.route("/bilans/secteur")
@login_required
@require_perm("bilans:view")
def bilan_secteur():
    scope = scope_for_user(current_user)
    years = list_exercice_years(scope)
    year_param = request.args.get("year")
    try:
        year = int(year_param) if year_param else years[0]
    except (TypeError, ValueError):
        year = years[0]
    if year not in years:
        abort(403)

    secteurs = list_secteurs(year, scope)

    # Choix secteur (finance/direction) via query param; responsable_secteur = auto
    selected = request.args.get("secteur")
    if scope.secteurs is not None:
        # responsable_secteur
        selected = scope.secteurs[0] if scope.secteurs else None
    if selected and selected not in secteurs:
        abort(403)
    if not selected and secteurs:
        selected = secteurs[0]

    data = compute_bilan_secteur(year, selected, scope) if selected else None

    return render_template(
        "bilans_secteur.html",
        year=year,
        years=years,
        secteurs=secteurs,
        selected_secteur=selected,
        data=data,
        scope=scope,
    )


@bp.route("/bilans/subvention")
@login_required
@require_perm("bilans:view")
def bilan_subvention():
    scope = scope_for_user(current_user)
    years = list_exercice_years(scope)
    year_param = request.args.get("year")
    try:
        year = int(year_param) if year_param else years[0]
    except (TypeError, ValueError):
        year = years[0]
    if year not in years:
        abort(403)

    subventions = list_subventions(year, scope)

    # Choix subvention (id) via query param
    selected_id = request.args.get("id")
    selected = None
    if selected_id:
        try:
            sid = int(selected_id)
        except ValueError:
            sid = None
        if sid:
            selected = next((s for s in subventions if s["id"] == sid), None)
            if not selected:
                abort(403)
    if not selected and subventions:
        selected = subventions[0]

    data = compute_bilan_subvention(year, selected["id"], scope) if selected else None

    return render_template(
        "bilans_subvention.html",
        year=year,
        years=years,
        subventions=subventions,
        selected_subvention=selected,
        data=data,
        scope=scope,
    )


@bp.route("/bilans/qualite")
@login_required
@require_perm("bilans:view")
def qualite():
    scope = scope_for_user(current_user)
    years = list_exercice_years(scope)
    year_param = request.args.get("year")
    try:
        year = int(year_param) if year_param else years[0]
    except (TypeError, ValueError):
        year = years[0]
    if year not in years:
        abort(403)

    data = compute_qualite_gestion(year, scope)
    return render_template("bilans_qualite.html", year=year, years=years, data=data, scope=scope)


@bp.route("/bilans/inventaire")
@login_required
@require_perm("bilans:view")
def inventaire():
    scope = scope_for_user(current_user)
    years = list_exercice_years(scope)
    year_param = request.args.get("year")
    try:
        year = int(year_param) if year_param else years[0]
    except (TypeError, ValueError):
        year = years[0]
    if year not in years:
        abort(403)

    data = compute_stats_inventaire(year, scope)
    return render_template("bilans_inventaire.html", year=year, years=years, data=data, scope=scope)
