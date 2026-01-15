import os

from flask import Flask, url_for
from werkzeug.routing import BuildError

from sqlalchemy import text, inspect
from sqlalchemy.exc import OperationalError, ProgrammingError

from config import Config
from app.extensions import db, login_manager, csrf
from app.models import User


def create_app():
    app = Flask(__name__, instance_relative_config=True)
    app.config.from_object(Config)

    # Instance folder (sqlite db, uploads, etc.)
    os.makedirs(app.instance_path, exist_ok=True)

    # Extensions
    db.init_app(app)
    login_manager.init_app(app)
    csrf.init_app(app)
    login_manager.login_view = "auth.login"

    # ------------------------------------------------------------------
    # Jinja helper: safe_url_for
    # ------------------------------------------------------------------
    def safe_url_for(endpoint: str, fallback: str = "#", **values) -> str:
        try:
            return url_for(endpoint, **values)
        except BuildError:
            return fallback

    app.jinja_env.globals["safe_url_for"] = safe_url_for

    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))

    # ------------------------------------------------------------------
    # Blueprints
    # ------------------------------------------------------------------
    from app.auth.routes import bp as auth_bp
    from app.main.routes import bp as main_bp
    from app.budget.routes import bp as budget_bp
    from app.projets.routes import bp as projets_bp
    from app.admin.routes import bp as admin_bp
    from app.activite import bp as activite_bp
    from app.kiosk import bp as kiosk_bp
    from app.statsimpact.routes import bp as statsimpact_bp
    from app.bilans.routes import bp as bilans_bp
    from app.inventaire.routes import bp as inventaire_bp
    from app.inventaire_materiel.routes import bp as inventaire_materiel_bp
    from app.participants.routes import bp as participants_bp
    from app.launcher import bp as launcher_bp
    from app.pedagogie.routes import bp as pedagogie_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(main_bp)
    app.register_blueprint(budget_bp)
    app.register_blueprint(projets_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(activite_bp)
    app.register_blueprint(kiosk_bp)
    app.register_blueprint(statsimpact_bp)
    app.register_blueprint(bilans_bp)
    app.register_blueprint(inventaire_bp)
    app.register_blueprint(inventaire_materiel_bp)
    app.register_blueprint(participants_bp)
    app.register_blueprint(launcher_bp)
    app.register_blueprint(pedagogie_bp)

    # ------------------------------------------------------------------
    # RBAC helpers
    # ------------------------------------------------------------------
    from app.rbac import bootstrap_rbac, can

    @app.context_processor
    def _inject_rbac_helpers():
        return {"can": can}

    # ------------------------------------------------------------------
    # ensure_schema : migrations légères SQLite / Postgres
    # ------------------------------------------------------------------
    def ensure_schema():
        dialect = db.engine.dialect.name
        insp = inspect(db.engine)

        def has_table(name):
            try:
                return insp.has_table(name)
            except Exception:
                return False

        def get_cols(table):
            if not has_table(table):
                return set()
            return {c["name"] for c in insp.get_columns(table)}

        def exec_sql(sql):
            db.session.execute(text(sql))

        def add_col(table, col, sql_sqlite, sql_pg):
            if col in get_cols(table):
                return
            if dialect == "sqlite":
                exec_sql(sql_sqlite)
            else:
                exec_sql(sql_pg)

        # --------------------------------------------------------------
        # 0) LEGACY : colonne user.role (OBLIGATOIRE pour le boot)
        # --------------------------------------------------------------
        try:
            add_col(
                "user",
                "role",
                'ALTER TABLE "user" ADD COLUMN role VARCHAR(50) NOT NULL DEFAULT "responsable_secteur"',
                'ALTER TABLE "user" ADD COLUMN IF NOT EXISTS role VARCHAR(50) NOT NULL DEFAULT \'responsable_secteur\'',
            )
            db.session.commit()
        except Exception:
            db.session.rollback()

        # --------------------------------------------------------------
        # 1) Exemple : colonne nature sur ligne_budget
        # --------------------------------------------------------------
        try:
            add_col(
                "ligne_budget",
                "nature",
                "ALTER TABLE ligne_budget ADD COLUMN nature VARCHAR(10) NOT NULL DEFAULT 'charge'",
                "ALTER TABLE ligne_budget ADD COLUMN IF NOT EXISTS nature VARCHAR(10) NOT NULL DEFAULT 'charge'",
            )
            db.session.commit()
        except Exception:
            db.session.rollback()

    # ------------------------------------------------------------------
    # INIT DB (ORDRE CRUCIAL)
    # ------------------------------------------------------------------
    with app.app_context():
        # 1) Créer les tables
        db.create_all()

        # 2) Garantir le schéma minimal (user.role AVANT RBAC)
        ensure_schema()

        # 3) Bootstrap RBAC (peut maintenant query User sans crash)
        bootstrap_rbac()

        # 4) Bootstrap Secteurs (depuis la config, non-destructif)
        from app.secteurs import bootstrap_secteurs_from_config
        bootstrap_secteurs_from_config()

        print("DB URI =", db.engine.url)
        print("DB DIALECT =", db.engine.dialect.name)

        @app.context_processor
        def inject_secteurs():
            # Secteurs canoniques pour les formulaires.
            # Source: DB (Secteur) avec fallback config.
            from app.secteurs import get_secteur_labels
            return {"SECTEURS": get_secteur_labels(active_only=True)}


        return app
