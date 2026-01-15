import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# Dossier "data" optionnel (utile si tu pack en exe ou si tu veux stocker ailleurs)
DEFAULT_DATA_DIR = os.path.join(os.path.dirname(BASE_DIR), "data")  # ex: si ton exe est dans C:\AppGestion\app\
DATA_DIR = os.environ.get("APP_DATA_DIR", DEFAULT_DATA_DIR)
os.makedirs(DATA_DIR, exist_ok=True)


class Config:
    SECRET_KEY = os.environ.get(
        "SECRET_KEY",
        "Uneapplicationdesuivibudgétairequisimplifielaviedetoutlemondenormalement",
    )

    # --- DB -----------------------------------------------------------------
    # Objectif :
    # - Priorité aux variables d'environnement (Postgres ou autre)
    # - Fallback SQLite local si rien n'est défini
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    INSTANCE_DIR = os.path.join(BASE_DIR, "instance")
    os.makedirs(INSTANCE_DIR, exist_ok=True)

    # Fallback SQLite unique (stable, version Flask standard)
    DB_PATH = os.path.join(INSTANCE_DIR, "database.db")
    _default_sqlite_uri = "sqlite:///" + DB_PATH.replace("\\", "/")

    # Priorité aux variables d'environnement (dans cet ordre)
    _db_url = (
        os.environ.get("SQLALCHEMY_DATABASE_URI")
        or os.environ.get("DATABASE_URL")
        or _default_sqlite_uri
    )

    # Compat anciens formats (Heroku-like)
    if _db_url.startswith("postgres://"):
        _db_url = _db_url.replace("postgres://", "postgresql://", 1)

    SQLALCHEMY_DATABASE_URI = _db_url

    # --- Domaines / constantes ----------------------------------------------
    SECTEURS = [
        "Numérique",
        "Familles",
        "EPE",
        "Santé Transition",
        "Insertion Sociale et Professionnelle",
        "Animation Globale",
    ]

    # SMTP optionnel (envoi des feuilles d'émargement)
    MAIL_HOST = os.environ.get("MAIL_HOST", "")
    MAIL_PORT = int(os.environ.get("MAIL_PORT", "587"))
    MAIL_USERNAME = os.environ.get("MAIL_USERNAME", "")
    MAIL_PASSWORD = os.environ.get("MAIL_PASSWORD", "")
    MAIL_USE_TLS = os.environ.get("MAIL_USE_TLS", "1") in {"1", "true", "True", "yes", "YES"}
    MAIL_SENDER = os.environ.get("MAIL_SENDER", "")

    # URL publique (LAN) de l'application, utilisée pour générer des QR codes.
    # Exemple : http://erp-cgb:8000 ou http://192.168.1.10:8000
    PUBLIC_BASE_URL = os.environ.get("ERP_PUBLIC_BASE_URL", "")
