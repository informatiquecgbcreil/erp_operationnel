from datetime import datetime
from datetime import date
import json
from werkzeug.security import generate_password_hash, check_password_hash
from app.extensions import db

# ---------- USERS ----------
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(180), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(255), nullable=False)
    nom = db.Column(db.String(120), nullable=False, default="Utilisateur")
    role = db.Column(db.String(40), nullable=False, default="responsable_secteur")
    secteur_assigne = db.Column(db.String(80), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    # Flask-Login
    @property
    def is_authenticated(self):
        return True

    @property
    def is_active(self):
        return True

    @property
    def is_anonymous(self):
        return False

    def get_id(self):
        return str(self.id)

    def set_password(self, password: str):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)

    # RBAC helpers (roles/permissions)
    def has_perm(self, code: str) -> bool:
        codes: set[str] = set()
        for role in getattr(self, "roles", []) or []:
            for p in getattr(role, "permissions", []) or []:
                codes.add(p.code)
        return code in codes

    @property
    def role_codes(self) -> list[str]:
        return sorted([r.code for r in getattr(self, "roles", []) or []])


    def has_role(self, code: str | None) -> bool:
        """Compat helper: True si l'utilisateur possède le rôle `code`.

        - D'abord via RBAC (User.roles -> Role.code)
        - Puis fallback legacy (User.role string) pour ne pas exploser les anciennes routes
        - Gère quelques alias historiques (directrice -> direction, financière -> finance, etc.)
        """
        if not code:
            return False

        c = (code or "").strip().lower()

        aliases = {
            # historiques
            "directrice": "direction",
            "directeur": "direction",
            "financiere": "finance",
            "financière": "finance",
            "responsable_secteurs": "responsable_secteur",
        }
        c = aliases.get(c, c)

        # RBAC (relation roles)
        try:
            for r in (getattr(self, "roles", []) or []):
                rc = (getattr(r, "code", "") or "").strip().lower()
                rc = aliases.get(rc, rc)
                if rc == c:
                    return True
        except Exception:
            pass

        # Legacy
        legacy = (getattr(self, "role", None) or "").strip().lower()
        legacy = aliases.get(legacy, legacy)
        return legacy == c


# =========================================================
# RBAC (Roles & Permissions)
# ---------------------------------------------------------
# Objectif: remplacer progressivement la logique "role = string" par
# une vraie gestion fine des permissions.
# Le champ User.role reste (compatibilité), mais on le mappe vers
# un ou plusieurs Roles.

user_roles = db.Table(
    "user_roles",
    db.Column("user_id", db.Integer, db.ForeignKey("user.id", ondelete="CASCADE"), primary_key=True),
    db.Column("role_id", db.Integer, db.ForeignKey("role.id", ondelete="CASCADE"), primary_key=True),
)

role_permissions = db.Table(
    "role_permissions",
    db.Column("role_id", db.Integer, db.ForeignKey("role.id", ondelete="CASCADE"), primary_key=True),
    db.Column("permission_id", db.Integer, db.ForeignKey("permission.id", ondelete="CASCADE"), primary_key=True),
)


class Role(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(60), unique=True, nullable=False, index=True)  # ex: "finance", "admin_tech"
    label = db.Column(db.String(120), nullable=False, default="Rôle")

    permissions = db.relationship(
        "Permission",
        secondary=role_permissions,
        lazy="subquery",
        backref=db.backref("roles", lazy=True),
    )

    def __repr__(self) -> str:
        return f"<Role {self.code}>"


class Permission(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(120), unique=True, nullable=False, index=True)  # ex: "subventions:edit"
    label = db.Column(db.String(200), nullable=False, default="Permission")
    category = db.Column(db.String(60), nullable=True, index=True)  # ex: "Subventions"

    def __repr__(self) -> str:
        return f"<Perm {self.code}>"


class Secteur(db.Model):
    """Secteur métier (liste administrable).

    ⚠️ Pour limiter le refacto, on conserve *label* comme valeur utilisée
    partout (Subvention.secteur, Projet.secteur, etc.).
    Le champ `code` sert surtout d'identifiant stable/slug.
    """

    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(80), unique=True, nullable=False, index=True)
    label = db.Column(db.String(120), unique=True, nullable=False, index=True)
    is_active = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self) -> str:
        return f"<Secteur {self.code} ({'on' if self.is_active else 'off'})>"


# Relation User.roles (déclarée après Role)
User.roles = db.relationship(
    "Role",
    secondary=user_roles,
    lazy="subquery",
    backref=db.backref("users", lazy=True),
)


# ---------- PEDAGOGIE ----------
class Referentiel(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nom = db.Column(db.String(120), nullable=False)
    description = db.Column(db.Text, nullable=True)


class Competence(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    referentiel_id = db.Column(db.Integer, db.ForeignKey("referentiel.id"), nullable=False)
    code = db.Column(db.String(40), nullable=False)
    nom = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text, nullable=True)

    referentiel = db.relationship("Referentiel", backref=db.backref("competences", cascade="all, delete-orphan"))


projet_competence = db.Table(
    "projet_competence",
    db.Column("projet_id", db.Integer, db.ForeignKey("projet.id"), primary_key=True),
    db.Column("competence_id", db.Integer, db.ForeignKey("competence.id"), primary_key=True),
)


atelier_competence = db.Table(
    "atelier_competence",
    db.Column("atelier_id", db.Integer, db.ForeignKey("atelier_activite.id"), primary_key=True),
    db.Column("competence_id", db.Integer, db.ForeignKey("competence.id"), primary_key=True),
)


session_competence = db.Table(
    "session_competence",
    db.Column("session_id", db.Integer, db.ForeignKey("session_activite.id"), primary_key=True),
    db.Column("competence_id", db.Integer, db.ForeignKey("competence.id"), primary_key=True),
)


objectif_competence = db.Table(
    "objectif_competence",
    db.Column("objectif_id", db.Integer, db.ForeignKey("objectif.id"), primary_key=True),
    db.Column("competence_id", db.Integer, db.ForeignKey("competence.id"), primary_key=True),
)


class Objectif(db.Model):
    __tablename__ = "objectif"
    id = db.Column(db.Integer, primary_key=True)
    parent_id = db.Column(db.Integer, db.ForeignKey("objectif.id"), nullable=True)
    type = db.Column(db.String(30), nullable=False)  # general | specifique | operationnel
    titre = db.Column(db.String(255), nullable=False)
    description = db.Column(db.Text, nullable=True)
    seuil_validation = db.Column(db.Float, nullable=False, default=60.0)

    projet_id = db.Column(db.Integer, db.ForeignKey("projet.id"), nullable=True)
    atelier_id = db.Column(db.Integer, db.ForeignKey("atelier_activite.id"), nullable=True)
    session_id = db.Column(db.Integer, db.ForeignKey("session_activite.id"), nullable=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    parent = db.relationship("Objectif", remote_side=[id], backref=db.backref("enfants", cascade="all, delete-orphan"))
    projet = db.relationship("Projet")
    atelier = db.relationship("AtelierActivite")
    session = db.relationship("SessionActivite")
    competences = db.relationship(
        "Competence",
        secondary=objectif_competence,
        backref=db.backref("objectifs", lazy="dynamic"),
    )


# ---------- PROJETS ----------
class Projet(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nom = db.Column(db.String(200), nullable=False)
    secteur = db.Column(db.String(80), nullable=False)
    description = db.Column(db.Text, nullable=True)

    cr_filename = db.Column(db.String(255), nullable=True)
    cr_original_name = db.Column(db.String(255), nullable=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    subventions = db.relationship("SubventionProjet", back_populates="projet", cascade="all, delete-orphan")
    # AAP / Budget projet (charges/produits/ventilations)
    charges_projet = db.relationship("ChargeProjet", back_populates="projet", cascade="all, delete-orphan")
    produits_projet = db.relationship("ProduitProjet", back_populates="projet", cascade="all, delete-orphan")
    competences = db.relationship(
        "Competence",
        secondary=projet_competence,
        backref=db.backref("projets", lazy="dynamic"),
    )

    @property
    def total_demande(self):
        return round(sum(float(sp.subvention.montant_demande or 0) for sp in self.subventions), 2)

    @property
    def total_attribue(self):
        return round(sum(float(sp.subvention.montant_attribue or 0) for sp in self.subventions), 2)

    @property
    def total_recu(self):
        return round(sum(float(sp.subvention.montant_recu or 0) for sp in self.subventions), 2)

    @property
    def total_reel_lignes(self):
        return round(sum(float(sp.subvention.total_reel_lignes or 0) for sp in self.subventions), 2)

    @property
    def total_engage(self):
        return round(sum(float(sp.subvention.total_engage or 0) for sp in self.subventions), 2)

    @property
    def total_reste(self):
        return round(sum(float(sp.subvention.total_reste or 0) for sp in self.subventions), 2)


    # -----------------------------
    # Budget AAP (par projet)
    # -----------------------------
    @property
    def total_charges_previsionnel(self):
        return round(sum(float(c.montant_previsionnel or 0) for c in self.charges_projet), 2)

    @property
    def total_charges_reel(self):
        return round(sum(float(c.montant_reel or 0) for c in self.charges_projet), 2)

    @property
    def total_produits_demandes(self):
        return round(sum(float(p.montant_demande or 0) for p in self.produits_projet), 2)

    @property
    def total_produits_accordes(self):
        return round(sum(float(p.montant_accorde or 0) for p in self.produits_projet), 2)

    @property
    def total_produits_recus(self):
        return round(sum(float(p.montant_recu or 0) for p in self.produits_projet), 2)

    @property
    def reste_a_financer(self):
        # basé sur l'accordé (et non la demande)
        return round(float(self.total_charges_previsionnel or 0) - float(self.total_produits_accordes or 0), 2)




class ChargeProjet(db.Model):
    __tablename__ = "charge_projet"
    id = db.Column(db.Integer, primary_key=True)
    projet_id = db.Column(db.Integer, db.ForeignKey("projet.id"), nullable=False)

    # bloc = directe / indirecte (comme le tableau AAP)
    bloc = db.Column(db.String(20), nullable=False, default="directe")  # directe | indirecte
    # code plan comptable : 60/61/62/63/64/65/...
    code_plan = db.Column(db.String(20), nullable=False, default="60")

    libelle = db.Column(db.String(255), nullable=False)

    montant_previsionnel = db.Column(db.Float, default=0.0)
    montant_reel = db.Column(db.Float, default=0.0)

    commentaire = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    projet = db.relationship("Projet", back_populates="charges_projet")
    ventilations = db.relationship("VentilationProjet", back_populates="charge", cascade="all, delete-orphan")
    depenses = db.relationship("Depense", back_populates="charge_projet", passive_deletes=True)

    @property
    def ventile(self):
        return round(sum(float(v.montant_ventile or 0) for v in self.ventilations), 2)

    @property
    def reste_a_financer(self):
        return round(float(self.montant_previsionnel or 0) - float(self.ventile or 0), 2)

    @property
    def engage(self):
        # engagement réel via les dépenses rattachées à cette charge
        return round(sum(float(d.montant or 0) for d in self.depenses if not d.est_supprimee), 2)

    @property
    def reste_a_engager(self):
        base = float(self.montant_reel or 0) if float(self.montant_reel or 0) > 0 else float(self.montant_previsionnel or 0)
        return round(base - float(self.engage or 0), 2)


class ProduitProjet(db.Model):
    __tablename__ = "produit_projet"
    id = db.Column(db.Integer, primary_key=True)
    projet_id = db.Column(db.Integer, db.ForeignKey("projet.id"), nullable=False)

    financeur = db.Column(db.String(255), nullable=False)
    categorie = db.Column(db.String(50), nullable=False, default="autre")  # etat/region/departement/commune/caf/europe/prive/autofinancement/...
    statut = db.Column(db.String(30), nullable=False, default="prevu")  # prevu/demande/accorde/partiel/refuse

    montant_demande = db.Column(db.Float, default=0.0)
    montant_accorde = db.Column(db.Float, default=0.0)
    montant_recu = db.Column(db.Float, default=0.0)

    reference_dossier = db.Column(db.String(120), nullable=True)
    commentaire = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    projet = db.relationship("Projet", back_populates="produits_projet")
    ventilations = db.relationship("VentilationProjet", back_populates="produit", cascade="all, delete-orphan")

    @property
    def ventile(self):
        return round(sum(float(v.montant_ventile or 0) for v in self.ventilations), 2)

    @property
    def reste_a_ventiler(self):
        return round(float(self.montant_accorde or 0) - float(self.ventile or 0), 2)


class VentilationProjet(db.Model):
    __tablename__ = "ventilation_projet"
    id = db.Column(db.Integer, primary_key=True)
    charge_id = db.Column(db.Integer, db.ForeignKey("charge_projet.id", ondelete="CASCADE"), nullable=False)
    produit_id = db.Column(db.Integer, db.ForeignKey("produit_projet.id", ondelete="CASCADE"), nullable=False)
    montant_ventile = db.Column(db.Float, default=0.0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    charge = db.relationship("ChargeProjet", back_populates="ventilations")
    produit = db.relationship("ProduitProjet", back_populates="ventilations")

class SubventionProjet(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    projet_id = db.Column(db.Integer, db.ForeignKey("projet.id"), nullable=False)
    subvention_id = db.Column(db.Integer, db.ForeignKey("subvention.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    projet = db.relationship("Projet", back_populates="subventions")
    subvention = db.relationship("Subvention", back_populates="projets")

    __table_args__ = (
        db.UniqueConstraint("projet_id", "subvention_id", name="uq_projet_subvention"),
    )



# ---------- LIENS PROJET <-> ATELIERS (activité) ----------
class ProjetAtelier(db.Model):
    __tablename__ = "projet_atelier"
    id = db.Column(db.Integer, primary_key=True)
    projet_id = db.Column(db.Integer, db.ForeignKey("projet.id"), nullable=False, index=True)
    atelier_id = db.Column(db.Integer, db.ForeignKey("atelier_activite.id"), nullable=False, index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    projet = db.relationship("Projet", backref=db.backref("ateliers", cascade="all, delete-orphan"))
    atelier = db.relationship("AtelierActivite")

    __table_args__ = (
        db.UniqueConstraint("projet_id", "atelier_id", name="uq_projet_atelier"),
    )


# ---------- INDICATEURS DE PROJET ----------
class ProjetIndicateur(db.Model):
    __tablename__ = "projet_indicateur"
    id = db.Column(db.Integer, primary_key=True)
    projet_id = db.Column(db.Integer, db.ForeignKey("projet.id"), nullable=False, index=True)

    # template (V1)
    code = db.Column(db.String(60), nullable=False)
    label = db.Column(db.String(200), nullable=False)
    is_active = db.Column(db.Boolean, default=True, nullable=False)
    params_json = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    projet = db.relationship("Projet", backref=db.backref("indicateurs", cascade="all, delete-orphan"))

    __table_args__ = (
        db.UniqueConstraint("projet_id", "code", name="uq_projet_indicateur_code"),
    )

    def params(self):
        try:
            return json.loads(self.params_json or "{}")
        except Exception:
            return {}




# ---------- SUBVENTIONS / BUDGET ----------
class Subvention(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nom = db.Column(db.String(200), nullable=False)
    secteur = db.Column(db.String(80), nullable=False)
    annee_exercice = db.Column(db.Integer, nullable=False, default=2025)

    montant_demande = db.Column(db.Float, default=0.0)
    montant_attribue = db.Column(db.Float, default=0.0)
    montant_recu = db.Column(db.Float, default=0.0)

    est_archive = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    lignes = db.relationship("LigneBudget", backref="source_sub", cascade="all, delete-orphan")
    projets = db.relationship("SubventionProjet", back_populates="subvention", cascade="all, delete-orphan")

    @property
    def total_base_lignes(self):
        # compat: total des CHARGES (lignes nature=charge)
        return round(sum(float(l.montant_base or 0) for l in self.lignes if getattr(l, "nature", "charge") == "charge"), 2)

    @property
    def total_reel_lignes(self):
        # compat: total des CHARGES (lignes nature=charge)
        return round(sum(float(l.montant_reel or 0) for l in self.lignes if getattr(l, "nature", "charge") == "charge"), 2)


    @property
    def total_base_produits(self):
        return round(sum(float(l.montant_base or 0) for l in self.lignes if getattr(l, "nature", "charge") == "produit"), 2)

    @property
    def total_reel_produits(self):
        return round(sum(float(l.montant_reel or 0) for l in self.lignes if getattr(l, "nature", "charge") == "produit"), 2)

    @property
    def solde_base(self):
        # Produits - Charges
        return round(float(self.total_base_produits or 0) - float(self.total_base_lignes or 0), 2)

    @property
    def solde_reel(self):
        # Produits - Charges
        return round(float(self.total_reel_produits or 0) - float(self.total_reel_lignes or 0), 2)
    @property
    def total_engage(self):
        return round(sum(float(l.engage or 0) for l in self.lignes if getattr(l, "nature", "charge") == "charge"), 2)

    @property
    def total_reste(self):
        return round(sum(float(l.reste or 0) for l in self.lignes if getattr(l, "nature", "charge") == "charge"), 2)


class LigneBudget(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    subvention_id = db.Column(db.Integer, db.ForeignKey("subvention.id"), nullable=False)

    # nature = charge (compte 6*) ou produit (compte 7*)
    nature = db.Column(db.String(10), nullable=False, default="charge")  # charge | produit

    compte = db.Column(db.String(20), nullable=False, default="60")
    libelle = db.Column(db.String(200), nullable=False)

    montant_base = db.Column(db.Float, default=0.0)
    montant_reel = db.Column(db.Float, default=0.0)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    depenses = db.relationship("Depense", backref="budget_source", cascade="all, delete-orphan")

    @property
    def engage(self):
        # engage / reste n'ont de sens que pour les CHARGES
        if getattr(self, "nature", "charge") != "charge":
            return 0.0
        # BLINDAGE: on ne compte pas les dépenses soft-delete
        return round(sum(float(d.montant or 0) for d in self.depenses if not d.est_supprimee), 2)

    @property
    def reste(self):
        if getattr(self, "nature", "charge") != "charge":
            return 0.0
        return round(float(self.montant_reel or 0) - float(self.engage or 0), 2)


class Depense(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ligne_budget_id = db.Column(db.Integer, db.ForeignKey("ligne_budget.id"), nullable=True)
    # Nouveau (AAP/Projets) : rattachement direct à une charge projet
    charge_projet_id = db.Column(db.Integer, db.ForeignKey("charge_projet.id", ondelete="SET NULL"), nullable=True)

    # Provenance facture / inventaire
    facture_ligne_id = db.Column(db.Integer, db.ForeignKey("facture_ligne.id", ondelete="SET NULL"), nullable=True)

    libelle = db.Column(db.String(255), nullable=False)
    montant = db.Column(db.Float, default=0.0)

    # infos finance-friendly (non obligatoires pour l’instant)
    fournisseur = db.Column(db.String(180), nullable=True)
    reference_piece = db.Column(db.String(120), nullable=True)  # n° facture / reçu / référence
    mode_paiement = db.Column(db.String(50), nullable=True)     # CB / Virement / Espèces / Autre

    date_paiement = db.Column(db.Date, nullable=True)
    type_depense = db.Column(db.String(80), default="Fonctionnement")

    # workflow / blindage
    statut = db.Column(db.String(30), nullable=False, default="valide")  # brouillon / valide
    anomalie = db.Column(db.String(255), nullable=True)
    est_supprimee = db.Column(db.Boolean, default=False)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    documents = db.relationship("DepenseDocument", backref="depense", cascade="all, delete-orphan")
    inventaire_items = db.relationship("InventaireItem", backref="depense", passive_deletes=True)
    # relation SQLAlchemy (nécessaire pour back_populates depuis ChargeProjet)
    charge_projet = db.relationship("ChargeProjet", back_populates="depenses")


class DepenseDocument(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    depense_id = db.Column(db.Integer, db.ForeignKey("depense.id"), nullable=False)

    filename = db.Column(db.String(255), nullable=False)
    original_name = db.Column(db.String(255), nullable=False)

    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)


# ---------- FACTURES / INVENTAIRE ----------
class FactureAchat(db.Model):
    __tablename__ = "facture_achat"

    id = db.Column(db.Integer, primary_key=True)
    secteur_principal = db.Column(db.String(80), nullable=False)
    fournisseur = db.Column(db.String(180), nullable=True)
    reference_facture = db.Column(db.String(120), nullable=True)
    date_facture = db.Column(db.Date, nullable=True)

    statut = db.Column(db.String(30), nullable=False, default="brouillon")  # brouillon / valide

    filename = db.Column(db.String(255), nullable=True)
    original_name = db.Column(db.String(255), nullable=True)

    created_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    lignes = db.relationship("FactureLigne", backref="facture", cascade="all, delete-orphan")

    @property
    def total(self):
        return round(sum(float(l.montant_ligne or 0) for l in self.lignes), 2)


class FactureLigne(db.Model):
    __tablename__ = "facture_ligne"

    id = db.Column(db.Integer, primary_key=True)
    facture_id = db.Column(db.Integer, db.ForeignKey("facture_achat.id"), nullable=False)
    secteur = db.Column(db.String(80), nullable=False)

    financement_type = db.Column(db.String(30), nullable=False, default="subvention")  # subvention / fonds_propres / don / autre
    a_ventiler = db.Column(db.Boolean, default=False)

    libelle = db.Column(db.String(255), nullable=False)
    quantite = db.Column(db.Integer, nullable=False, default=1)
    prix_unitaire = db.Column(db.Float, default=0.0)
    montant_ligne = db.Column(db.Float, default=0.0)

    ligne_budget_id = db.Column(db.Integer, db.ForeignKey("ligne_budget.id"), nullable=True)
    # Nouveau (AAP/Projets) : rattachement direct à une charge projet
    charge_projet_id = db.Column(db.Integer, db.ForeignKey("charge_projet.id", ondelete="SET NULL"), nullable=True)
    subvention_id = db.Column(db.Integer, db.ForeignKey("subvention.id"), nullable=False)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    depenses = db.relationship("Depense", backref="facture_ligne", passive_deletes=True)
    inventaire_items = db.relationship("InventaireItem", backref="facture_ligne", passive_deletes=True)


class InventaireItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    secteur = db.Column(db.String(80), nullable=False)
    id_interne = db.Column(db.String(64), nullable=False, unique=True)

    categorie = db.Column(db.String(120), nullable=True)
    designation = db.Column(db.String(255), nullable=False)
    marque = db.Column(db.String(120), nullable=True)
    modele = db.Column(db.String(120), nullable=True)

    quantite = db.Column(db.Integer, nullable=False, default=1)
    numero_serie = db.Column(db.String(180), nullable=True)
    etat = db.Column(db.String(50), nullable=False, default="OK")
    localisation = db.Column(db.String(255), nullable=True)

    valeur_unitaire = db.Column(db.Float, nullable=True)
    date_entree = db.Column(db.Date, nullable=True)
    notes = db.Column(db.Text, nullable=True)

    facture_ligne_id = db.Column(db.Integer, db.ForeignKey("facture_ligne.id", ondelete="SET NULL"), nullable=True)
    depense_id = db.Column(db.Integer, db.ForeignKey("depense.id", ondelete="SET NULL"), nullable=True)

    created_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


# ==========================================================
# ===============  ACTIVITÉ / ÉMARGEMENT  ==================
# ==========================================================

class Quartier(db.Model):
    __tablename__ = "quartier"
    id = db.Column(db.Integer, primary_key=True)
    ville = db.Column(db.String(80), nullable=False, default="Creil")
    nom = db.Column(db.String(120), nullable=False)
    is_qpv = db.Column(db.Boolean, default=False)

    __table_args__ = (
        db.UniqueConstraint("ville", "nom", name="uq_quartier_ville_nom"),
    )


class Participant(db.Model):
    __tablename__ = "participant"
    id = db.Column(db.Integer, primary_key=True)

    nom = db.Column(db.String(120), nullable=False)
    prenom = db.Column(db.String(120), nullable=False)
    adresse = db.Column(db.String(255), nullable=True)
    ville = db.Column(db.String(120), nullable=True)
    email = db.Column(db.String(180), nullable=True)
    telephone = db.Column(db.String(60), nullable=True)
    genre = db.Column(db.String(20), nullable=True)
    date_naissance = db.Column(db.Date, nullable=True)

    # Type de public (ex: H/S/B/A/P). Par défaut: H
    type_public = db.Column(db.String(2), nullable=False, default="H")

    quartier_id = db.Column(db.Integer, db.ForeignKey("quartier.id"), nullable=True)
    quartier = db.relationship("Quartier")

    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # Pour permettre la création "en avance" (avant toute présence) tout en respectant
    # le cloisonnement par secteur en rôle responsable_secteur.
    created_by_user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    created_secteur = db.Column(db.String(80), nullable=True)

    @property
    def is_creil(self):
        return (self.ville or "").strip().lower() == "creil"

    @property
    def is_qpv(self):
        return bool(self.quartier and self.quartier.is_qpv)

    @property
    def age(self):
        if not self.date_naissance:
            return None
        today = date.today()
        years = today.year - self.date_naissance.year
        if (today.month, today.day) < (self.date_naissance.month, self.date_naissance.day):
            years -= 1
        return years


class AtelierActivite(db.Model):
    __tablename__ = "atelier_activite"
    id = db.Column(db.Integer, primary_key=True)
    secteur = db.Column(db.String(80), nullable=False, index=True)
    nom = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text, nullable=True)
    type_atelier = db.Column(db.String(30), nullable=False, default="COLLECTIF")
    # COLLECTIF: nb places. INDIVIDUEL_MENSUEL: heures dispo / mois.
    capacite_defaut = db.Column(db.Integer, nullable=True)
    heures_dispo_defaut_mois = db.Column(db.Float, nullable=True)
    duree_defaut_minutes = db.Column(db.Integer, nullable=True)

    motifs_json = db.Column(db.Text, nullable=True)  # liste JSON de motifs (dropdown)

    modele_docx_collectif = db.Column(db.String(255), nullable=True)
    modele_docx_individuel = db.Column(db.String(255), nullable=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    # Soft-delete (safe during tests / RGPD / audit)
    is_deleted = db.Column(db.Boolean, nullable=False, default=False, index=True)
    deleted_at = db.Column(db.DateTime, nullable=True)

    sessions = db.relationship("SessionActivite", backref="atelier", cascade="all, delete-orphan")
    competences = db.relationship(
        "Competence",
        secondary=atelier_competence,
        backref=db.backref("ateliers", lazy="dynamic"),
    )

    def motifs(self):
        try:
            return json.loads(self.motifs_json or "[]")
        except Exception:
            return []


class SessionActivite(db.Model):
    __tablename__ = "session_activite"
    id = db.Column(db.Integer, primary_key=True)
    atelier_id = db.Column(db.Integer, db.ForeignKey("atelier_activite.id"), nullable=False)
    secteur = db.Column(db.String(80), nullable=False, index=True)
    session_type = db.Column(db.String(30), nullable=False, default="COLLECTIF")
    # COLLECTIF
    date_session = db.Column(db.Date, nullable=True, index=True)
    heure_debut = db.Column(db.String(10), nullable=True)
    heure_fin = db.Column(db.String(10), nullable=True)
    capacite = db.Column(db.Integer, nullable=True)
    statut = db.Column(db.String(20), nullable=False, default="realisee")  # realisee / annulee

    # INDIVIDUEL_MENSUEL (rdv)
    rdv_date = db.Column(db.Date, nullable=True, index=True)
    rdv_debut = db.Column(db.String(10), nullable=True)
    rdv_fin = db.Column(db.String(10), nullable=True)
    duree_minutes = db.Column(db.Integer, nullable=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    # Soft-delete (safe during tests)
    is_deleted = db.Column(db.Boolean, nullable=False, default=False, index=True)
    deleted_at = db.Column(db.DateTime, nullable=True)

    # KIOSQUE (public) : émargement via /kiosk sans exposer l'app complète
    kiosk_open = db.Column(db.Boolean, default=False, index=True)
    kiosk_pin = db.Column(db.String(10), nullable=True, index=True)
    kiosk_token = db.Column(db.String(64), nullable=True, index=True)
    kiosk_opened_at = db.Column(db.DateTime, nullable=True)

    presences = db.relationship("PresenceActivite", backref="session", cascade="all, delete-orphan")
    competences = db.relationship(
        "Competence",
        secondary=session_competence,
        backref=db.backref("sessions", lazy="dynamic"),
    )


class AtelierCapaciteMois(db.Model):
    __tablename__ = "atelier_capacite_mois"
    id = db.Column(db.Integer, primary_key=True)
    atelier_id = db.Column(db.Integer, db.ForeignKey("atelier_activite.id"), nullable=False)
    annee = db.Column(db.Integer, nullable=False)
    mois = db.Column(db.Integer, nullable=False)
    heures_dispo = db.Column(db.Float, nullable=False, default=0.0)
    locked = db.Column(db.Boolean, default=False)

    __table_args__ = (
        db.UniqueConstraint("atelier_id", "annee", "mois", name="uq_atelier_capacite_mois"),
    )


class PresenceActivite(db.Model):
    __tablename__ = "presence_activite"
    id = db.Column(db.Integer, primary_key=True)
    session_id = db.Column(db.Integer, db.ForeignKey("session_activite.id"), nullable=False)
    participant_id = db.Column(db.Integer, db.ForeignKey("participant.id"), nullable=False)
    participant = db.relationship("Participant")

    # Motif (liste + autre)
    motif = db.Column(db.String(180), nullable=True)
    motif_autre = db.Column(db.String(255), nullable=True)

    # signature: stockée en fichier (temp), ici juste le chemin
    signature_path = db.Column(db.String(255), nullable=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    __table_args__ = (
        db.UniqueConstraint("session_id", "participant_id", name="uq_presence_session_participant"),
    )


class Evaluation(db.Model):
    __tablename__ = "evaluation"
    id = db.Column(db.Integer, primary_key=True)
    participant_id = db.Column(db.Integer, db.ForeignKey("participant.id"), nullable=False)
    competence_id = db.Column(db.Integer, db.ForeignKey("competence.id"), nullable=False)
    session_id = db.Column(db.Integer, db.ForeignKey("session_activite.id"), nullable=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    etat = db.Column(db.Integer, nullable=False, default=0)  # 0=Non acquis, 1=En cours, 2=Acquis, 3=Expert
    date_evaluation = db.Column(db.Date, nullable=False, default=date.today)
    commentaire = db.Column(db.Text, nullable=True)

    participant = db.relationship("Participant")
    competence = db.relationship("Competence")
    session = db.relationship("SessionActivite")
    user = db.relationship("User")

    __table_args__ = (
        db.UniqueConstraint("participant_id", "competence_id", "session_id", name="uq_eval_participant_competence_session"),
    )


class ArchiveEmargement(db.Model):
    __tablename__ = "archive_emargement"
    id = db.Column(db.Integer, primary_key=True)
    secteur = db.Column(db.String(80), nullable=False, index=True)
    atelier_id = db.Column(db.Integer, db.ForeignKey("atelier_activite.id"), nullable=False)
    atelier = db.relationship("AtelierActivite")
    # pour collectif : session_id ; pour individuel mensuel : null
    session_id = db.Column(db.Integer, db.ForeignKey("session_activite.id"), nullable=True)
    annee = db.Column(db.Integer, nullable=False)
    mois = db.Column(db.Integer, nullable=True)

    docx_path = db.Column(db.String(255), nullable=True)
    pdf_path = db.Column(db.String(255), nullable=True)

    # Option : version corrigée manuellement (upload après édition Word)
    corrected_docx_path = db.Column(db.String(255), nullable=True)
    corrected_pdf_path = db.Column(db.String(255), nullable=True)

    # Suivi envoi mail
    last_emailed_to = db.Column(db.String(255), nullable=True)
    last_emailed_at = db.Column(db.DateTime, nullable=True)
    status = db.Column(db.String(20), nullable=False, default="open")  # open/locked
    created_at = db.Column(db.DateTime, default=datetime.utcnow)



class PeriodeFinancement(db.Model):
    """Périodes enregistrées (souvent calées sur un financeur) pour filtrer les stats.

    - sectorisée : une période peut être rattachée à un secteur (ou globale si secteur=None)
    - RGPD : ne contient pas de données personnelles
    """
    __tablename__ = "periode_financement"
    id = db.Column(db.Integer, primary_key=True)
    secteur = db.Column(db.String(80), nullable=True, index=True)  # None = global
    nom = db.Column(db.String(255), nullable=False)
    date_debut = db.Column(db.Date, nullable=False, index=True)
    date_fin = db.Column(db.Date, nullable=False, index=True)

    created_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    is_deleted = db.Column(db.Boolean, default=False, nullable=False, index=True)

    def __repr__(self):
        return f"<PeriodeFinancement {self.id} {self.nom} {self.date_debut}..{self.date_fin} secteur={self.secteur}>"

# ---------------------------------------------------------------------
# RBAC COMPAT: provide User.role property
# ---------------------------------------------------------------------
# Depuis l'introduction du RBAC (User.roles), certaines branches/test DB
# peuvent ne plus avoir la colonne legacy `User.role`. Or beaucoup de routes
# utilisent encore `current_user.role`.
#
# Objectif: éviter les 500/AttributeError en fournissant un fallback READ-ONLY
# (et une compat "directrice" vs "direction") tant que la migration complète
# n'est pas terminée.
#
# IMPORTANT: on n'écrase PAS un attribut SQLAlchemy existant. On ne crée ce
# fallback QUE si `User.role` n'existe pas déjà.
# ---------------------------------------------------------------------

def _role_compat_get(u) -> str:
    # 1) si une colonne legacy existe (cas ancien), on la privilégie
    legacy = getattr(u, "__dict__", {}).get("role", None)
    if legacy:
        return legacy

    # 2) sinon, on dérive depuis RBAC: premier rôle, sinon responsable_secteur
    codes = []
    try:
        codes = [r.code for r in (getattr(u, "roles", []) or []) if getattr(r, "code", None)]
    except Exception:
        codes = []
    code = (codes[0] if codes else "responsable_secteur")

    # 3) compat historique : certaines routes comparent à "directrice"
    mapping = {
        "direction": "directrice",
    }
    return mapping.get(code, code)

def _role_compat_set(u, value: str):
    # Si la colonne legacy n'existe pas, on ignore (fallback read-only).
    try:
        if "role" in getattr(u, "__dict__", {}):
            u.__dict__["role"] = value
    except Exception:
        pass

try:
    # Ne créer la propriété QUE si SQLAlchemy n'a pas déjà mappé `role`
    if not hasattr(User, "role"):
        User.role = property(_role_compat_get, _role_compat_set)  # type: ignore[attr-defined]
except Exception:
    # On ne doit jamais empêcher l'app de démarrer pour un souci de compat.
    pass
