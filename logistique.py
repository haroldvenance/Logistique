"""
PONDA LOGISTIQUE URBAIN — Système de Gestion
Version 2.4 — Améliorations :
  - Gestion robuste des erreurs (transactions, contraintes)
  - Documentation complète (docstrings et commentaires)
  - Factorisation des onglets via classe parente BaseOnglet
  - Transactions explicites dans Database
  - Optimisation des alertes : rafraîchissement partiel du dashboard
  - Réorganisation des imports (math uniquement dans PieChart)
  - Mise à jour dynamique du nom de l'entreprise dans toute l'interface
  - Adaptation pour déploiement portable (PyInstaller)
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime, date, timedelta
import os
import shutil
import threading
import time
import sys
import logging
import platform

# ── Détermination du dossier des données (portable) ─────────────────────────
def get_data_dir():
    """Retourne le dossier où doivent être stockées les données persistantes.
    En mode développement : dossier du script.
    En mode exécutable PyInstaller (--onefile) : dossier de l'exécutable."""
    if getattr(sys, 'frozen', False):
        # Exécutable unique PyInstaller
        return os.path.dirname(sys.executable)
    else:
        # Mode développement
        return os.path.dirname(os.path.abspath(__file__))

DATA_DIR = get_data_dir()
DB_PATH  = os.path.join(DATA_DIR, "ponda_data.db")
BAK_DIR  = os.path.join(DATA_DIR, "backups")
LOG_DIR  = os.path.join(DATA_DIR, "logs")

# Création des dossiers nécessaires
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(BAK_DIR, exist_ok=True)

# ── Son portable ─────────────────────────────────────────────────────────────
def play_beep():
    """Émet un bip système de manière portable."""
    if platform.system() == 'Windows':
        try:
            import winsound
            winsound.Beep(800, 500)
        except:
            print('\a', end='', flush=True)
    else:
        print('\a', end='', flush=True)

# ── Logging ─────────────────────────────────────────────────────────────────
LOG_FILE = os.path.join(LOG_DIR, "ponda.log")
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def log_action(action):
    """Enregistre une action dans le fichier de log."""
    logging.info(action)

# ── Excel ────────────────────────────────────────────────────────────────────
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    messagebox.showwarning("Module manquant", "openpyxl n'est pas installé. L'export Excel ne fonctionnera pas.\nInstallez-le avec : pip install openpyxl")

# ── PDF ──────────────────────────────────────────────────────────────────────
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Table, TableStyle, HRFlowable, PageBreak,
)

# ── CONSTANTES ───────────────────────────────────────────────────────────────
C = {
    "primaire":   "#1a3c5e",
    "secondaire": "#2e86c1",
    "accent":     "#f39c12",
    "succes":     "#27ae60",
    "danger":     "#e74c3c",
    "violet":     "#8e44ad",
    "neutre":     "#95a5a6",
    "fond":       "#f0f4f8",
    "blanc":      "#ffffff",
    "texte":      "#2c3e50",
    "bordure":    "#d5d8dc",
    "alerte_bg":  "#fdedec",
    "alerte_fg":  "#c0392b",
}

STATUTS = ["En attente", "Livré", "Échec", "Reporté"]
S_COL = {
    "En attente": "#f39c12",
    "Livré":      "#27ae60",
    "Échec":      "#e74c3c",
    "Reporté":    "#8e44ad",
}
S_BG = {
    "En attente": "#fef9e7",
    "Livré":      "#d5f5e3",
    "Échec":      "#fadbd8",
    "Reporté":    "#f3e5f5",
}

MOIS_FR = ["","Janvier","Février","Mars","Avril","Mai","Juin",
           "Juillet","Août","Septembre","Octobre","Novembre","Décembre"]

# ═══════════════════════════════════════════════════════════════════════════════
#  BASE DE DONNÉES (avec transactions explicites et gestion d'erreurs)
# ═══════════════════════════════════════════════════════════════════════════════
class Database:
    """Gestionnaire de la base de données SQLite."""
    
    def __init__(self):
        self.conn = sqlite3.connect(DB_PATH, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA journal_mode=WAL")
        self._create_tables()
        self._seed()

    def _create_tables(self):
        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS livreurs (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            nom       TEXT NOT NULL,
            prenom    TEXT NOT NULL,
            telephone TEXT,
            secteur   TEXT,
            actif     INTEGER DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS clients (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            civilite   TEXT DEFAULT 'M.',
            nom        TEXT NOT NULL,
            entreprise TEXT,
            quartier   TEXT,
            adresse    TEXT,
            telephone  TEXT,
            email      TEXT
        );
        CREATE TABLE IF NOT EXISTS quartiers (
            id  INTEGER PRIMARY KEY AUTOINCREMENT,
            nom TEXT NOT NULL UNIQUE
        );
        CREATE TABLE IF NOT EXISTS bons (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            numero       TEXT UNIQUE NOT NULL,
            client_id    INTEGER REFERENCES clients(id) ON DELETE RESTRICT,
            quartier_id  INTEGER REFERENCES quartiers(id),
            adresse      TEXT,
            telephone    TEXT,
            heure_prevue TEXT,
            livreur_id   INTEGER REFERENCES livreurs(id),
            statut       TEXT DEFAULT 'En attente',
            observations TEXT DEFAULT '',
            date_bon     TEXT DEFAULT (date('now')),
            date_modif   TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS incidents (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            bon_id      INTEGER REFERENCES bons(id) ON DELETE CASCADE,
            type_inc    TEXT DEFAULT 'Autre',
            description TEXT NOT NULL,
            resolution  TEXT DEFAULT '',
            resolu      INTEGER DEFAULT 0,
            date_inc    TEXT DEFAULT (datetime('now')),
            date_res    TEXT
        );
        CREATE TABLE IF NOT EXISTS config (
            cle   TEXT PRIMARY KEY,
            valeur TEXT
        );
        INSERT OR IGNORE INTO config VALUES('backup_interval_min','60');
        INSERT OR IGNORE INTO config VALUES('tarif_livraison','2500');
        INSERT OR IGNORE INTO config VALUES('nom_entreprise','Ponda Logistique Urbain');
        INSERT OR IGNORE INTO config VALUES('adresse_entreprise','Quartier Akwa, Douala');
        INSERT OR IGNORE INTO config VALUES('tel_entreprise','699 00 00 00');
        INSERT OR IGNORE INTO config VALUES('email_entreprise','contact@pondalogistique.cm');
        """)
        self.conn.commit()

    def _seed(self):
        """Insère des données de test si les tables sont vides."""
        cur = self.conn.cursor()
        for l in [
            ("Messi",  "Jean-Pierre", "699001122", "Bastos"),
            ("Fouda",  "Éric",        "699003344", "Melen"),
            ("Ekani",  "Sandrine",    "699005566", "Omnisport"),
            ("Mbarga", "Fatima",      "699007788", "Biyem-Assi"),
            ("Tabi",   "Rodrigue",    "699009900", "Nlongkak"),
        ]:
            cur.execute("INSERT OR IGNORE INTO livreurs(nom,prenom,telephone,secteur) VALUES(?,?,?,?)", l)
        for q in ["Bastos","Melen","Omnisport","Biyem-Assi","Nlongkak","Akwa","Bonanjo","Bonapriso"]:
            cur.execute("INSERT OR IGNORE INTO quartiers(nom) VALUES(?)", (q,))
        for c in [
            ("M.",  "Mbarga", "Pharmacie Centrale", "Bastos",    "Rue de la Paix",    "699010203",""),
            ("Mme", "Ekani",  "Superette Mama",     "Melen",     "Av. des Cocotiers", "699040506",""),
            ("M.",  "Fouda",  "Boutique Fleurs",    "Omnisport", "Rue du Stade",      "699070809",""),
        ]:
            cur.execute("INSERT OR IGNORE INTO clients(civilite,nom,entreprise,quartier,adresse,telephone,email) VALUES(?,?,?,?,?,?,?)", c)
        self.conn.commit()

    def cfg(self, cle): 
        r = self.conn.execute("SELECT valeur FROM config WHERE cle=?", (cle,)).fetchone()
        return r[0] if r else ""

    def set_cfg(self, cle, valeur):
        self.conn.execute("INSERT OR REPLACE INTO config VALUES(?,?)", (cle, valeur))
        self.conn.commit()
        log_action(f"Configuration modifiée : {cle} = {valeur}")

    # --- Livreurs ---
    def get_livreurs(self, actif_only=True):
        q = "SELECT * FROM livreurs" + (" WHERE actif=1" if actif_only else "") + " ORDER BY nom"
        return self.conn.execute(q).fetchall()

    def add_livreur(self, nom, prenom, tel, secteur):
        self.conn.execute("INSERT INTO livreurs(nom,prenom,telephone,secteur) VALUES(?,?,?,?)",
                          (nom, prenom, tel, secteur))
        self.conn.commit()
        log_action(f"Ajout livreur : {nom} {prenom}")

    def update_livreur(self, lid, nom, prenom, tel, secteur, actif=1):
        self.conn.execute("UPDATE livreurs SET nom=?,prenom=?,telephone=?,secteur=?,actif=? WHERE id=?",
                          (nom, prenom, tel, secteur, actif, lid))
        self.conn.commit()
        log_action(f"Modification livreur id {lid}")

    # --- Clients ---
    def get_clients(self):
        return self.conn.execute("SELECT * FROM clients ORDER BY nom").fetchall()

    def add_client(self, civilite, nom, entreprise, quartier, adresse, telephone, email=""):
        self.conn.execute(
            "INSERT INTO clients(civilite,nom,entreprise,quartier,adresse,telephone,email) VALUES(?,?,?,?,?,?,?)",
            (civilite, nom, entreprise, quartier, adresse, telephone, email))
        self.conn.commit()
        log_action(f"Ajout client : {civilite} {nom}")

    def update_client(self, cid, civilite, nom, entreprise, quartier, adresse, telephone, email=""):
        self.conn.execute(
            "UPDATE clients SET civilite=?,nom=?,entreprise=?,quartier=?,adresse=?,telephone=?,email=? WHERE id=?",
            (civilite, nom, entreprise, quartier, adresse, telephone, email, cid))
        self.conn.commit()
        log_action(f"Modification client id {cid}")

    def delete_client(self, cid):
        """Supprime un client et tous ses bons (grâce à ON DELETE CASCADE)."""
        # Vérifier d'abord s'il existe des bons pour ce client
        bons = self.conn.execute("SELECT COUNT(*) FROM bons WHERE client_id=?", (cid,)).fetchone()[0]
        if bons > 0:
            raise ValueError(f"Ce client possède {bons} bon(s). Supprimez d'abord ses bons ou changez de client.")
        try:
            self.conn.execute("DELETE FROM clients WHERE id=?", (cid,))
            self.conn.commit()
            log_action(f"Suppression client id {cid}")
        except sqlite3.IntegrityError as e:
            self.conn.rollback()
            raise ValueError(f"Impossible de supprimer le client : {e}")

    def get_client_stats_mois(self, client_id, annee, mois):
        prefix = f"{annee}-{str(mois).zfill(2)}"
        return self.conn.execute("""
            SELECT COUNT(*) total,
                   SUM(CASE WHEN statut='Livré' THEN 1 ELSE 0 END) livres,
                   SUM(CASE WHEN statut='Échec' THEN 1 ELSE 0 END) echecs
            FROM bons WHERE client_id=? AND date_bon LIKE ?
        """, (client_id, f"{prefix}%")).fetchone()

    # --- Quartiers ---
    def get_quartiers(self):
        return self.conn.execute("SELECT * FROM quartiers ORDER BY nom").fetchall()

    # --- Bons ---
    def _bons_query(self, where, params):
        return self.conn.execute(f"""
            SELECT b.*,
                   COALESCE(c.civilite||' '||c.nom,'—')          AS client_nom,
                   COALESCE(c.entreprise,'')                       AS client_ent,
                   COALESCE(q.nom,'—')                             AS quartier_nom,
                   COALESCE(l.nom||' '||l.prenom,'Non assigné')   AS livreur_nom,
                   c.telephone AS client_tel, c.adresse AS client_adresse,
                   c.email AS client_email
            FROM bons b
            LEFT JOIN clients   c ON b.client_id   = c.id
            LEFT JOIN quartiers q ON b.quartier_id = q.id
            LEFT JOIN livreurs  l ON b.livreur_id  = l.id
            {where} ORDER BY b.heure_prevue
        """, params).fetchall()

    def get_bons_jour(self, jour=None):
        return self._bons_query("WHERE b.date_bon=?", (jour or date.today().isoformat(),))

    def get_bons_mois(self, annee, mois):
        return self._bons_query("WHERE b.date_bon LIKE ?", (f"{annee}-{str(mois).zfill(2)}%",))

    def get_bons_retard(self):
        now_h = datetime.now().strftime("%H:%M")
        today = date.today().isoformat()
        return self._bons_query(
            "WHERE b.statut='En attente' AND b.date_bon=? AND b.heure_prevue<=?",
            (today, now_h)
        )

    def add_bon(self, numero, client_id, quartier_id, adresse, telephone, heure_prevue, livreur_id, date_bon):
        try:
            self.conn.execute("""INSERT INTO bons(numero,client_id,quartier_id,adresse,telephone,
                heure_prevue,livreur_id,date_bon) VALUES(?,?,?,?,?,?,?,?)""",
                (numero, client_id, quartier_id, adresse, telephone, heure_prevue, livreur_id, date_bon))
            self.conn.commit()
            log_action(f"Ajout bon {numero}")
        except sqlite3.IntegrityError as e:
            self.conn.rollback()
            raise ValueError(f"Numéro de bon déjà existant ou clé étrangère invalide : {e}")

    def update_statut(self, bon_id, statut, obs=""):
        self.conn.execute("UPDATE bons SET statut=?,observations=?,date_modif=datetime('now') WHERE id=?",
                          (statut, obs, bon_id))
        self.conn.commit()
        log_action(f"Modification statut bon {bon_id} -> {statut}")

    def update_bon(self, bon_id, livreur_id, heure_prevue, obs):
        self.conn.execute("UPDATE bons SET livreur_id=?,heure_prevue=?,observations=?,date_modif=datetime('now') WHERE id=?",
                          (livreur_id, heure_prevue, obs, bon_id))
        self.conn.commit()
        log_action(f"Modification bon {bon_id}")

    def delete_bon(self, bon_id):
        try:
            self.conn.execute("DELETE FROM bons WHERE id=?", (bon_id,))
            self.conn.commit()
            log_action(f"Suppression bon {bon_id}")
        except sqlite3.IntegrityError as e:
            self.conn.rollback()
            raise ValueError(f"Impossible de supprimer le bon : {e}")

    def get_prochain_numero(self, jour=None):
        jour = jour or date.today().isoformat()
        n = self.conn.execute("SELECT COUNT(*) FROM bons WHERE date_bon=?", (jour,)).fetchone()[0]
        d = datetime.strptime(jour, "%Y-%m-%d")
        return f"BON-{d.strftime('%Y%m%d')}-{str(n+1).zfill(3)}"

    # --- Statistiques ---
    def get_stats_jour(self, jour=None):
        jour = jour or date.today().isoformat()
        r = self.conn.execute("""SELECT COUNT(*) total,
            SUM(CASE WHEN statut='Livré'     THEN 1 ELSE 0 END) livres,
            SUM(CASE WHEN statut='Échec'     THEN 1 ELSE 0 END) echecs,
            SUM(CASE WHEN statut='Reporté'   THEN 1 ELSE 0 END) reportes,
            SUM(CASE WHEN statut='En attente'THEN 1 ELSE 0 END) attente
            FROM bons WHERE date_bon=?""", (jour,)).fetchone()
        return dict(r) if r else {}

    def get_stats_mois(self, annee, mois):
        prefix = f"{annee}-{str(mois).zfill(2)}"
        r = self.conn.execute("""SELECT COUNT(*) total,
            SUM(CASE WHEN statut='Livré' THEN 1 ELSE 0 END) livres,
            SUM(CASE WHEN statut='Échec' THEN 1 ELSE 0 END) echecs,
            SUM(CASE WHEN statut='Reporté' THEN 1 ELSE 0 END) reportes
            FROM bons WHERE date_bon LIKE ?""", (f"{prefix}%",)).fetchone()
        return dict(r) if r else {}

    def get_stats_livreurs_jour(self, jour=None):
        jour = jour or date.today().isoformat()
        return self.conn.execute("""
            SELECT l.nom||' '||l.prenom livreur,
                   COUNT(*) total,
                   SUM(CASE WHEN b.statut='Livré' THEN 1 ELSE 0 END) livres,
                   SUM(CASE WHEN b.statut='Échec' THEN 1 ELSE 0 END) echecs
            FROM bons b JOIN livreurs l ON b.livreur_id=l.id
            WHERE b.date_bon=? GROUP BY l.id ORDER BY livres DESC
        """, (jour,)).fetchall()

    def get_stats_livreurs_mois(self, annee, mois):
        prefix = f"{annee}-{str(mois).zfill(2)}"
        return self.conn.execute("""
            SELECT l.nom||' '||l.prenom livreur, l.id livreur_id,
                   COUNT(*) total,
                   SUM(CASE WHEN b.statut='Livré' THEN 1 ELSE 0 END) livres,
                   SUM(CASE WHEN b.statut='Échec' THEN 1 ELSE 0 END) echecs,
                   SUM(CASE WHEN b.statut='Reporté' THEN 1 ELSE 0 END) reportes
            FROM bons b JOIN livreurs l ON b.livreur_id=l.id
            WHERE b.date_bon LIKE ? GROUP BY l.id ORDER BY livres DESC
        """, (f"{prefix}%",)).fetchall()

    def get_stats_quartiers_mois(self, annee, mois):
        prefix = f"{annee}-{str(mois).zfill(2)}"
        return self.conn.execute("""
            SELECT COALESCE(q.nom,'Inconnu') quartier,
                   COUNT(*) total,
                   SUM(CASE WHEN b.statut='Échec' THEN 1 ELSE 0 END) echecs
            FROM bons b LEFT JOIN quartiers q ON b.quartier_id=q.id
            WHERE b.date_bon LIKE ? GROUP BY q.id ORDER BY echecs DESC
        """, (f"{prefix}%",)).fetchall()

    def get_stats_quartiers_mois_simple(self, annee, mois):
        """Retourne liste (quartier, total, livres, echecs) pour le mois."""
        prefix = f"{annee}-{str(mois).zfill(2)}"
        return self.conn.execute("""
            SELECT COALESCE(q.nom,'Inconnu') quartier,
                   COUNT(*) total,
                   SUM(CASE WHEN b.statut='Livré' THEN 1 ELSE 0 END) livres,
                   SUM(CASE WHEN b.statut='Échec' THEN 1 ELSE 0 END) echecs
            FROM bons b LEFT JOIN quartiers q ON b.quartier_id=q.id
            WHERE b.date_bon LIKE ?
            GROUP BY q.id
            ORDER BY total DESC
        """, (f"{prefix}%",)).fetchall()

    # --- Incidents ---
    def get_incidents(self, resolu=None):
        w = "" if resolu is None else f" WHERE i.resolu={int(resolu)}"
        return self.conn.execute(f"""
            SELECT i.*, b.numero bon_numero,
                   COALESCE(l.nom||' '||l.prenom,'—') livreur_nom
            FROM incidents i
            LEFT JOIN bons b ON i.bon_id=b.id
            LEFT JOIN livreurs l ON b.livreur_id=l.id
            {w} ORDER BY i.date_inc DESC
        """).fetchall()

    def add_incident(self, bon_id, type_inc, description, resolution="", resolu=0):
        self.conn.execute(
            "INSERT INTO incidents(bon_id,type_inc,description,resolution,resolu) VALUES(?,?,?,?,?)",
            (bon_id or None, type_inc, description, resolution, resolu))
        self.conn.commit()
        log_action(f"Ajout incident : {type_inc} - {description[:30]}")

    def update_incident(self, iid, type_inc, description, resolution, resolu):
        date_res = datetime.now().isoformat() if resolu else None
        self.conn.execute(
            "UPDATE incidents SET type_inc=?,description=?,resolution=?,resolu=?,date_res=? WHERE id=?",
            (type_inc, description, resolution, resolu, date_res, iid))
        self.conn.commit()
        log_action(f"Modification incident {iid}")

    def delete_incident(self, iid):
        self.conn.execute("DELETE FROM incidents WHERE id=?", (iid,))
        self.conn.commit()
        log_action(f"Suppression incident {iid}")

    def count_incidents_ouverts(self):
        return self.conn.execute("SELECT COUNT(*) FROM incidents WHERE resolu=0").fetchone()[0]

    # --- Transactions multi-étapes (exemple) ---
    def execute_transaction(self, queries):
        """Exécute plusieurs requêtes dans une transaction. En cas d'erreur, rollback."""
        try:
            for sql, params in queries:
                self.conn.execute(sql, params)
            self.conn.commit()
            return True
        except Exception as e:
            self.conn.rollback()
            log_action(f"Erreur transaction : {e}")
            raise

# ═══════════════════════════════════════════════════════════════════════════════
#  SAUVEGARDE AUTOMATIQUE (inchangé)
# ═══════════════════════════════════════════════════════════════════════════════
class BackupManager:
    """Gère les sauvegardes automatiques et manuelles de la base de données."""
    def __init__(self, db, status_callback=None):
        self.db = db
        self.cb = status_callback
        self._stop = threading.Event()
        self._thread = None
        os.makedirs(BAK_DIR, exist_ok=True)

    def backup_now(self):
        """Effectue une sauvegarde immédiate."""
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = os.path.join(BAK_DIR, f"ponda_backup_{ts}.db")
        try:
            shutil.copy2(DB_PATH, dest)
            self._nettoyer_anciens(keep=10)
            if self.cb:
                self.cb(f"✅ Sauvegarde : {os.path.basename(dest)}")
            log_action(f"Sauvegarde effectuée : {dest}")
            return dest
        except Exception as e:
            if self.cb:
                self.cb(f"⚠️ Erreur sauvegarde : {e}")
            log_action(f"Erreur sauvegarde : {e}")
            return None

    def _nettoyer_anciens(self, keep=10):
        fichiers = sorted(
            [f for f in os.listdir(BAK_DIR) if f.endswith(".db")],
            reverse=True
        )
        for f in fichiers[keep:]:
            try:
                os.remove(os.path.join(BAK_DIR, f))
            except Exception:
                pass

    def start_auto(self, interval_min=60):
        self._stop.clear()
        def _loop():
            while not self._stop.wait(interval_min * 60):
                self.backup_now()
        self._thread = threading.Thread(target=_loop, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop.set()

    def liste_backups(self):
        if not os.path.exists(BAK_DIR):
            return []
        return sorted(
            [f for f in os.listdir(BAK_DIR) if f.endswith(".db")],
            reverse=True
        )

# ═══════════════════════════════════════════════════════════════════════════════
#  GÉNÉRATION PDF (documentation ajoutée)
# ═══════════════════════════════════════════════════════════════════════════════
class PDF:
    """Classe utilitaire pour générer des rapports PDF avec ReportLab."""
    BLEU  = colors.HexColor("#1a3c5e")
    CLAIR = colors.HexColor("#e8f4f8")
    VERT  = colors.HexColor("#27ae60")
    ROUGE = colors.HexColor("#e74c3c")
    OR    = colors.HexColor("#f39c12")
    VIOL  = colors.HexColor("#8e44ad")

    @staticmethod
    def _doc(path, titre):
        return SimpleDocTemplate(path, pagesize=A4,
            leftMargin=2*cm, rightMargin=2*cm,
            topMargin=2*cm, bottomMargin=2*cm, title=titre)

    @classmethod
    def _entete_story(cls, db, titre, sous_titre, story, styles):
        ts = ParagraphStyle("ts", fontSize=17, textColor=cls.BLEU, fontName="Helvetica-Bold", spaceAfter=2)
        ss = ParagraphStyle("ss", fontSize=9,  textColor=colors.grey, spaceAfter=14)
        story.append(Paragraph(db.cfg("nom_entreprise").upper(), ts))
        story.append(Paragraph(f"{titre} — {sous_titre}", ss))
        story.append(HRFlowable(width="100%", thickness=2, color=cls.BLEU))
        story.append(Spacer(1, 0.4*cm))

    @classmethod
    def _pied(cls, story):
        story.append(Spacer(1, 1*cm))
        story.append(HRFlowable(width="100%", thickness=1, color=cls.BLEU))
        story.append(Paragraph(
            f"Document généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')} — Confidentiel",
            ParagraphStyle("pied", fontSize=7, textColor=colors.grey, alignment=1)
        ))

    @classmethod
    def rapport_journalier(cls, db, jour, path):
        """Génère un rapport PDF pour une date donnée."""
        stats = db.get_stats_jour(jour)
        bons  = db.get_bons_jour(jour)
        lvrrs = db.get_stats_livreurs_jour(jour)
        total  = stats.get("total",0) or 0
        livres = stats.get("livres",0) or 0
        taux   = round(livres/total*100) if total else 0

        doc    = cls._doc(path, f"Rapport {jour}")
        styles = getSampleStyleSheet()
        story  = []
        h2 = ParagraphStyle("h2", fontSize=11, textColor=cls.BLEU, fontName="Helvetica-Bold",
                             spaceBefore=12, spaceAfter=5)
        cls._entete_story(db, "Rapport Journalier",
                          datetime.strptime(jour,"%Y-%m-%d").strftime("%d/%m/%Y"), story, styles)

        # KPIs
        story.append(Paragraph("1. SYNTHÈSE", h2))
        kpis = [
            ["Total bons","Livrés","Échecs","Reportés","En attente","Taux réussite"],
            [str(total), str(livres), str(stats.get("echecs",0) or 0),
             str(stats.get("reportes",0) or 0), str(stats.get("attente",0) or 0), f"{taux}%"]
        ]
        tk_tab = Table(kpis, colWidths=[2.8*cm]*6)
        tk_tab.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0), cls.BLEU),
            ("TEXTCOLOR", (0,0),(-1,0), colors.white),
            ("FONTNAME",  (0,0),(-1,0), "Helvetica-Bold"),
            ("BACKGROUND",(0,1),(-1,1), cls.CLAIR),
            ("FONTNAME",  (0,1),(-1,1), "Helvetica-Bold"),
            ("TEXTCOLOR", (5,1),(5,1),  cls.VERT if taux>=80 else cls.ROUGE),
            ("ALIGN",     (0,0),(-1,-1),"CENTER"),
            ("GRID",      (0,0),(-1,-1),0.5, colors.HexColor("#d5d8dc")),
            ("PADDING",   (0,0),(-1,-1),8),
        ]))
        story.append(tk_tab); story.append(Spacer(1,0.4*cm))

        # Livreurs
        story.append(Paragraph("2. PERFORMANCE PAR LIVREUR", h2))
        if lvrrs:
            lv = [["Livreur","Assignés","Livrés","Échecs","Taux"]]
            for r in lvrrs:
                tx = f"{round(r['livres']/r['total']*100)}%" if r["total"] else "0%"
                lv.append([r["livreur"],str(r["total"]),str(r["livres"]),str(r["echecs"]),tx])
            tl = Table(lv, colWidths=[6*cm,2.5*cm,2.5*cm,2.5*cm,2.5*cm])
            tl.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),cls.BLEU),("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#f9f9f9")]),
                ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#d5d8dc")),
                ("ALIGN",(1,0),(-1,-1),"CENTER"),("PADDING",(0,0),(-1,-1),6),
            ]))
            story.append(tl)
        story.append(Spacer(1,0.4*cm))

        # Détail bons
        story.append(Paragraph("3. DÉTAIL DES BONS", h2))
        if bons:
            bd = [["N° Bon","Client","Quartier","Livreur","Heure","Statut"]]
            for b in bons:
                bd.append([b["numero"], str(b["client_nom"] or "—")[:22],
                           str(b["quartier_nom"] or "—"), str(b["livreur_nom"] or "—")[:16],
                           b["heure_prevue"] or "—", b["statut"]])
            tb = Table(bd, colWidths=[3.2*cm,4*cm,2.5*cm,3.5*cm,1.8*cm,2.5*cm])
            row_styles = [
                ("BACKGROUND",(0,0),(-1,0),cls.BLEU),("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#f9f9f9")]),
                ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#d5d8dc")),
                ("FONTSIZE",(0,0),(-1,-1),8),("PADDING",(0,0),(-1,-1),5),
            ]
            for i,b in enumerate(bons,1):
                bg = {"Livré":colors.HexColor("#d5f5e3"),"Échec":colors.HexColor("#fadbd8"),
                      "Reporté":colors.HexColor("#f3e5f5")}.get(b["statut"],colors.white)
                row_styles.append(("BACKGROUND",(5,i),(5,i),bg))
            tb.setStyle(TableStyle(row_styles))
            story.append(tb)

        cas = [b for b in bons if b["observations"]]
        if cas:
            story.append(Spacer(1,0.4*cm))
            story.append(Paragraph("4. CAS PARTICULIERS", h2))
            for b in cas:
                story.append(Paragraph(
                    f"<b>{b['numero']}</b> ({b['statut']}) : {b['observations']}",
                    styles["Normal"]))
                story.append(Spacer(1,0.15*cm))

        cls._pied(story)
        doc.build(story)

    @classmethod
    def rapport_mensuel(cls, db, annee, mois, path):
        """Génère un rapport PDF pour le mois donné."""
        stats = db.get_stats_mois(annee, mois)
        bons  = db.get_bons_mois(annee, mois)
        lvrrs = db.get_stats_livreurs_mois(annee, mois)
        total  = stats.get("total",0) or 0
        livres = stats.get("livres",0) or 0
        taux   = round(livres/total*100) if total else 0

        doc    = cls._doc(path, f"Rapport {MOIS_FR[mois]} {annee}")
        styles = getSampleStyleSheet()
        story  = []
        h2 = ParagraphStyle("h2", fontSize=11, textColor=cls.BLEU, fontName="Helvetica-Bold",
                             spaceBefore=12, spaceAfter=5)
        cls._entete_story(db, "Rapport Mensuel",
                          f"{MOIS_FR[mois]} {annee}", story, styles)

        # KPIs
        story.append(Paragraph("1. SYNTHÈSE", h2))
        kpis = [
            ["Total bons","Livrés","Échecs","Reportés","Taux réussite"],
            [str(total), str(livres), str(stats.get("echecs",0) or 0),
             str(stats.get("reportes",0) or 0), f"{taux}%"]
        ]
        tk_tab = Table(kpis, colWidths=[3*cm]*5)
        tk_tab.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0), cls.BLEU),
            ("TEXTCOLOR", (0,0),(-1,0), colors.white),
            ("FONTNAME",  (0,0),(-1,0), "Helvetica-Bold"),
            ("BACKGROUND",(0,1),(-1,1), cls.CLAIR),
            ("FONTNAME",  (0,1),(-1,1), "Helvetica-Bold"),
            ("TEXTCOLOR", (4,1),(4,1),  cls.VERT if taux>=80 else cls.ROUGE),
            ("ALIGN",     (0,0),(-1,-1),"CENTER"),
            ("GRID",      (0,0),(-1,-1),0.5, colors.HexColor("#d5d8dc")),
            ("PADDING",   (0,0),(-1,-1),8),
        ]))
        story.append(tk_tab); story.append(Spacer(1,0.4*cm))

        # Livreurs
        story.append(Paragraph("2. PERFORMANCE PAR LIVREUR", h2))
        if lvrrs:
            lv = [["Livreur","Assignés","Livrés","Échecs","Taux"]]
            for r in lvrrs:
                tx = f"{round(r['livres']/r['total']*100)}%" if r["total"] else "0%"
                lv.append([r["livreur"],str(r["total"]),str(r["livres"]),str(r["echecs"]),tx])
            tl = Table(lv, colWidths=[6*cm,2.5*cm,2.5*cm,2.5*cm,2.5*cm])
            tl.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),cls.BLEU),("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#f9f9f9")]),
                ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#d5d8dc")),
                ("ALIGN",(1,0),(-1,-1),"CENTER"),("PADDING",(0,0),(-1,-1),6),
            ]))
            story.append(tl)
        story.append(Spacer(1,0.4*cm))

        # Détail bons (limité aux 30 premiers pour éviter un PDF trop long)
        story.append(Paragraph("3. DÉTAIL DES BONS (30 premiers)", h2))
        if bons:
            bd = [["N° Bon","Client","Quartier","Livreur","Heure","Statut"]]
            for b in bons[:30]:
                bd.append([b["numero"], str(b["client_nom"] or "—")[:22],
                           str(b["quartier_nom"] or "—"), str(b["livreur_nom"] or "—")[:16],
                           b["heure_prevue"] or "—", b["statut"]])
            tb = Table(bd, colWidths=[3.2*cm,4*cm,2.5*cm,3.5*cm,1.8*cm,2.5*cm])
            row_styles = [
                ("BACKGROUND",(0,0),(-1,0),cls.BLEU),("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#f9f9f9")]),
                ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#d5d8dc")),
                ("FONTSIZE",(0,0),(-1,-1),8),("PADDING",(0,0),(-1,-1),5),
            ]
            for i,b in enumerate(bons[:30],1):
                bg = {"Livré":colors.HexColor("#d5f5e3"),"Échec":colors.HexColor("#fadbd8"),
                      "Reporté":colors.HexColor("#f3e5f5")}.get(b["statut"],colors.white)
                row_styles.append(("BACKGROUND",(5,i),(5,i),bg))
            tb.setStyle(TableStyle(row_styles))
            story.append(tb)
            if len(bons) > 30:
                story.append(Paragraph("... et d'autres bons non affichés pour des raisons de taille.", styles["Normal"]))

        cls._pied(story)
        doc.build(story)

    @classmethod
    def facture(cls, db, bon_id, path, tarif=2500):
        """Génère une facture PDF pour un bon donné."""
        b = db.conn.execute("""
            SELECT b.*, c.civilite, c.nom c_nom, c.entreprise, c.adresse c_adresse,
                   c.telephone c_tel, c.email c_email, q.nom quartier_nom,
                   l.nom||' '||l.prenom livreur_nom
            FROM bons b
            LEFT JOIN clients c ON b.client_id=c.id
            LEFT JOIN quartiers q ON b.quartier_id=q.id
            LEFT JOIN livreurs l ON b.livreur_id=l.id
            WHERE b.id=?""", (bon_id,)).fetchone()
        if not b: return False

        doc = cls._doc(path, f"Facture {b['numero']}")
        styles = getSampleStyleSheet()
        story = []
        ts = ParagraphStyle("ts",fontSize=18,textColor=cls.BLEU,fontName="Helvetica-Bold")
        sub = ParagraphStyle("sub",fontSize=9,textColor=colors.grey)

        hdr = Table([[
            Paragraph(db.cfg("nom_entreprise").upper(), ts),
            Paragraph(f"<b>FACTURE N°</b><br/>{b['numero'].replace('BON','FAC')}",
                      ParagraphStyle("fn",fontSize=13,alignment=2,textColor=cls.BLEU))
        ],[
            Paragraph(f"{db.cfg('adresse_entreprise')}<br/>Tél: {db.cfg('tel_entreprise')}<br/>{db.cfg('email_entreprise')}", sub),
            Paragraph(f"Date : {datetime.strptime(b['date_bon'],'%Y-%m-%d').strftime('%d/%m/%Y')}",
                      ParagraphStyle("fd",fontSize=10,alignment=2))
        ]], colWidths=[10*cm, 7*cm])
        hdr.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
        story += [hdr, Spacer(1,0.4*cm),
                  HRFlowable(width="100%",thickness=2,color=cls.BLEU), Spacer(1,0.4*cm)]

        bold = ParagraphStyle("bl",fontSize=10,fontName="Helvetica-Bold")
        story.append(Paragraph("FACTURER À :", bold))
        story.append(Paragraph(
            f"{b['civilite']} {b['c_nom']}" + (f" — {b['entreprise']}" if b["entreprise"] else ""),
            styles["Normal"]))
        if b["c_adresse"]: story.append(Paragraph(b["c_adresse"], styles["Normal"]))
        story.append(Paragraph(b["quartier_nom"] or "", styles["Normal"]))
        story.append(Spacer(1,0.5*cm))

        montant = tarif; tva = round(montant*0.1925); total_ttc = montant+tva
        items = [
            ["Description","Qté","P.U (XAF)","Total (XAF)"],
            [f"Livraison urbaine\nBon {b['numero']} — {b['quartier_nom']}",
             "1", f"{montant:,}".replace(",","  "), f"{montant:,}".replace(",","  ")],
        ]
        ti = Table(items, colWidths=[8*cm,1.5*cm,3.5*cm,4*cm])
        ti.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),cls.BLEU),("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#d5d8dc")),
            ("ALIGN",(1,0),(-1,-1),"RIGHT"),("PADDING",(0,0),(-1,-1),8),
        ]))
        story.append(ti); story.append(Spacer(1,0.3*cm))

        tots = Table([
            ["","Sous-total :", f"{montant:,} XAF".replace(",","  ")],
            ["","TVA (19,25%) :", f"{tva:,} XAF".replace(",","  ")],
            ["","TOTAL TTC :", f"{total_ttc:,} XAF".replace(",","  ")],
        ], colWidths=[8*cm,4*cm,5*cm])
        tots.setStyle(TableStyle([
            ("ALIGN",(1,0),(-1,-1),"RIGHT"),
            ("FONTNAME",(0,2),(-1,2),"Helvetica-Bold"),
            ("FONTSIZE",(0,2),(-1,2),12),
            ("TEXTCOLOR",(0,2),(-1,2),cls.BLEU),
            ("BACKGROUND",(0,2),(-1,2),cls.CLAIR),
            ("PADDING",(0,0),(-1,-1),5),
        ]))
        story.append(tots)
        story.append(Spacer(1,2*cm))
        story.append(HRFlowable(width="100%",thickness=1,color=cls.BLEU))
        story.append(Paragraph("Merci de votre confiance — "+db.cfg("nom_entreprise"),
                                ParagraphStyle("merci",fontSize=9,textColor=colors.grey,alignment=1)))
        doc.build(story)
        return True

    @classmethod
    def publipostage(cls, db, annee, mois, path, tarif=2500):
        """Génère des courriers de bilan pour tous les clients ayant eu des livraisons ce mois."""
        clients = db.get_clients()
        styles  = getSampleStyleSheet()
        doc     = cls._doc(path, f"Courriers Clients {MOIS_FR[mois]} {annee}")
        story   = []
        ts      = ParagraphStyle("ts",fontSize=14,textColor=cls.BLEU,fontName="Helvetica-Bold")
        sub     = ParagraphStyle("sub",fontSize=9,textColor=colors.grey)
        corps   = ParagraphStyle("corps",fontSize=10,leading=16,spaceAfter=10)
        sig     = ParagraphStyle("sig",fontSize=10,leading=16)

        entreprise = db.cfg("nom_entreprise")
        adresse_e  = db.cfg("adresse_entreprise")
        tel_e      = db.cfg("tel_entreprise")
        email_e    = db.cfg("email_entreprise")
        date_str   = datetime.now().strftime("%d/%m/%Y")

        nb_lettres = 0
        for cl in clients:
            stats = db.get_client_stats_mois(cl["id"], annee, mois)
            total  = stats["total"]  or 0
            livres = stats["livres"] or 0
            echecs = stats["echecs"] or 0
            if total == 0:
                continue

            taux = round(livres/total*100) if total else 0
            nom_client = f"{cl['civilite']} {cl['nom']}"
            ent_client = cl["entreprise"] or ""

            entete_tab = Table([[
                Paragraph(f"<b>{entreprise.upper()}</b><br/>{adresse_e}<br/>Tél : {tel_e}<br/>{email_e}", sub),
                Paragraph(f"Douala, le {date_str}", ParagraphStyle("dr",fontSize=10,alignment=2)),
            ]], colWidths=[9*cm,8*cm])
            entete_tab.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
            story.append(entete_tab)
            story.append(Spacer(1,0.6*cm))
            story.append(Paragraph(f"<b>À l'attention de :</b>", ParagraphStyle("at",fontSize=9,textColor=colors.grey)))
            story.append(Paragraph(f"<b>{nom_client}</b>", ParagraphStyle("dc",fontSize=11,fontName="Helvetica-Bold")))
            if ent_client:
                story.append(Paragraph(ent_client, styles["Normal"]))
            if cl["quartier"]:
                story.append(Paragraph(f"{cl['quartier']}, Douala", styles["Normal"]))
            story.append(Spacer(1,0.6*cm))
            story.append(HRFlowable(width="100%",thickness=1,color=colors.HexColor("#d5d8dc")))
            story.append(Spacer(1,0.4*cm))

            story.append(Paragraph(
                f"<b>Objet :</b> Bilan mensuel de vos livraisons — {MOIS_FR[mois]} {annee}",
                ParagraphStyle("obj",fontSize=10,fontName="Helvetica-Bold",spaceAfter=12)))

            story.append(Paragraph(f"{nom_client},", corps))
            story.append(Paragraph(
                f"Nous avons le plaisir de vous adresser le bilan de vos livraisons pour le mois "
                f"de <b>{MOIS_FR[mois]} {annee}</b>.",
                corps))

            bilan = [
                ["Bons traités", "Livrés", "Échecs", "Taux de réussite"],
                [str(total), str(livres), str(echecs), f"{taux}%"],
            ]
            tb = Table(bilan, colWidths=[4*cm]*4)
            tb.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),cls.BLEU),("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("BACKGROUND",(0,1),(-1,1),cls.CLAIR),
                ("FONTNAME",(0,1),(-1,1),"Helvetica-Bold"),
                ("TEXTCOLOR",(3,1),(3,1),cls.VERT if taux>=80 else cls.ROUGE),
                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#d5d8dc")),
                ("PADDING",(0,0),(-1,-1),8),
            ]))
            story.append(tb); story.append(Spacer(1,0.4*cm))

            if echecs > 0:
                story.append(Paragraph(
                    f"Nous notons {echecs} livraison(s) non aboutie(s) ce mois et nous nous engageons "
                    "à améliorer notre taux de réussite pour le mois à venir.",
                    corps))

            story.append(Paragraph(
                "Nous vous remercions de la confiance que vous accordez à nos services "
                "et restons à votre disposition pour tout renseignement.",
                corps))
            story.append(Spacer(1,0.6*cm))
            story.append(Paragraph("Veuillez agréer, " + nom_client +
                                    ", l'expression de nos salutations distinguées.", corps))
            story.append(Spacer(1,0.8*cm))
            story.append(Paragraph("Pour Ponda Logistique Urbain,", sig))
            story.append(Paragraph("<b>Le Chef d'Équipe</b>", sig))
            story.append(PageBreak())
            nb_lettres += 1

        if nb_lettres == 0:
            story.append(Paragraph("Aucun client avec livraisons ce mois.", styles["Normal"]))

        doc.build(story)
        return nb_lettres

# ═══════════════════════════════════════════════════════════════════════════════
#  WIDGETS RÉUTILISABLES (inchangés)
# ═══════════════════════════════════════════════════════════════════════════════
class Carte(tk.Frame):
    """Carte KPI affichant un titre et une valeur sur fond coloré."""
    def __init__(self, p, titre, valeur, couleur, **kw):
        super().__init__(p, bg=couleur, padx=16, pady=12, **kw)
        tk.Label(self,text=titre,bg=couleur,fg="white",font=("Helvetica",8)).pack(anchor="w")
        self._v=tk.Label(self,text=str(valeur),bg=couleur,fg="white",font=("Helvetica",24,"bold"))
        self._v.pack(anchor="w")
    def update(self,v): self._v.config(text=str(v))

class Btn(tk.Button):
    """Bouton standard avec thème Ponda."""
    def __init__(self,p,**kw):
        kw.setdefault("bg",C["secondaire"]); kw.setdefault("fg","white")
        kw.setdefault("font",("Helvetica",9,"bold")); kw.setdefault("relief","flat")
        kw.setdefault("padx",12); kw.setdefault("pady",6); kw.setdefault("cursor","hand2")
        kw.setdefault("activebackground",C["primaire"]); kw.setdefault("activeforeground","white")
        super().__init__(p,**kw)

class BtnDanger(tk.Button):
    """Bouton rouge pour les actions dangereuses."""
    def __init__(self,p,**kw):
        kw.setdefault("bg",C["danger"]); kw.setdefault("fg","white")
        kw.setdefault("font",("Helvetica",9,"bold")); kw.setdefault("relief","flat")
        kw.setdefault("padx",12); kw.setdefault("pady",6); kw.setdefault("cursor","hand2")
        kw.setdefault("activebackground","#c0392b"); kw.setdefault("activeforeground","white")
        super().__init__(p,**kw)

class BtnVert(tk.Button):
    """Bouton vert pour actions positives."""
    def __init__(self,p,**kw):
        kw.setdefault("bg",C["succes"]); kw.setdefault("fg","white")
        kw.setdefault("font",("Helvetica",9,"bold")); kw.setdefault("relief","flat")
        kw.setdefault("padx",12); kw.setdefault("pady",6); kw.setdefault("cursor","hand2")
        super().__init__(p,**kw)

def barre_titre(parent, icone, texte, bg=None):
    """Crée une barre de titre avec une icône et un texte."""
    bg = bg or C["primaire"]
    fr = tk.Frame(parent, bg=bg, pady=10)
    fr.pack(fill="x")
    tk.Label(fr, text=f"{icone}  {texte}", bg=bg, fg="white",
             font=("Helvetica",13,"bold")).pack(side="left", padx=18)
    return fr

# ═══════════════════════════════════════════════════════════════════════════════
#  DIALOGUES RÉUTILISABLES (inchangés)
# ═══════════════════════════════════════════════════════════════════════════════
def champ_label(fr, label, row, widget):
    """Place un label et un widget sur la même ligne."""
    tk.Label(fr,text=label,bg=C["fond"],fg=C["texte"],
             font=("Helvetica",9,"bold")).grid(row=row,column=0,sticky="w",padx=12,pady=4)
    widget.grid(row=row,column=1,sticky="ew",padx=12,pady=4)
    return widget

def entry(parent, val=""):
    """Crée un champ de saisie avec une valeur par défaut."""
    e = tk.Entry(parent, font=("Helvetica",10), relief="solid", bd=1)
    if val: e.insert(0, val)
    return e

def combo(parent, values, val=""):
    """Crée une liste déroulante."""
    cb = ttk.Combobox(parent, values=values, state="readonly", font=("Helvetica",10))
    if val and val in values: cb.set(val)
    elif values: cb.current(0)
    return cb

# ═══════════════════════════════════════════════════════════════════════════════
#  CLASSE PARENTE POUR LES ONGLETS (factorisation)
# ═══════════════════════════════════════════════════════════════════════════════
class BaseOnglet(tk.Frame):
    """Classe de base pour tous les onglets. Fournit des méthodes communes."""
    def __init__(self, parent, db, **kwargs):
        super().__init__(parent, bg=C["fond"], **kwargs)
        self.db = db
        self._map = {}  # mapping entre iid et données
        self._build()   # à implémenter par les sous-classes
        self.refresh()  # à implémenter

    def _build(self):
        """Construit l'interface de l'onglet. À surcharger."""
        raise NotImplementedError

    def refresh(self):
        """Rafraîchit les données de l'onglet. À surcharger."""
        raise NotImplementedError

    def _sel(self):
        """Retourne l'élément sélectionné dans le Treeview principal."""
        # Si la classe a un attribut _tree, on l'utilise
        if hasattr(self, '_tree') and self._tree.selection():
            iid = self._tree.selection()[0]
            return self._map.get(iid)
        return None

    def _notify_change(self):
        """Appelée après une modification pour rafraîchir et déclencher un callback."""
        self.refresh()
        # Si un callback est défini dans l'onglet, on l'appelle
        if hasattr(self, 'on_change') and self.on_change:
            self.on_change()

# ═══════════════════════════════════════════════════════════════════════════════
#  ONGLET DASHBOARD (optimisation alertes)
# ═══════════════════════════════════════════════════════════════════════════════
class OngletDashboard(BaseOnglet):
    """Tableau de bord avec KPI et performances des livreurs."""
    def __init__(self, parent, db):
        super().__init__(parent, db)

    def _build(self):
        bt = barre_titre(self, "🏠", "Tableau de bord")
        self._lbl_date = tk.Label(bt,bg=C["primaire"],fg="#aed6f1",font=("Helvetica",9))
        self._lbl_date.pack(side="right",padx=18)

        self._fr_alerte = tk.Frame(self,bg=C["alerte_bg"],pady=6)
        self._lbl_alerte = tk.Label(self._fr_alerte,bg=C["alerte_bg"],fg=C["alerte_fg"],
                                     font=("Helvetica",9,"bold"))
        self._lbl_alerte.pack(side="left",padx=16)
        Btn(self._fr_alerte,text="Voir →",bg=C["alerte_fg"],
            command=self._voir_alertes,padx=8,pady=2).pack(side="right",padx=12)

        fr_kpi = tk.Frame(self,bg=C["fond"])
        fr_kpi.pack(fill="x",padx=20,pady=14)
        self._c = {}
        for titre, cle, col in [
            ("BONS DU JOUR","total",   C["primaire"]),
            ("LIVRÉS",      "livres",  C["succes"]),
            ("ÉCHECS",      "echecs",  C["danger"]),
            ("EN ATTENTE",  "attente", C["accent"]),
            ("TAUX RÉUSSITE","taux",   C["secondaire"]),
            ("⚠️ RETARDS",  "retards", "#c0392b"),
        ]:
            c = Carte(fr_kpi,titre,"—",col)
            c.pack(side="left",expand=True,fill="both",padx=5,pady=2)
            self._c[cle] = c

        tk.Label(self,text="Performance livreurs — aujourd'hui",bg=C["fond"],fg=C["texte"],
                 font=("Helvetica",11,"bold")).pack(anchor="w",padx=22,pady=(6,2))
        fr = tk.Frame(self,bg=C["fond"])
        fr.pack(fill="both",expand=True,padx=20,pady=(0,10))
        cols = ("Livreur","Assignés","Livrés","Échecs","Taux")
        self._tree = ttk.Treeview(fr,columns=cols,show="headings",height=7)
        for c in cols:
            self._tree.heading(c,text=c)
            self._tree.column(c,width=130,anchor="center")
        self._tree.column("Livreur",width=200,anchor="w")
        sb = ttk.Scrollbar(fr,orient="vertical",command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.pack(side="left",fill="both",expand=True)
        sb.pack(side="right",fill="y")
        self._tree.tag_configure("bien",background="#d5f5e3")
        self._tree.tag_configure("moyen",background="#fadbd8")

        Btn(self,text="🔄  Actualiser",command=self.refresh).pack(anchor="e",padx=22,pady=8)

    def _voir_alertes(self):
        retards = self.db.get_bons_retard()
        if not retards:
            messagebox.showinfo("Alertes","Aucun retard en ce moment.")
            return
        dlg = tk.Toplevel(self)
        dlg.title("🚨 Bons en retard")
        dlg.geometry("700x350")
        dlg.configure(bg=C["fond"])
        dlg.grab_set()
        tk.Label(dlg,text=f"🚨  {len(retards)} bon(s) dépassé(s) non livrés",
                 bg=C["alerte_bg"],fg=C["alerte_fg"],
                 font=("Helvetica",11,"bold")).pack(fill="x",pady=6)
        cols = ("N° Bon","Client","Livreur","Heure prévue","Observations")
        tree = ttk.Treeview(dlg,columns=cols,show="headings",height=10)
        for c in cols: tree.heading(c,text=c); tree.column(c,width=130)
        tree.column("N° Bon",width=140); tree.column("Client",width=180)
        for b in retards:
            tree.insert("","end",values=(
                b["numero"],str(b["client_nom"] or "—")[:24],
                str(b["livreur_nom"] or "—")[:18],
                b["heure_prevue"] or "—", b["observations"] or ""))
        tree.pack(fill="both",expand=True,padx=12,pady=8)
        Btn(dlg,text="Fermer",command=dlg.destroy).pack(pady=8)

    def refresh(self):
        """Rafraîchissement complet du dashboard."""
        aujourd = date.today()
        self._lbl_date.config(text=aujourd.strftime("%A %d %B %Y").capitalize())

        stats  = self.db.get_stats_jour()
        total  = stats.get("total",0) or 0
        livres = stats.get("livres",0) or 0
        taux   = f"{round(livres/total*100)}%" if total else "—%"
        retards = len(self.db.get_bons_retard())

        self._c["total"].update(total)
        self._c["livres"].update(livres)
        self._c["echecs"].update(stats.get("echecs",0) or 0)
        self._c["attente"].update(stats.get("attente",0) or 0)
        self._c["taux"].update(taux)
        self._c["retards"].update(retards)

        if retards > 0:
            self._lbl_alerte.config(text=f"⚠️  {retards} bon(s) en retard non livrés !")
            self._fr_alerte.pack(fill="x",after=self.winfo_children()[0])
        else:
            self._fr_alerte.pack_forget()

        for row in self._tree.get_children(): self._tree.delete(row)
        for r in self.db.get_stats_livreurs_jour():
            tx = f"{round(r['livres']/r['total']*100)}%" if r["total"] else "0%"
            tag = "bien" if r["total"] and r["livres"]/r["total"]>=0.8 else "moyen"
            self._tree.insert("","end",values=(r["livreur"],r["total"],r["livres"],r["echecs"],tx),tags=(tag,))

    def update_alertes(self):
        """Met à jour uniquement les indicateurs d'alerte (retards) sans recharger tout."""
        retards = len(self.db.get_bons_retard())
        self._c["retards"].update(retards)
        if retards > 0:
            self._lbl_alerte.config(text=f"⚠️  {retards} bon(s) en retard non livrés !")
            if not self._fr_alerte.winfo_ismapped():
                self._fr_alerte.pack(fill="x",after=self.winfo_children()[0])
        else:
            if self._fr_alerte.winfo_ismapped():
                self._fr_alerte.pack_forget()

# ═══════════════════════════════════════════════════════════════════════════════
#  ONGLET BONS (hérite de BaseOnglet)
# ═══════════════════════════════════════════════════════════════════════════════
class DialogBon(tk.Toplevel):
    """Dialogue pour ajouter/modifier un bon."""
    def __init__(self,parent,db,bon=None,on_save=None):
        super().__init__(parent)
        self.db=db; self.bon=bon; self.on_save=on_save
        self.title("Modifier" if bon else "Nouveau bon")
        self.geometry("530x580"); self.resizable(False,False)
        self.configure(bg=C["fond"]); self.grab_set()
        self._build()
        if bon: self._remplir()

    def _build(self):
        fr=tk.Frame(self,bg=C["fond"]); fr.pack(fill="both",expand=True,padx=12,pady=10)
        fr.columnconfigure(1,weight=1)
        livreurs=self.db.get_livreurs(); clients=self.db.get_clients(); quartiers=self.db.get_quartiers()
        self._lv_ids=[l["id"] for l in livreurs]; self._lv_noms=[f"{l['nom']} {l['prenom']}" for l in livreurs]
        self._cl_ids=[c["id"] for c in clients];  self._cl_noms=[f"{c['nom']}{' — '+c['entreprise'] if c['entreprise'] else ''}" for c in clients]
        self._qt_ids=[q["id"] for q in quartiers]; self._qt_noms=[q["nom"] for q in quartiers]

        self._e_num = champ_label(fr,"N° Bon *",0,entry(fr,self.db.get_prochain_numero()))
        self._e_date= champ_label(fr,"Date *",1,entry(fr,date.today().isoformat()))
        self._cb_cl = champ_label(fr,"Client *",2,combo(fr,self._cl_noms))
        self._cb_qt = champ_label(fr,"Quartier *",3,combo(fr,self._qt_noms))
        self._e_adr = champ_label(fr,"Adresse",4,entry(fr))
        self._e_tel = champ_label(fr,"Téléphone",5,entry(fr))
        self._e_h   = champ_label(fr,"Heure prévue",6,entry(fr,"08h00"))
        self._cb_lv = champ_label(fr,"Livreur *",7,combo(fr,self._lv_noms))

        tk.Label(fr,text="Observations",bg=C["fond"],fg=C["texte"],
                 font=("Helvetica",9,"bold")).grid(row=8,column=0,sticky="nw",padx=12,pady=4)
        self._obs=tk.Text(fr,height=3,font=("Helvetica",10),relief="solid",bd=1)
        self._obs.grid(row=8,column=1,sticky="ew",padx=12,pady=4)

        fr_btn=tk.Frame(self,bg=C["fond"]); fr_btn.pack(fill="x",padx=12,pady=8)
        Btn(fr_btn,text="💾  Enregistrer",command=self._save).pack(side="right",padx=6)
        tk.Button(fr_btn,text="Annuler",command=self.destroy,bg=C["bordure"],relief="flat",padx=12,pady=6).pack(side="right")

    def _remplir(self):
        b=self.bon
        for e,v in [(self._e_num,b["numero"]),(self._e_date,b["date_bon"]),(self._e_adr,b["adresse"] or ""),
                    (self._e_tel,b["telephone"] or ""),(self._e_h,b["heure_prevue"] or "")]:
            e.delete(0,"end"); e.insert(0,v)
        if b["client_id"] and b["client_id"] in self._cl_ids: self._cb_cl.current(self._cl_ids.index(b["client_id"]))
        if b["quartier_id"] and b["quartier_id"] in self._qt_ids: self._cb_qt.current(self._qt_ids.index(b["quartier_id"]))
        if b["livreur_id"] and b["livreur_id"] in self._lv_ids: self._cb_lv.current(self._lv_ids.index(b["livreur_id"]))
        self._obs.insert("1.0", b["observations"] or "")

    def _save(self):
        num=self._e_num.get().strip(); d=self._e_date.get().strip()
        if not num or self._cb_cl.current()<0 or self._cb_qt.current()<0 or self._cb_lv.current()<0:
            messagebox.showerror("Requis","N° Bon, Client, Quartier et Livreur sont obligatoires."); return
        try: datetime.strptime(d,"%Y-%m-%d")
        except ValueError: messagebox.showerror("Date","Format invalide (AAAA-MM-JJ)"); return
        obs=self._obs.get("1.0","end").strip()
        cl_id=self._cl_ids[self._cb_cl.current()]; qt_id=self._qt_ids[self._cb_qt.current()]
        lv_id=self._lv_ids[self._cb_lv.current()]
        try:
            if self.bon:
                self.db.update_bon(self.bon["id"],lv_id,self._e_h.get().strip(),obs)
            else:
                self.db.add_bon(num,cl_id,qt_id,self._e_adr.get().strip(),self._e_tel.get().strip(),
                                self._e_h.get().strip(),lv_id,d)
        except ValueError as e:
            messagebox.showerror("Erreur", str(e))
            return
        except sqlite3.IntegrityError:
            messagebox.showerror("Doublon",f"Le numéro '{num}' existe déjà.")
            return
        if self.on_save: self.on_save()
        self.destroy()

class DialogStatut(tk.Toplevel):
    """Dialogue pour changer le statut d'un bon."""
    def __init__(self,parent,db,bon,on_save=None):
        super().__init__(parent)
        self.db=db; self.bon=bon; self.on_save=on_save
        self.title(f"Statut — {bon['numero']}")
        self.geometry("400x310"); self.resizable(False,False)
        self.configure(bg=C["fond"]); self.grab_set()
        self._build()

    def _build(self):
        fr=tk.Frame(self,bg=C["fond"]); fr.pack(fill="both",expand=True,padx=20,pady=16)
        tk.Label(fr,text=f"Bon : {self.bon['numero']}",bg=C["fond"],fg=C["primaire"],
                 font=("Helvetica",12,"bold")).pack(anchor="w",pady=(0,10))
        self._vs=tk.StringVar(value=self.bon["statut"])
        for s in STATUTS:
            tk.Radiobutton(fr,text=s,variable=self._vs,value=s,bg=C["fond"],
                           fg=S_COL.get(s,C["texte"]),font=("Helvetica",10,"bold"),
                           selectcolor=C["fond"]).pack(anchor="w",padx=10)
        tk.Label(fr,text="Observations :",bg=C["fond"],font=("Helvetica",9,"bold")).pack(anchor="w",pady=(10,4))
        self._txt=tk.Text(fr,height=3,font=("Helvetica",10),relief="solid",bd=1); self._txt.pack(fill="x")
        if self.bon["observations"]: self._txt.insert("1.0",self.bon["observations"])
        fr_btn=tk.Frame(self,bg=C["fond"]); fr_btn.pack(fill="x",padx=20,pady=8)
        Btn(fr_btn,text="✅  Valider",command=self._save).pack(side="right",padx=6)
        tk.Button(fr_btn,text="Annuler",command=self.destroy,bg=C["bordure"],relief="flat",padx=12,pady=6).pack(side="right")

    def _save(self):
        self.db.update_statut(self.bon["id"],self._vs.get(),self._txt.get("1.0","end").strip())
        if self.on_save: self.on_save()
        self.destroy()

class OngletBons(BaseOnglet):
    """Onglet de gestion des bons de livraison."""
    def __init__(self,parent,db,on_change=None):
        self.on_change = on_change
        super().__init__(parent, db)
        self._last_date = None

    def _build(self):
        bt=barre_titre(self,"📦","Gestion des Bons")
        fr_r=tk.Frame(bt,bg=C["primaire"]); fr_r.pack(side="right",padx=10)
        tk.Label(fr_r,text="Date :",bg=C["primaire"],fg="white").pack(side="left",padx=4)
        self._e_d=tk.Entry(fr_r,width=12,font=("Helvetica",9)); self._e_d.insert(0,date.today().isoformat()); self._e_d.pack(side="left",padx=4)
        Btn(fr_r,text="🔍",command=self.refresh,padx=8,pady=3).pack(side="left",padx=4)
        Btn(fr_r,text="➕  Nouveau",command=self._nouveau,pady=3).pack(side="left",padx=4)

        fr_f=tk.Frame(self,bg=C["fond"]); fr_f.pack(fill="x",padx=16,pady=5)
        tk.Label(fr_f,text="Filtrer :",bg=C["fond"],fg=C["texte"],font=("Helvetica",9)).pack(side="left")
        self._vf=tk.StringVar(value="Tous")
        for s in ["Tous"]+STATUTS:
            tk.Radiobutton(fr_f,text=s,variable=self._vf,value=s,bg=C["fond"],
                           fg=S_COL.get(s,C["secondaire"]),font=("Helvetica",9),
                           selectcolor=C["fond"],command=self.refresh).pack(side="left",padx=8)

        cols=("N° Bon","Client","Quartier","Heure","Livreur","Statut","Observations")
        fr_t=tk.Frame(self,bg=C["fond"]); fr_t.pack(fill="both",expand=True,padx=16,pady=4)
        self._tree=ttk.Treeview(fr_t,columns=cols,show="headings",height=16)
        for c,w in zip(cols,[130,180,110,70,150,90,200]):
            self._tree.heading(c,text=c)
            anc="w" if c in ("Client","Livreur","Observations") else "center"
            self._tree.column(c,width=w,anchor=anc)
        sv=ttk.Scrollbar(fr_t,orient="vertical",command=self._tree.yview)
        sh=ttk.Scrollbar(fr_t,orient="horizontal",command=self._tree.xview)
        self._tree.configure(yscrollcommand=sv.set,xscrollcommand=sh.set)
        self._tree.grid(row=0,column=0,sticky="nsew"); sv.grid(row=0,column=1,sticky="ns"); sh.grid(row=1,column=0,sticky="ew")
        fr_t.rowconfigure(0,weight=1); fr_t.columnconfigure(0,weight=1)
        for s in STATUTS: self._tree.tag_configure(s,background=S_BG[s])

        fr_a=tk.Frame(self,bg=C["fond"]); fr_a.pack(fill="x",padx=16,pady=8)
        Btn(fr_a,text="✏️  Modifier statut",command=self._statut).pack(side="left",padx=4)
        Btn(fr_a,text="📝  Éditer",command=self._editer).pack(side="left",padx=4)
        BtnDanger(fr_a,text="🗑️  Supprimer",command=self._supprimer).pack(side="left",padx=4)
        Btn(fr_a,text="🔄  Actualiser",command=self.refresh).pack(side="right",padx=4)

    def refresh(self):
        jour=self._e_d.get().strip() or date.today().isoformat()
        bons=self.db.get_bons_jour(jour)
        if self._vf.get()!="Tous": bons=[b for b in bons if b["statut"]==self._vf.get()]
        self._map={}
        for r in self._tree.get_children(): self._tree.delete(r)
        for b in bons:
            iid=self._tree.insert("","end",values=(
                b["numero"],str(b["client_nom"] or "—")[:28],b["quartier_nom"] or "—",
                b["heure_prevue"] or "—",str(b["livreur_nom"] or "—")[:22],b["statut"],b["observations"] or ""),
                tags=(b["statut"],))
            self._map[iid]=dict(b)

    def _notify(self):
        self.refresh()
        if self.on_change: self.on_change()

    def _nouveau(self): DialogBon(self,self.db,on_save=self._notify)
    def _statut(self):
        b=self._sel()
        if b: DialogStatut(self,self.db,b,on_save=self._notify)
    def _editer(self):
        b=self._sel()
        if b: DialogBon(self,self.db,bon=b,on_save=self._notify)
    def _supprimer(self):
        b=self._sel()
        if b and messagebox.askyesno("Confirmer",f"Supprimer {b['numero']} ?"):
            try:
                self.db.delete_bon(b["id"])
                self._notify()
            except ValueError as e:
                messagebox.showerror("Erreur", str(e))

# ═══════════════════════════════════════════════════════════════════════════════
#  ONGLET CLIENTS (hérite de BaseOnglet)
# ═══════════════════════════════════════════════════════════════════════════════
class OngletClients(BaseOnglet):
    """Onglet de gestion des clients."""
    def __init__(self,parent,db):
        super().__init__(parent,db)

    def _build(self):
        bt=barre_titre(self,"🏪","Gestion des Clients")
        Btn(bt,text="➕  Nouveau client",command=self._ajouter,pady=3).pack(side="right",padx=12)

        fr_rech=tk.Frame(self,bg=C["fond"]); fr_rech.pack(fill="x",padx=16,pady=6)
        tk.Label(fr_rech,text="🔍 Rechercher :",bg=C["fond"],fg=C["texte"],
                 font=("Helvetica",9)).pack(side="left")
        self._e_rech=tk.Entry(fr_rech,font=("Helvetica",10),relief="solid",bd=1,width=24)
        self._e_rech.pack(side="left",padx=8)
        self._e_rech.bind("<KeyRelease>",lambda e: self.refresh())

        cols=("Civilité","Nom","Entreprise","Quartier","Téléphone","Email")
        fr=tk.Frame(self,bg=C["fond"]); fr.pack(fill="both",expand=True,padx=16,pady=4)
        self._tree=ttk.Treeview(fr,columns=cols,show="headings",height=16)
        for c,w in zip(cols,[60,160,200,120,110,160]):
            self._tree.heading(c,text=c); self._tree.column(c,width=w)
        sb=ttk.Scrollbar(fr,orient="vertical",command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.pack(side="left",fill="both",expand=True); sb.pack(side="right",fill="y")

        fr_a=tk.Frame(self,bg=C["fond"]); fr_a.pack(fill="x",padx=16,pady=8)
        Btn(fr_a,text="✏️  Modifier",command=self._modifier).pack(side="left",padx=4)
        BtnDanger(fr_a,text="🗑️  Supprimer",command=self._supprimer).pack(side="left",padx=4)
        Btn(fr_a,text="📊  Historique livraisons",command=self._historique).pack(side="left",padx=4)
        Btn(fr_a,text="🔄  Actualiser",command=self.refresh).pack(side="right",padx=4)

    def refresh(self):
        rech=(self._e_rech.get().strip().lower() if hasattr(self,"_e_rech") else "")
        self._map={}
        for r in self._tree.get_children(): self._tree.delete(r)
        for c in self.db.get_clients():
            if rech and rech not in (c["nom"]+" "+(c["entreprise"] or "")).lower(): continue
            iid=self._tree.insert("","end",values=(
                c["civilite"],c["nom"],c["entreprise"] or "",c["quartier"] or "",
                c["telephone"] or "",c["email"] or ""))
            self._map[iid]=dict(c)

    def _dialog(self, cl=None):
        dlg=tk.Toplevel(self); dlg.title("Modifier client" if cl else "Nouveau client")
        dlg.geometry("440x380"); dlg.configure(bg=C["fond"]); dlg.grab_set()
        fr=tk.Frame(dlg,bg=C["fond"]); fr.pack(fill="both",expand=True,padx=16,pady=12)
        fr.columnconfigure(1,weight=1)

        civs=["M.","Mme","Dr","Prof."]
        e_civ=champ_label(fr,"Civilité",0,combo(fr,civs,cl["civilite"] if cl else "M."))
        e_nom=champ_label(fr,"Nom *",1,entry(fr,cl["nom"] if cl else ""))
        e_ent=champ_label(fr,"Entreprise",2,entry(fr,cl["entreprise"] if cl and cl["entreprise"] else ""))
        e_qt =champ_label(fr,"Quartier",3,entry(fr,cl["quartier"] if cl and cl["quartier"] else ""))
        e_adr=champ_label(fr,"Adresse",4,entry(fr,cl["adresse"] if cl and cl["adresse"] else ""))
        e_tel=champ_label(fr,"Téléphone",5,entry(fr,cl["telephone"] if cl and cl["telephone"] else ""))
        e_em =champ_label(fr,"Email",6,entry(fr,cl["email"] if cl and cl.get("email") else ""))

        def save():
            nom=e_nom.get().strip()
            if not nom: messagebox.showerror("Requis","Le nom est obligatoire."); return
            try:
                if cl:
                    self.db.update_client(cl["id"],e_civ.get(),nom,e_ent.get().strip(),
                                          e_qt.get().strip(),e_adr.get().strip(),
                                          e_tel.get().strip(),e_em.get().strip())
                else:
                    self.db.add_client(e_civ.get(),nom,e_ent.get().strip(),
                                       e_qt.get().strip(),e_adr.get().strip(),
                                       e_tel.get().strip(),e_em.get().strip())
                self.refresh()
                dlg.destroy()
            except ValueError as e:
                messagebox.showerror("Erreur", str(e))

        fr_b=tk.Frame(dlg,bg=C["fond"]); fr_b.pack(fill="x",padx=16,pady=8)
        Btn(fr_b,text="💾  Enregistrer",command=save).pack(side="right",padx=4)
        tk.Button(fr_b,text="Annuler",command=dlg.destroy,bg=C["bordure"],relief="flat",padx=12,pady=6).pack(side="right")

    def _ajouter(self): self._dialog()
    def _modifier(self):
        c=self._sel()
        if c: self._dialog(c)

    def _supprimer(self):
        c=self._sel()
        if c and messagebox.askyesno("Confirmer",f"Supprimer {c['nom']} et tous ses bons ?"):
            try:
                self.db.delete_client(c["id"])
                self.refresh()
            except ValueError as e:
                messagebox.showerror("Erreur", str(e))

    def _historique(self):
        cl=self._sel()
        if not cl: return
        dlg=tk.Toplevel(self); dlg.title(f"Historique — {cl['nom']}")
        dlg.geometry("680x400"); dlg.configure(bg=C["fond"]); dlg.grab_set()
        tk.Label(dlg,text=f"Livraisons de {cl['civilite']} {cl['nom']}",bg=C["fond"],
                 fg=C["primaire"],font=("Helvetica",12,"bold")).pack(pady=8)
        bons=self.db._bons_query("WHERE b.client_id=? ORDER BY b.date_bon DESC",(cl["id"],))
        cols=("Date","N° Bon","Quartier","Livreur","Statut")
        tree=ttk.Treeview(dlg,columns=cols,show="headings",height=14)
        for c_,w in zip(cols,[100,140,110,160,90]):
            tree.heading(c_,text=c_); tree.column(c_,width=w)
        for s in STATUTS: tree.tag_configure(s,background=S_BG[s])
        for b in bons:
            tree.insert("","end",values=(b["date_bon"],b["numero"],b["quartier_nom"] or "—",
                                          b["livreur_nom"] or "—",b["statut"]),tags=(b["statut"],))
        tree.pack(fill="both",expand=True,padx=12,pady=4)
        Btn(dlg,text="Fermer",command=dlg.destroy).pack(pady=8)

# ═══════════════════════════════════════════════════════════════════════════════
#  ONGLET INCIDENTS (hérite de BaseOnglet)
# ═══════════════════════════════════════════════════════════════════════════════
TYPES_INC = ["Adresse incorrecte","Client absent","Colis abîmé","Retard",
             "Livreur en panne","Refus de livraison","Autre"]

class OngletIncidents(BaseOnglet):
    """Onglet de gestion des incidents."""
    def __init__(self,parent,db,on_change=None):
        self.on_change = on_change
        super().__init__(parent, db)

    def _build(self):
        bt=barre_titre(self,"🚨","Historique des Incidents")
        Btn(bt,text="➕  Signaler incident",command=self._ajouter,pady=3).pack(side="right",padx=12)

        fr_f=tk.Frame(self,bg=C["fond"]); fr_f.pack(fill="x",padx=16,pady=6)
        tk.Label(fr_f,text="Afficher :",bg=C["fond"],fg=C["texte"],font=("Helvetica",9)).pack(side="left")
        self._vf=tk.StringVar(value="Tous")
        for label,val in [("Tous","Tous"),("En cours","En cours"),("Résolus","Résolus")]:
            tk.Radiobutton(fr_f,text=label,variable=self._vf,value=val,bg=C["fond"],
                           font=("Helvetica",9),selectcolor=C["fond"],
                           command=self.refresh).pack(side="left",padx=10)

        cols=("Date","Bon N°","Type","Description","Livreur","Résolu","Résolution")
        fr=tk.Frame(self,bg=C["fond"]); fr.pack(fill="both",expand=True,padx=16,pady=4)
        self._tree=ttk.Treeview(fr,columns=cols,show="headings",height=16)
        for c,w in zip(cols,[130,110,130,220,150,65,200]):
            self._tree.heading(c,text=c); self._tree.column(c,width=w)
        sb=ttk.Scrollbar(fr,orient="vertical",command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.pack(side="left",fill="both",expand=True); sb.pack(side="right",fill="y")
        self._tree.tag_configure("resolu",background="#d5f5e3")
        self._tree.tag_configure("ouvert",background="#fadbd8")

        fr_a=tk.Frame(self,bg=C["fond"]); fr_a.pack(fill="x",padx=16,pady=8)
        Btn(fr_a,text="✏️  Modifier/Résoudre",command=self._modifier).pack(side="left",padx=4)
        BtnDanger(fr_a,text="🗑️  Supprimer",command=self._supprimer).pack(side="left",padx=4)
        Btn(fr_a,text="🔄  Actualiser",command=self.refresh).pack(side="right",padx=4)

    def refresh(self):
        filtre=self._vf.get()
        resolu=None if filtre=="Tous" else (1 if filtre=="Résolus" else 0)
        self._map={}
        for r in self._tree.get_children(): self._tree.delete(r)
        for i in self.db.get_incidents(resolu):
            d=i["date_inc"][:16] if i["date_inc"] else "—"
            res="✅ Oui" if i["resolu"] else "❌ Non"
            tag="resolu" if i["resolu"] else "ouvert"
            iid=self._tree.insert("","end",values=(
                d, i["bon_numero"] or "—", i["type_inc"],
                str(i["description"])[:40], i["livreur_nom"] or "—",
                res, str(i["resolution"] or "")[:40]), tags=(tag,))
            self._map[iid]=dict(i)
        if self.on_change: self.on_change()

    def _dialog(self, inc=None):
        dlg=tk.Toplevel(self); dlg.title("Modifier incident" if inc else "Signaler un incident")
        dlg.geometry("500x440"); dlg.configure(bg=C["fond"]); dlg.grab_set()
        fr=tk.Frame(dlg,bg=C["fond"]); fr.pack(fill="both",expand=True,padx=16,pady=12)
        fr.columnconfigure(1,weight=1)

        bons_jour=self.db.get_bons_jour()
        bon_noms=["(Aucun bon lié)"]+[b["numero"] for b in bons_jour]
        bon_ids=[None]+[b["id"] for b in bons_jour]

        e_bon=champ_label(fr,"Bon lié (optionnel)",0,combo(fr,bon_noms))
        if inc and inc.get("bon_id") and inc["bon_id"] in bon_ids:
            e_bon.current(bon_ids.index(inc["bon_id"]))

        e_type=champ_label(fr,"Type *",1,combo(fr,TYPES_INC,inc["type_inc"] if inc else ""))

        tk.Label(fr,text="Description *",bg=C["fond"],fg=C["texte"],
                 font=("Helvetica",9,"bold")).grid(row=2,column=0,sticky="nw",padx=12,pady=4)
        e_desc=tk.Text(fr,height=4,font=("Helvetica",10),relief="solid",bd=1)
        e_desc.grid(row=2,column=1,sticky="ew",padx=12,pady=4)
        if inc: e_desc.insert("1.0",inc["description"])

        tk.Label(fr,text="Résolution",bg=C["fond"],fg=C["texte"],
                 font=("Helvetica",9,"bold")).grid(row=3,column=0,sticky="nw",padx=12,pady=4)
        e_res=tk.Text(fr,height=4,font=("Helvetica",10),relief="solid",bd=1)
        e_res.grid(row=3,column=1,sticky="ew",padx=12,pady=4)
        if inc: e_res.insert("1.0",inc["resolution"] or "")

        v_resolu=tk.BooleanVar(value=bool(inc["resolu"]) if inc else False)
        tk.Checkbutton(fr,text="Incident résolu ✅",variable=v_resolu,bg=C["fond"],
                       font=("Helvetica",10)).grid(row=4,column=1,sticky="w",padx=12,pady=6)

        def save():
            desc=e_desc.get("1.0","end").strip()
            if not desc: messagebox.showerror("Requis","La description est obligatoire."); return
            bon_id=bon_ids[e_bon.current()]
            res_txt=e_res.get("1.0","end").strip()
            try:
                if inc:
                    self.db.update_incident(inc["id"],e_type.get(),desc,res_txt,int(v_resolu.get()))
                else:
                    self.db.add_incident(bon_id,e_type.get(),desc,res_txt,int(v_resolu.get()))
                self.refresh()
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Erreur", str(e))

        fr_b=tk.Frame(dlg,bg=C["fond"]); fr_b.pack(fill="x",padx=16,pady=8)
        Btn(fr_b,text="💾  Enregistrer",command=save).pack(side="right",padx=4)
        tk.Button(fr_b,text="Annuler",command=dlg.destroy,bg=C["bordure"],relief="flat",padx=12,pady=6).pack(side="right")

    def _ajouter(self): self._dialog()
    def _modifier(self):
        i=self._sel()
        if i: self._dialog(i)
    def _supprimer(self):
        i=self._sel()
        if i and messagebox.askyesno("Confirmer","Supprimer cet incident ?"):
            try:
                self.db.delete_incident(i["id"])
                self.refresh()
            except Exception as e:
                messagebox.showerror("Erreur", str(e))

# ═══════════════════════════════════════════════════════════════════════════════
#  ONGLET STATISTIQUES (hérite de BaseOnglet)
# ═══════════════════════════════════════════════════════════════════════════════
class BarChart(tk.Canvas):
    """Diagramme en barres simple."""
    def __init__(self,parent,data,couleurs,**kw):
        kw.setdefault("bg",C["fond"]); kw.setdefault("width",580); kw.setdefault("height",220)
        super().__init__(parent,**kw)
        self._data=data; self._cols=couleurs
        self.bind("<Configure>",lambda e: self._draw())
        self._draw()

    def update_data(self,data,couleurs=None):
        self._data=data
        if couleurs: self._cols=couleurs
        self._draw()

    def _draw(self):
        self.delete("all")
        W=self.winfo_width() or 580; H=self.winfo_height() or 220
        pad_l=40; pad_b=40; pad_t=20; pad_r=20
        w=W-pad_l-pad_r; h=H-pad_b-pad_t
        if not self._data: return
        labels=[str(d[0]) for d in self._data]
        values=[max(0,float(d[1])) for d in self._data]
        max_v=max(values) if values else 1
        n=len(values); bw=max(12,w//n-8); gap=max(4,(w-(bw*n))//(n+1))

        self.create_line(pad_l,pad_t,pad_l,H-pad_b,fill="#aaa",width=1)
        self.create_line(pad_l,H-pad_b,W-pad_r,H-pad_b,fill="#aaa",width=1)

        for i in range(1,5):
            y=pad_t+h*(1-i/4); v=round(max_v*i/4)
            self.create_line(pad_l,y,W-pad_r,y,fill="#e0e0e0",dash=(3,3))
            self.create_text(pad_l-4,y,text=str(v),anchor="e",font=("Helvetica",7),fill="#666")

        cols=self._cols
        for i,(lbl,val) in enumerate(zip(labels,values)):
            x=pad_l+gap*(i+1)+bw*i
            bar_h=(val/max_v)*h if max_v else 0
            y1=H-pad_b-bar_h; y2=H-pad_b
            col=cols[i%len(cols)] if cols else C["secondaire"]
            self.create_rectangle(x,y1,x+bw,y2,fill=col,outline="white",width=1)
            if val>0:
                self.create_text(x+bw//2,y1-6,text=str(int(val)),font=("Helvetica",7,"bold"),fill=C["texte"])
            lbl_s=lbl[:10] if len(lbl)>10 else lbl
            self.create_text(x+bw//2,H-pad_b+10,text=lbl_s,font=("Helvetica",7),fill=C["texte"],angle=0)

class OngletStatistiques(BaseOnglet):
    """Onglet de statistiques et comparaisons."""
    def __init__(self,parent,db):
        super().__init__(parent,db)

    def _build(self):
        barre_titre(self,"📊","Statistiques & Comparaisons")

        fr_s=tk.Frame(self,bg=C["fond"]); fr_s.pack(fill="x",padx=16,pady=8)
        annee_act=date.today().year; mois_act=date.today().month
        annees=[str(y) for y in range(annee_act-2,annee_act+1)]

        tk.Label(fr_s,text="Mois A :",bg=C["fond"],fg=C["texte"],font=("Helvetica",9,"bold")).pack(side="left")
        self._cb_a1=combo(fr_s,[f"{MOIS_FR[m]} {y}" for y in reversed(range(annee_act-1,annee_act+1)) for m in range(1,13)])
        self._cb_a1.pack(side="left",padx=6)
        idx1=next((i for i,v in enumerate(self._cb_a1["values"]) if MOIS_FR[mois_act] in v and str(annee_act) in v),0)
        self._cb_a1.current(idx1)

        tk.Label(fr_s,text="Mois B :",bg=C["fond"],fg=C["texte"],font=("Helvetica",9,"bold")).pack(side="left",padx=(20,0))
        self._cb_a2=combo(fr_s,[f"{MOIS_FR[m]} {y}" for y in reversed(range(annee_act-1,annee_act+1)) for m in range(1,13)])
        self._cb_a2.pack(side="left",padx=6)
        mois_prec=mois_act-1 or 12; annee_prec=annee_act if mois_act>1 else annee_act-1
        idx2=next((i for i,v in enumerate(self._cb_a2["values"]) if MOIS_FR[mois_prec] in v and str(annee_prec) in v),0)
        self._cb_a2.current(idx2)

        Btn(fr_s,text="📊  Comparer",command=self._comparer).pack(side="left",padx=16)
        Btn(fr_s,text="📎  Exporter Excel",command=self._export_excel).pack(side="left",padx=8)

        self._nb=ttk.Notebook(self)
        self._nb.pack(fill="both",expand=True,padx=16,pady=8)

        self._fr_comp=tk.Frame(self._nb,bg=C["fond"])
        self._fr_lv  =tk.Frame(self._nb,bg=C["fond"])
        self._fr_qt  =tk.Frame(self._nb,bg=C["fond"])
        self._nb.add(self._fr_comp,text="  📈 Comparaison globale  ")
        self._nb.add(self._fr_lv,  text="  👥 Par livreur  ")
        self._nb.add(self._fr_qt,  text="  🗺️ Par quartier  ")

        # Pas d'appel à _comparer ici car refresh() le fera
        # self._comparer()  # <-- supprimé pour éviter double exécution

    def refresh(self):
        """Rafraîchit les statistiques."""
        self._comparer()

    def _parse_mois(self, val):
        parts=val.split(); mois=MOIS_FR.index(parts[0]); annee=int(parts[1])
        return annee, mois

    def _comparer(self):
        a1,m1=self._parse_mois(self._cb_a1.get())
        a2,m2=self._parse_mois(self._cb_a2.get())
        s1=self.db.get_stats_mois(a1,m1); s2=self.db.get_stats_mois(a2,m2)
        lbl1=f"{MOIS_FR[m1]} {a1}"; lbl2=f"{MOIS_FR[m2]} {a2}"

        self._build_comp(s1,s2,lbl1,lbl2)
        self._build_livreurs(a1,m1,a2,m2,lbl1,lbl2)
        self._build_quartiers(a1,m1,a2,m2,lbl1,lbl2)

    def _build_comp(self,s1,s2,lbl1,lbl2):
        for w in self._fr_comp.winfo_children(): w.destroy()
        t1=s1.get("total",0) or 0; l1=s1.get("livres",0) or 0
        t2=s2.get("total",0) or 0; l2=s2.get("livres",0) or 0
        tx1=round(l1/t1*100) if t1 else 0; tx2=round(l2/t2*100) if t2 else 0

        tk.Label(self._fr_comp,text="Synthèse comparative",bg=C["fond"],fg=C["primaire"],
                 font=("Helvetica",11,"bold")).pack(pady=(12,4))
        fr_t=tk.Frame(self._fr_comp,bg=C["fond"]); fr_t.pack(fill="x",padx=16,pady=4)
        cols=("Indicateur",lbl1,lbl2,"Évolution")
        tree=ttk.Treeview(fr_t,columns=cols,show="headings",height=6)
        for c,w in zip(cols,[200,160,160,120]):
            tree.heading(c,text=c); tree.column(c,width=w,anchor="center")
        tree.column("Indicateur",anchor="w")
        rows=[("Bons traités",str(t1),str(t2),self._evo(t1,t2)),
              ("Livrés",str(l1),str(l2),self._evo(l1,l2)),
              ("Échecs",str(s1.get("echecs",0) or 0),str(s2.get("echecs",0) or 0),""),
              ("Reportés",str(s1.get("reportes",0) or 0),str(s2.get("reportes",0) or 0),""),
              ("Taux réussite",f"{tx1}%",f"{tx2}%",self._evo(tx1,tx2,pct=True))]
        for r in rows: tree.insert("","end",values=r)
        tree.pack(fill="x")

        tk.Label(self._fr_comp,text="Bons traités & livrés par mois",bg=C["fond"],fg=C["texte"],
                 font=("Helvetica",10,"bold")).pack(pady=(14,4))
        chart_data=[(lbl1+" traités",t1),(lbl1+" livrés",l1),
                    (lbl2+" traités",t2),(lbl2+" livrés",l2)]
        chart=BarChart(self._fr_comp,chart_data,[C["secondaire"],C["succes"],C["accent"],C["violet"]],height=200)
        chart.pack(fill="x",padx=16,pady=4)

    def _build_livreurs(self,a1,m1,a2,m2,lbl1,lbl2):
        for w in self._fr_lv.winfo_children(): w.destroy()
        lv1=self.db.get_stats_livreurs_mois(a1,m1)
        lv2=self.db.get_stats_livreurs_mois(a2,m2)
        lv2_dict={r["livreur_id"]:r for r in lv2}

        fr=tk.Frame(self._fr_lv,bg=C["fond"]); fr.pack(fill="both",expand=True,padx=16,pady=12)
        cols=("Livreur",f"Bons {lbl1}",f"Livrés {lbl1}",f"Taux {lbl1}",
              f"Bons {lbl2}",f"Livrés {lbl2}",f"Taux {lbl2}","Évolution")
        tree=ttk.Treeview(fr,columns=cols,show="headings",height=12)
        for c in cols: tree.heading(c,text=c); tree.column(c,width=90,anchor="center")
        tree.column("Livreur",width=160,anchor="w")
        tree.column(f"Bons {lbl1}",width=80); tree.column(f"Bons {lbl2}",width=80)

        for r1 in lv1:
            r2=lv2_dict.get(r1["livreur_id"])
            t1=r1["total"] or 0; l1=r1["livres"] or 0; tx1=round(l1/t1*100) if t1 else 0
            t2=r2["total"] if r2 else 0; l2=r2["livres"] if r2 else 0
            tx2=round(l2/t2*100) if t2 else 0
            evo=self._evo(tx1,tx2,pct=True)
            tree.insert("","end",values=(r1["livreur"],t1,l1,f"{tx1}%",t2,l2,f"{tx2}%",evo))
        sb=ttk.Scrollbar(fr,orient="vertical",command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left",fill="both",expand=True); sb.pack(side="right",fill="y")

        tk.Label(self._fr_lv,text=f"Taux de réussite — {lbl1}",bg=C["fond"],fg=C["texte"],
                 font=("Helvetica",10,"bold")).pack(pady=(4,2))
        chart_data=[(r["livreur"].split()[0],round(r["livres"]/r["total"]*100) if r["total"] else 0)
                    for r in lv1]
        BarChart(self._fr_lv,chart_data,[C["succes"],C["secondaire"],C["accent"],C["violet"],C["danger"]],
                 height=180).pack(fill="x",padx=16,pady=4)

    def _build_quartiers(self,a1,m1,a2,m2,lbl1,lbl2):
        for w in self._fr_qt.winfo_children(): w.destroy()
        qt1=self.db.get_stats_quartiers_mois(a1,m1)
        qt2=self.db.get_stats_quartiers_mois(a2,m2)
        qt2_dict={r["quartier"]:r for r in qt2}

        fr=tk.Frame(self._fr_qt,bg=C["fond"]); fr.pack(fill="x",padx=16,pady=12)
        cols=("Quartier",f"Bons {lbl1}",f"Échecs {lbl1}",f"Bons {lbl2}",f"Échecs {lbl2}")
        tree=ttk.Treeview(fr,columns=cols,show="headings",height=10)
        for c in cols: tree.heading(c,text=c); tree.column(c,width=120,anchor="center")
        tree.column("Quartier",width=160,anchor="w")
        for r1 in qt1:
            r2=qt2_dict.get(r1["quartier"])
            tree.insert("","end",values=(r1["quartier"],r1["total"],r1["echecs"],
                                          r2["total"] if r2 else 0, r2["echecs"] if r2 else 0))
        sb=ttk.Scrollbar(fr,orient="vertical",command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left",fill="both",expand=True); sb.pack(side="right",fill="y")

        tk.Label(self._fr_qt,text=f"Échecs par quartier — {lbl1}",bg=C["fond"],fg=C["texte"],
                 font=("Helvetica",10,"bold")).pack(pady=(8,2))
        cd=[(r["quartier"],r["echecs"]) for r in qt1 if r["echecs"]>0]
        if cd:
            BarChart(self._fr_qt,cd,[C["danger"]],height=180).pack(fill="x",padx=16,pady=4)
        else:
            tk.Label(self._fr_qt,text="Aucun échec ce mois 🎉",bg=C["fond"],
                     fg=C["succes"],font=("Helvetica",11)).pack(pady=20)

    def _evo(self,v1,v2,pct=False):
        if not v2: return "—"
        diff=v1-v2
        if diff>0: sym=f"▲ +{diff}" + ("%" if pct else "")
        elif diff<0: sym=f"▼ {diff}" + ("%" if pct else "")
        else: sym="= 0"
        return sym

    def _export_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Erreur", "openpyxl n'est pas installé.\nInstallez-le avec : pip install openpyxl")
            return
        try:
            a1,m1=self._parse_mois(self._cb_a1.get())
            a2,m2=self._parse_mois(self._cb_a2.get())
            s1=self.db.get_stats_mois(a1,m1); s2=self.db.get_stats_mois(a2,m2)
            lv1=self.db.get_stats_livreurs_mois(a1,m1)
            lv2=self.db.get_stats_livreurs_mois(a2,m2)
            qt1=self.db.get_stats_quartiers_mois(a1,m1)
            qt2=self.db.get_stats_quartiers_mois(a2,m2)

            wb = Workbook()
            # Feuille synthèse
            ws = wb.active
            ws.title = "Synthèse"
            ws.append(["Indicateur", f"{MOIS_FR[m1]} {a1}", f"{MOIS_FR[m2]} {a2}", "Évolution"])
            ws.append(["Bons traités", s1.get("total",0), s2.get("total",0), self._evo(s1.get("total",0), s2.get("total",0))])
            ws.append(["Livrés", s1.get("livres",0), s2.get("livres",0), self._evo(s1.get("livres",0), s2.get("livres",0))])
            ws.append(["Échecs", s1.get("echecs",0), s2.get("echecs",0), self._evo(s1.get("echecs",0), s2.get("echecs",0))])
            ws.append(["Reportés", s1.get("reportes",0), s2.get("reportes",0), self._evo(s1.get("reportes",0), s2.get("reportes",0))])
            ws.append(["Taux réussite", f"{round(s1.get('livres',0)/(s1.get('total',0) or 1)*100)}%",
                       f"{round(s2.get('livres',0)/(s2.get('total',0) or 1)*100)}%",
                       self._evo(round(s1.get('livres',0)/(s1.get('total',0) or 1)*100),
                                 round(s2.get('livres',0)/(s2.get('total',0) or 1)*100), pct=True)])

            # Mise en forme
            for col in range(1,5):
                ws.column_dimensions[get_column_letter(col)].width = 15
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="1a3c5e", end_color="1a3c5e", fill_type="solid")
                cell.alignment = Alignment(horizontal="center")
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    cell.alignment = Alignment(horizontal="center")

            # Feuille livreurs
            ws_lv = wb.create_sheet("Livreurs")
            ws_lv.append(["Livreur", f"Bons {MOIS_FR[m1]}", f"Livrés {MOIS_FR[m1]}", f"Taux {MOIS_FR[m1]}",
                          f"Bons {MOIS_FR[m2]}", f"Livrés {MOIS_FR[m2]}", f"Taux {MOIS_FR[m2]}", "Évolution"])
            lv2_dict = {r["livreur_id"]:r for r in lv2}
            for r1 in lv1:
                r2=lv2_dict.get(r1["livreur_id"])
                t1=r1["total"] or 0; l1=r1["livres"] or 0; tx1=round(l1/t1*100) if t1 else 0
                t2=r2["total"] if r2 else 0; l2=r2["livres"] if r2 else 0
                tx2=round(l2/t2*100) if t2 else 0
                evo=self._evo(tx1,tx2,pct=True)
                ws_lv.append([r1["livreur"], t1, l1, f"{tx1}%", t2, l2, f"{tx2}%", evo])

            # Feuille quartiers
            ws_qt = wb.create_sheet("Quartiers")
            ws_qt.append(["Quartier", f"Bons {MOIS_FR[m1]}", f"Échecs {MOIS_FR[m1]}",
                          f"Bons {MOIS_FR[m2]}", f"Échecs {MOIS_FR[m2]}"])
            qt2_dict = {r["quartier"]:r for r in qt2}
            for r1 in qt1:
                r2=qt2_dict.get(r1["quartier"])
                ws_qt.append([r1["quartier"], r1["total"], r1["echecs"],
                              r2["total"] if r2 else 0, r2["echecs"] if r2 else 0])

            # Sauvegarde
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
                                                initialfile=f"Statistiques_{MOIS_FR[m1]}_{a1}_vs_{MOIS_FR[m2]}_{a2}.xlsx")
            if path:
                wb.save(path)
                messagebox.showinfo("Export Excel", f"Export réussi vers {os.path.basename(path)}")
                log_action(f"Export Excel : {path}")
        except Exception as e:
            messagebox.showerror("Erreur export", str(e))

# ═══════════════════════════════════════════════════════════════════════════════
#  ONGLET LIVREURS (hérite de BaseOnglet)
# ═══════════════════════════════════════════════════════════════════════════════
class OngletLivreurs(BaseOnglet):
    """Onglet de gestion des livreurs."""
    def __init__(self,parent,db):
        super().__init__(parent,db)

    def _build(self):
        bt=barre_titre(self,"👥","Gestion des Livreurs")
        Btn(bt,text="➕  Ajouter",command=self._ajouter,pady=3).pack(side="right",padx=12)
        cols=("Nom","Prénom","Téléphone","Secteur","Actif")
        fr=tk.Frame(self,bg=C["fond"]); fr.pack(fill="both",expand=True,padx=16,pady=16)
        self._tree=ttk.Treeview(fr,columns=cols,show="headings",height=18)
        for c,w in zip(cols,[150,150,120,160,60]):
            self._tree.heading(c,text=c); self._tree.column(c,width=w)
        sb=ttk.Scrollbar(fr,orient="vertical",command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.pack(side="left",fill="both",expand=True); sb.pack(side="right",fill="y")
        fr_b=tk.Frame(self,bg=C["fond"]); fr_b.pack(fill="x",padx=16,pady=8)
        Btn(fr_b,text="✏️  Modifier",command=self._modifier).pack(side="left",padx=4)

    def refresh(self):
        self._map={}
        for r in self._tree.get_children(): self._tree.delete(r)
        for l in self.db.get_livreurs(actif_only=False):
            iid=self._tree.insert("","end",values=(l["nom"],l["prenom"],l["telephone"] or "",
                                                    l["secteur"] or "","✅" if l["actif"] else "❌"))
            self._map[iid]=dict(l)

    def _dialog(self, lv=None):
        dlg=tk.Toplevel(self); dlg.title("Modifier" if lv else "Nouveau livreur")
        dlg.geometry("400x340"); dlg.configure(bg=C["fond"]); dlg.grab_set()
        fr=tk.Frame(dlg,bg=C["fond"]); fr.pack(fill="both",expand=True,padx=20,pady=16)
        fr.columnconfigure(1,weight=1)
        e_nom=champ_label(fr,"Nom *",0,entry(fr,lv["nom"] if lv else ""))
        e_pre=champ_label(fr,"Prénom *",1,entry(fr,lv["prenom"] if lv else ""))
        e_tel=champ_label(fr,"Téléphone",2,entry(fr,lv["telephone"] if lv and lv["telephone"] else ""))
        e_sec=champ_label(fr,"Secteur",3,entry(fr,lv["secteur"] if lv and lv["secteur"] else ""))
        v_act=tk.BooleanVar(value=bool(lv["actif"]) if lv else True)
        tk.Checkbutton(fr,text="Livreur actif",variable=v_act,bg=C["fond"],
                       font=("Helvetica",10)).grid(row=4,column=1,sticky="w",padx=12,pady=8)

        def save():
            nom=e_nom.get().strip(); pre=e_pre.get().strip()
            if not nom or not pre: messagebox.showerror("Requis","Nom et prénom obligatoires."); return
            try:
                if lv:
                    self.db.update_livreur(lv["id"],nom,pre,e_tel.get().strip(),e_sec.get().strip(),int(v_act.get()))
                else:
                    self.db.add_livreur(nom,pre,e_tel.get().strip(),e_sec.get().strip())
                self.refresh()
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Erreur", str(e))

        fr_b=tk.Frame(dlg,bg=C["fond"]); fr_b.pack(fill="x",padx=20,pady=8)
        Btn(fr_b,text="💾  Enregistrer",command=save).pack(side="right",padx=4)
        tk.Button(fr_b,text="Annuler",command=dlg.destroy,bg=C["bordure"],relief="flat",padx=12,pady=6).pack(side="right")

    def _ajouter(self): self._dialog()
    def _modifier(self):
        s=self._tree.selection()
        if not s: messagebox.showwarning("Sélection","Sélectionnez un livreur."); return
        self._dialog(self._map[s[0]])

# ═══════════════════════════════════════════════════════════════════════════════
#  ONGLET CARTE (avec PieChart utilisant math)
# ═══════════════════════════════════════════════════════════════════════════════
import math  # import localisé pour le PieChart

class PieChart(tk.Canvas):
    """Diagramme circulaire pour visualiser les proportions."""
    def __init__(self, parent, data, colors=None, **kw):
        kw.setdefault("bg", C["fond"])
        kw.setdefault("width", 400)
        kw.setdefault("height", 400)
        super().__init__(parent, **kw)
        self._data = data
        self._colors = colors or [C["primaire"], C["secondaire"], C["accent"], C["succes"], C["danger"], C["violet"], C["neutre"]]
        self.bind("<Configure>", lambda e: self._draw())
        self._draw()

    def _draw(self):
        """Dessine le camembert en fonction des données."""
        self.delete("all")
        W = self.winfo_width()
        H = self.winfo_height()
        center_x = W//2
        center_y = H//2
        radius = min(W, H)//2 - 20
        if not self._data:
            self.create_text(center_x, center_y, text="Aucune donnée", fill=C["texte"])
            return
        total = sum(v for _,v in self._data)
        if total == 0:
            self.create_text(center_x, center_y, text="Aucune livraison", fill=C["texte"])
            return
        start = 0
        for i, (label, value) in enumerate(self._data):
            angle = (value / total) * 360
            col = self._colors[i % len(self._colors)]
            self.create_arc(center_x-radius, center_y-radius, center_x+radius, center_y+radius,
                            start=start, extent=angle, fill=col, outline="white", width=2)
            mid_angle = start + angle/2
            # Positionner le texte à 60% du rayon
            x = center_x + (radius*0.6) * math.cos(math.radians(mid_angle))
            y = center_y + (radius*0.6) * math.sin(math.radians(mid_angle))
            self.create_text(x, y, text=f"{label}\n{value}", fill="white", font=("Helvetica", 8, "bold"), justify="center")
            start += angle

class OngletCarte(BaseOnglet):
    """Onglet carte des livraisons par quartier."""
    def __init__(self, parent, db):
        super().__init__(parent, db)

    def _build(self):
        barre_titre(self, "🗺️", "Carte des livraisons par quartier")

        fr_sel = tk.Frame(self, bg=C["fond"])
        fr_sel.pack(fill="x", padx=16, pady=8)
        tk.Label(fr_sel, text="Mois :", bg=C["fond"], fg=C["texte"], font=("Helvetica",9,"bold")).pack(side="left")
        mois_noms = [f"{MOIS_FR[m]} {y}" for y in range(date.today().year-1, date.today().year+1) for m in range(1,13)]
        self._cb_mois = combo(fr_sel, mois_noms)
        cur_label = f"{MOIS_FR[date.today().month]} {date.today().year}"
        if cur_label in mois_noms:
            self._cb_mois.set(cur_label)
        self._cb_mois.pack(side="left", padx=8)
        Btn(fr_sel, text="🔄 Actualiser", command=self._refresh).pack(side="left", padx=8)

        main = tk.Frame(self, bg=C["fond"])
        main.pack(fill="both", expand=True, padx=16, pady=8)

        frame_left = tk.Frame(main, bg=C["fond"])
        frame_left.pack(side="left", fill="both", expand=True, padx=4)
        tk.Label(frame_left, text="Livraisons par quartier", bg=C["fond"], fg=C["primaire"],
                 font=("Helvetica",11,"bold")).pack(anchor="w")
        cols = ("Quartier", "Bons", "Livrés", "Échecs")
        self._tree = ttk.Treeview(frame_left, columns=cols, show="headings", height=12)
        for c,w in zip(cols, [150,80,80,80]):
            self._tree.heading(c, text=c)
            self._tree.column(c, width=w, anchor="center")
        self._tree.column("Quartier", anchor="w")
        self._tree.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(frame_left, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        frame_right = tk.Frame(main, bg=C["fond"])
        frame_right.pack(side="right", fill="both", expand=True, padx=4)
        tk.Label(frame_right, text="Répartition des livraisons", bg=C["fond"], fg=C["primaire"],
                 font=("Helvetica",11,"bold")).pack(anchor="w")
        self._pie = PieChart(frame_right, [], width=350, height=350)
        self._pie.pack(fill="both", expand=True)

    def refresh(self):
        """Rafraîchissement complet de l'onglet (appelé par _refresh)."""
        self._refresh()

    def _refresh(self):
        val = self._cb_mois.get()
        parts = val.split()
        mois = MOIS_FR.index(parts[0])
        annee = int(parts[1])
        data = self.db.get_stats_quartiers_mois_simple(annee, mois)

        for row in self._tree.get_children():
            self._tree.delete(row)
        for q in data:
            self._tree.insert("", "end", values=(q["quartier"], q["total"], q["livres"], q["echecs"]))

        pie_data = [(q["quartier"], q["total"]) for q in data if q["total"] > 0]
        self._pie._data = pie_data
        self._pie._draw()

# ═══════════════════════════════════════════════════════════════════════════════
#  ONGLET RAPPORTS (hérite de BaseOnglet)
# ═══════════════════════════════════════════════════════════════════════════════
class OngletRapports(BaseOnglet):
    """Onglet de génération de rapports, factures et sauvegarde."""
    def __init__(self,parent,db,backup_mgr,status_cb=None):
        self.bak = backup_mgr
        self.status_cb = status_cb
        super().__init__(parent, db)
        self._bons_fac_map={}

    def _build(self):
        barre_titre(self,"📄","Rapports, Factures & Publipostage")
        canvas=tk.Canvas(self,bg=C["fond"],highlightthickness=0)
        sb=ttk.Scrollbar(self,orient="vertical",command=canvas.yview)
        self._inner=tk.Frame(canvas,bg=C["fond"])
        self._inner.bind("<Configure>",lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0),window=self._inner,anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left",fill="both",expand=True)
        sb.pack(side="right",fill="y")

        self._build_rapport()
        self._build_rapport_mensuel()
        self._build_facture()
        self._build_publipostage()
        self._build_sauvegarde()

    def _section(self,titre,emoji=""):
        fr=tk.LabelFrame(self._inner,text=f"  {emoji}  {titre}  ",bg=C["fond"],fg=C["primaire"],
                          font=("Helvetica",10,"bold"),padx=14,pady=10)
        fr.pack(fill="x",padx=20,pady=8)
        return fr

    def _build_rapport(self):
        fr=tk.LabelFrame(self._inner,text="  📋  Rapport Journalier PDF  ",bg=C["fond"],
                          fg=C["primaire"],font=("Helvetica",10,"bold"),padx=14,pady=10)
        fr.pack(fill="x",padx=20,pady=8); fr.columnconfigure(2,weight=1)
        tk.Label(fr,text="Date :",bg=C["fond"],font=("Helvetica",9,"bold")).grid(row=0,column=0,sticky="w",pady=6)
        self._e_drpt=tk.Entry(fr,font=("Helvetica",10),width=13,relief="solid",bd=1)
        self._e_drpt.insert(0,date.today().isoformat()); self._e_drpt.grid(row=0,column=1,padx=8,sticky="w")
        Btn(fr,text="📥  Générer PDF",command=self._gen_rapport).grid(row=0,column=2,padx=4,sticky="w")
        self._lbl_rpt=tk.Label(fr,text="",bg=C["fond"],fg=C["succes"],font=("Helvetica",9))
        self._lbl_rpt.grid(row=1,column=0,columnspan=3,sticky="w",pady=2)

    def _build_rapport_mensuel(self):
        fr=tk.LabelFrame(self._inner,text="  📅  Rapport Mensuel PDF  ",bg=C["fond"],
                          fg=C["primaire"],font=("Helvetica",10,"bold"),padx=14,pady=10)
        fr.pack(fill="x",padx=20,pady=8); fr.columnconfigure(3,weight=1)
        tk.Label(fr,text="Mois :",bg=C["fond"],font=("Helvetica",9,"bold")).grid(row=0,column=0,sticky="w",pady=6)
        mois_noms=[f"{MOIS_FR[m]} {y}" for y in range(date.today().year-1,date.today().year+1) for m in range(1,13)]
        self._cb_mois_rpt=combo(fr,mois_noms)
        cur_label=f"{MOIS_FR[date.today().month]} {date.today().year}"
        if cur_label in mois_noms: self._cb_mois_rpt.set(cur_label)
        self._cb_mois_rpt.grid(row=0,column=1,padx=8,sticky="w")
        Btn(fr,text="📥  Générer PDF",command=self._gen_rapport_mensuel).grid(row=0,column=2,padx=4,sticky="w")
        self._lbl_rpt_m=tk.Label(fr,text="",bg=C["fond"],fg=C["succes"],font=("Helvetica",9))
        self._lbl_rpt_m.grid(row=1,column=0,columnspan=3,sticky="w",pady=2)

    def _build_facture(self):
        fr=tk.LabelFrame(self._inner,text="  🧾  Factures PDF  ",bg=C["fond"],
                          fg=C["primaire"],font=("Helvetica",10,"bold"),padx=14,pady=10)
        fr.pack(fill="x",padx=20,pady=4); fr.columnconfigure(2,weight=1)
        tk.Label(fr,text="Date bons :",bg=C["fond"],font=("Helvetica",9,"bold")).grid(row=0,column=0,sticky="w",pady=6)
        self._e_dfac=tk.Entry(fr,font=("Helvetica",10),width=13,relief="solid",bd=1)
        self._e_dfac.insert(0,date.today().isoformat()); self._e_dfac.grid(row=0,column=1,padx=8,sticky="w")
        Btn(fr,text="🔍  Charger",command=self._charger_bons).grid(row=0,column=2,padx=4,sticky="w")
        tk.Label(fr,text="Tarif (XAF) :",bg=C["fond"],font=("Helvetica",9,"bold")).grid(row=1,column=0,sticky="w",pady=4)
        self._e_tarif=tk.Entry(fr,font=("Helvetica",10),width=10,relief="solid",bd=1)
        self._e_tarif.insert(0,self.db.cfg("tarif_livraison")); self._e_tarif.grid(row=1,column=1,padx=8,sticky="w")
        cols=("N° Bon","Client","Statut","Livreur")
        self._tree_fac=ttk.Treeview(fr,columns=cols,show="headings",height=6)
        for c in cols: self._tree_fac.heading(c,text=c); self._tree_fac.column(c,width=160)
        self._tree_fac.grid(row=2,column=0,columnspan=4,sticky="ew",pady=6)
        Btn(fr,text="🧾  Générer facture (sélection)",command=self._gen_facture).grid(row=3,column=0,columnspan=2,sticky="w",pady=4)
        self._lbl_fac=tk.Label(fr,text="",bg=C["fond"],fg=C["succes"],font=("Helvetica",9))
        self._lbl_fac.grid(row=4,column=0,columnspan=3,sticky="w")

    def _build_publipostage(self):
        fr=tk.LabelFrame(self._inner,text="  📨  Publipostage Clients  ",bg=C["fond"],
                          fg=C["primaire"],font=("Helvetica",10,"bold"),padx=14,pady=10)
        fr.pack(fill="x",padx=20,pady=4); fr.columnconfigure(3,weight=1)

        mois_noms=[f"{MOIS_FR[m]} {y}" for y in range(date.today().year-1,date.today().year+1) for m in range(1,13)]
        tk.Label(fr,text="Mois :",bg=C["fond"],font=("Helvetica",9,"bold")).grid(row=0,column=0,sticky="w",pady=6)
        self._cb_pub=combo(fr,mois_noms)
        cur_label=f"{MOIS_FR[date.today().month]} {date.today().year}"
        if cur_label in mois_noms: self._cb_pub.set(cur_label)
        self._cb_pub.grid(row=0,column=1,padx=8,sticky="w")

        tk.Label(fr,text="Tarif (XAF) :",bg=C["fond"],font=("Helvetica",9,"bold")).grid(row=0,column=2,sticky="w",padx=(16,0))
        self._e_tpub=tk.Entry(fr,font=("Helvetica",10),width=10,relief="solid",bd=1)
        self._e_tpub.insert(0,self.db.cfg("tarif_livraison")); self._e_tpub.grid(row=0,column=3,padx=8,sticky="w")

        Btn(fr,text="📨  Générer courriers PDF",command=self._gen_publipostage,
            bg=C["violet"]).grid(row=1,column=0,columnspan=2,sticky="w",pady=6)
        self._lbl_pub=tk.Label(fr,text="Un courrier par client actif ce mois, en un seul PDF.",
                                bg=C["fond"],fg=C["neutre"],font=("Helvetica",8))
        self._lbl_pub.grid(row=2,column=0,columnspan=4,sticky="w")

    def _build_sauvegarde(self):
        fr=tk.LabelFrame(self._inner,text="  💾  Sauvegarde Automatique  ",bg=C["fond"],
                          fg=C["primaire"],font=("Helvetica",10,"bold"),padx=14,pady=10)
        fr.pack(fill="x",padx=20,pady=4); fr.columnconfigure(2,weight=1)

        tk.Label(fr,text="Intervalle (min) :",bg=C["fond"],font=("Helvetica",9,"bold")).grid(row=0,column=0,sticky="w",pady=6)
        self._e_interval=tk.Entry(fr,font=("Helvetica",10),width=8,relief="solid",bd=1)
        self._e_interval.insert(0,self.db.cfg("backup_interval_min")); self._e_interval.grid(row=0,column=1,padx=8,sticky="w")

        Btn(fr,text="💾  Sauvegarder maintenant",command=self._backup_now).grid(row=0,column=2,padx=8,sticky="w")
        Btn(fr,text="⚙️  Appliquer intervalle",command=self._apply_interval,
            bg=C["neutre"]).grid(row=0,column=3,padx=4,sticky="w")

        tk.Label(fr,text="Sauvegardes disponibles :",bg=C["fond"],
                 font=("Helvetica",9,"bold")).grid(row=1,column=0,columnspan=4,sticky="w",pady=(10,4))
        self._lb_bak=tk.Listbox(fr,height=5,font=("Helvetica",9),relief="solid",bd=1)
        self._lb_bak.grid(row=2,column=0,columnspan=3,sticky="ew",pady=4)

        fr_b2=tk.Frame(fr,bg=C["fond"]); fr_b2.grid(row=3,column=0,columnspan=4,sticky="w",pady=4)
        Btn(fr_b2,text="♻️  Restaurer sélectionné",command=self._restaurer,bg=C["accent"]).pack(side="left",padx=4)
        Btn(fr_b2,text="📁  Ouvrir dossier backups",command=self._ouvrir_dossier,
            bg=C["neutre"]).pack(side="left",padx=4)
        self._lbl_bak=tk.Label(fr,text="",bg=C["fond"],fg=C["succes"],font=("Helvetica",9))
        self._lbl_bak.grid(row=4,column=0,columnspan=4,sticky="w")
        self._refresh_bak()

    def _refresh_bak(self):
        self._lb_bak.delete(0,"end")
        for f in self.bak.liste_backups(): self._lb_bak.insert("end",f)

    def refresh(self):
        """Rafraîchissement de l'onglet (nécessaire pour BaseOnglet)."""
        self._refresh_bak()

    def _gen_rapport(self):
        jour=self._e_drpt.get().strip()
        try: datetime.strptime(jour,"%Y-%m-%d")
        except: messagebox.showerror("Format","Date invalide (AAAA-MM-JJ)"); return
        path=filedialog.asksaveasfilename(defaultextension=".pdf",filetypes=[("PDF","*.pdf")],
                                           initialfile=f"Rapport_{jour}.pdf")
        if not path: return
        try:
            PDF.rapport_journalier(self.db,jour,path)
            self._lbl_rpt.config(text=f"✅ {os.path.basename(path)}")
            if messagebox.askyesno("Ouvrir","Ouvrir le rapport ?"):
                self._ouvrir_pdf(path)
        except Exception as e: messagebox.showerror("Erreur",str(e))

    def _gen_rapport_mensuel(self):
        val=self._cb_mois_rpt.get()
        parts=val.split(); mois=MOIS_FR.index(parts[0]); annee=int(parts[1])
        path=filedialog.asksaveasfilename(defaultextension=".pdf",filetypes=[("PDF","*.pdf")],
                                           initialfile=f"Rapport_{MOIS_FR[mois]}_{annee}.pdf")
        if not path: return
        try:
            PDF.rapport_mensuel(self.db,annee,mois,path)
            self._lbl_rpt_m.config(text=f"✅ {os.path.basename(path)}")
            if messagebox.askyesno("Ouvrir","Ouvrir le rapport ?"):
                self._ouvrir_pdf(path)
        except Exception as e: messagebox.showerror("Erreur",str(e))

    def _charger_bons(self):
        bons=self.db.get_bons_jour(self._e_dfac.get().strip())
        self._bons_fac_map={}
        for r in self._tree_fac.get_children(): self._tree_fac.delete(r)
        for b in bons:
            iid=self._tree_fac.insert("","end",values=(b["numero"],str(b["client_nom"] or "—")[:26],b["statut"],str(b["livreur_nom"] or "—")[:20]))
            self._bons_fac_map[iid]=dict(b)
        self._lbl_fac.config(text=f"{len(bons)} bon(s) chargé(s).")

    def _gen_facture(self):
        sel=self._tree_fac.selection()
        if not sel: messagebox.showwarning("Sélection","Sélectionnez un bon."); return
        bon=self._bons_fac_map[sel[0]]
        try: tarif=int(self._e_tarif.get().strip())
        except: messagebox.showerror("Tarif","Entrez un nombre entier."); return
        path=filedialog.asksaveasfilename(defaultextension=".pdf",filetypes=[("PDF","*.pdf")],
                                           initialfile=f"Facture_{bon['numero']}.pdf")
        if not path: return
        try:
            PDF.facture(self.db,bon["id"],path,tarif)
            self._lbl_fac.config(text=f"✅ Facture : {os.path.basename(path)}")
            if messagebox.askyesno("Ouvrir","Ouvrir la facture ?"):
                self._ouvrir_pdf(path)
        except Exception as e: messagebox.showerror("Erreur",str(e))

    def _gen_publipostage(self):
        val=self._cb_pub.get()
        parts=val.split(); mois=MOIS_FR.index(parts[0]); annee=int(parts[1])
        try: tarif=int(self._e_tpub.get().strip())
        except: messagebox.showerror("Tarif","Entrez un nombre entier."); return
        path=filedialog.asksaveasfilename(defaultextension=".pdf",filetypes=[("PDF","*.pdf")],
                                           initialfile=f"Courriers_{MOIS_FR[mois]}_{annee}.pdf")
        if not path: return
        try:
            n=PDF.publipostage(self.db,annee,mois,path,tarif)
            msg=f"✅ {n} courrier(s) générés : {os.path.basename(path)}" if n else "⚠️ Aucun client avec livraisons ce mois."
            self._lbl_pub.config(text=msg,fg=C["succes"] if n else C["accent"])
            if n and messagebox.askyesno("Ouvrir","Ouvrir les courriers ?"):
                self._ouvrir_pdf(path)
        except Exception as e: messagebox.showerror("Erreur PDF",str(e))

    def _backup_now(self):
        path=self.bak.backup_now()
        if path:
            self._lbl_bak.config(text=f"✅ {os.path.basename(path)}")
            self._refresh_bak()

    def _apply_interval(self):
        try:
            mins=int(self._e_interval.get().strip())
            if mins<1: raise ValueError
        except: messagebox.showerror("Intervalle","Entrez un nombre entier ≥ 1."); return
        self.db.set_cfg("backup_interval_min",str(mins))
        self.bak.stop(); self.bak.start_auto(mins)
        self._lbl_bak.config(text=f"✅ Sauvegarde auto toutes les {mins} min.")

    def _restaurer(self):
        sel=self._lb_bak.curselection()
        if not sel: messagebox.showwarning("Sélection","Sélectionnez un fichier backup."); return
        fname=self._lb_bak.get(sel[0])
        src=os.path.join(BAK_DIR,fname)
        if messagebox.askyesno("Restaurer",f"Restaurer {fname} ?\nL'application se fermera pour appliquer."):
            shutil.copy2(src,DB_PATH)
            messagebox.showinfo("Restauré","Base restaurée. Relancez l'application.")
            import sys; sys.exit(0)

    def _ouvrir_dossier(self):
        os.makedirs(BAK_DIR,exist_ok=True)
        os.startfile(BAK_DIR) if os.name=="nt" else os.system(f"xdg-open '{BAK_DIR}'")

    @staticmethod
    def _ouvrir_pdf(path):
        os.startfile(path) if os.name=="nt" else os.system(f"xdg-open '{path}'")

# ═══════════════════════════════════════════════════════════════════════════════
#  ONGLET CONFIGURATION (hérite de BaseOnglet)
# ═══════════════════════════════════════════════════════════════════════════════
class OngletConfig(BaseOnglet):
    """Onglet de configuration de l'application."""
    def __init__(self, parent, db, backup_mgr, status_cb, update_header_cb=None):
        self.bak = backup_mgr
        self.status_cb = status_cb
        self.update_header_cb = update_header_cb   # callback pour rafraîchir l'en-tête
        super().__init__(parent, db)

    def _build(self):
        barre_titre(self, "⚙️", "Configuration")

        # Nom entreprise
        fr_nom = tk.Frame(self, bg=C["fond"])
        fr_nom.pack(fill="x", padx=20, pady=8)
        tk.Label(fr_nom, text="Nom de l'entreprise :", bg=C["fond"], font=("Helvetica",9,"bold")).pack(side="left")
        self.e_nom = entry(fr_nom, self.db.cfg("nom_entreprise"))
        self.e_nom.pack(side="left", padx=10, fill="x", expand=True)

        # Adresse
        fr_adr = tk.Frame(self, bg=C["fond"])
        fr_adr.pack(fill="x", padx=20, pady=8)
        tk.Label(fr_adr, text="Adresse :", bg=C["fond"], font=("Helvetica",9,"bold")).pack(side="left")
        self.e_adr = entry(fr_adr, self.db.cfg("adresse_entreprise"))
        self.e_adr.pack(side="left", padx=10, fill="x", expand=True)

        # Téléphone
        fr_tel = tk.Frame(self, bg=C["fond"])
        fr_tel.pack(fill="x", padx=20, pady=8)
        tk.Label(fr_tel, text="Téléphone :", bg=C["fond"], font=("Helvetica",9,"bold")).pack(side="left")
        self.e_tel = entry(fr_tel, self.db.cfg("tel_entreprise"))
        self.e_tel.pack(side="left", padx=10, fill="x", expand=True)

        # Email
        fr_mail = tk.Frame(self, bg=C["fond"])
        fr_mail.pack(fill="x", padx=20, pady=8)
        tk.Label(fr_mail, text="Email :", bg=C["fond"], font=("Helvetica",9,"bold")).pack(side="left")
        self.e_mail = entry(fr_mail, self.db.cfg("email_entreprise"))
        self.e_mail.pack(side="left", padx=10, fill="x", expand=True)

        # Tarif livraison
        fr_tarif = tk.Frame(self, bg=C["fond"])
        fr_tarif.pack(fill="x", padx=20, pady=8)
        tk.Label(fr_tarif, text="Tarif livraison (XAF) :", bg=C["fond"], font=("Helvetica",9,"bold")).pack(side="left")
        self.e_tarif = entry(fr_tarif, self.db.cfg("tarif_livraison"))
        self.e_tarif.pack(side="left", padx=10, fill="x", expand=True)

        # Intervalle backup
        fr_back = tk.Frame(self, bg=C["fond"])
        fr_back.pack(fill="x", padx=20, pady=8)
        tk.Label(fr_back, text="Intervalle sauvegarde (minutes) :", bg=C["fond"], font=("Helvetica",9,"bold")).pack(side="left")
        self.e_back = entry(fr_back, self.db.cfg("backup_interval_min"))
        self.e_back.pack(side="left", padx=10, fill="x", expand=True)

        # Boutons
        fr_btn = tk.Frame(self, bg=C["fond"])
        fr_btn.pack(pady=20)
        Btn(fr_btn, text="💾  Enregistrer", command=self._save).pack(side="left", padx=10)
        Btn(fr_btn, text="🔄  Réinitialiser", command=self._reset, bg=C["neutre"]).pack(side="left", padx=10)

    def refresh(self):
        """Rafraîchissement des champs de configuration."""
        pass  # Les champs sont mis à jour via les appels à db.cfg

    def _save(self):
        self.db.set_cfg("nom_entreprise", self.e_nom.get().strip())
        self.db.set_cfg("adresse_entreprise", self.e_adr.get().strip())
        self.db.set_cfg("tel_entreprise", self.e_tel.get().strip())
        self.db.set_cfg("email_entreprise", self.e_mail.get().strip())
        try:
            tarif = int(self.e_tarif.get().strip())
            self.db.set_cfg("tarif_livraison", str(tarif))
        except:
            messagebox.showerror("Erreur", "Tarif invalide, doit être un nombre entier.")
            return
        try:
            mins = int(self.e_back.get().strip())
            if mins < 1:
                raise ValueError
            self.db.set_cfg("backup_interval_min", str(mins))
            self.bak.stop()
            self.bak.start_auto(mins)
        except:
            messagebox.showerror("Erreur", "Intervalle invalide, doit être un entier >= 1.")
            return
        messagebox.showinfo("Configuration", "Configuration enregistrée.")
        if self.status_cb:
            self.status_cb("✅ Configuration mise à jour")
        # Mettre à jour l'en-tête de l'application si le callback est fourni
        if self.update_header_cb:
            self.update_header_cb()

    def _reset(self):
        defaults = {
            "nom_entreprise": "Ponda Logistique Urbain",
            "adresse_entreprise": "Quartier Akwa, Douala",
            "tel_entreprise": "699 00 00 00",
            "email_entreprise": "contact@pondalogistique.cm",
            "tarif_livraison": "2500",
            "backup_interval_min": "60"
        }
        for key, val in defaults.items():
            self.db.set_cfg(key, val)
        self.e_nom.delete(0, tk.END); self.e_nom.insert(0, defaults["nom_entreprise"])
        self.e_adr.delete(0, tk.END); self.e_adr.insert(0, defaults["adresse_entreprise"])
        self.e_tel.delete(0, tk.END); self.e_tel.insert(0, defaults["tel_entreprise"])
        self.e_mail.delete(0, tk.END); self.e_mail.insert(0, defaults["email_entreprise"])
        self.e_tarif.delete(0, tk.END); self.e_tarif.insert(0, defaults["tarif_livraison"])
        self.e_back.delete(0, tk.END); self.e_back.insert(0, defaults["backup_interval_min"])
        try:
            self.bak.stop()
            self.bak.start_auto(int(defaults["backup_interval_min"]))
        except:
            pass
        messagebox.showinfo("Configuration", "Réinitialisation aux valeurs par défaut.")
        if self.status_cb:
            self.status_cb("✅ Configuration réinitialisée")
        if self.update_header_cb:
            self.update_header_cb()

# ═══════════════════════════════════════════════════════════════════════════════
#  APPLICATION PRINCIPALE
# ═══════════════════════════════════════════════════════════════════════════════
class AppPonda(tk.Tk):
    """Application principale."""
    def __init__(self):
        super().__init__()
        self.title("Ponda Logistique Urbain — Système de Gestion v2.4")
        self.geometry("1180x760"); self.minsize(950,620)
        self.configure(bg=C["fond"])
        self._style()
        self.db=Database()
        self.bak=BackupManager(self.db,status_callback=self._set_status)
        self._build()
        # Démarrer sauvegarde auto
        try:
            mins=int(self.db.cfg("backup_interval_min") or 60)
        except Exception:
            mins=60
        self.bak.start_auto(mins)
        # Initialisation du compteur de retards avant le premier appel à _check_alertes
        self._last_retard = 0
        # Vérification alertes toutes les 2 min (optimisée)
        self._check_alertes()

    def _style(self):
        s=ttk.Style(self); s.theme_use("clam")
        s.configure("TNotebook",background=C["fond"],borderwidth=0)
        s.configure("TNotebook.Tab",background=C["bordure"],foreground=C["texte"],
                    font=("Helvetica",9,"bold"),padding=[14,7])
        s.map("TNotebook.Tab",background=[("selected",C["primaire"])],
              foreground=[("selected","white")])
        s.configure("Treeview",background="white",fieldbackground="white",
                    rowheight=26,font=("Helvetica",9))
        s.configure("Treeview.Heading",background=C["primaire"],foreground="white",
                    font=("Helvetica",9,"bold"))
        s.map("Treeview",background=[("selected",C["secondaire"])])

    def _update_header(self):
        """Met à jour l'en-tête principal avec le nom de l'entreprise depuis la base."""
        nom = self.db.cfg("nom_entreprise")
        self._header_label.config(text=f"🚚  {nom.upper()}")
        self.title(f"{nom} — Système de Gestion v2.4")

    def _build(self):
        # En-tête
        ent=tk.Frame(self,bg=C["primaire"],height=52); ent.pack(fill="x"); ent.pack_propagate(False)
        self._header_label = tk.Label(ent, text="", bg=C["primaire"], fg="white",
                                      font=("Helvetica",15,"bold"))
        self._header_label.pack(side="left", padx=20, pady=12)
        tk.Label(ent,text="Système de Gestion — v2.4",bg=C["primaire"],fg="#aed6f1",
                 font=("Helvetica",9)).pack(side="right",padx=20)
        self._update_header()  # initialisation

        # Notebook
        self.nb=ttk.Notebook(self); self.nb.pack(fill="both",expand=True)

        def on_change(): self.dash.refresh()
        def on_incident_change(): self._refresh_badge()

        self.dash   =OngletDashboard(self.nb,self.db)
        self.bons   =OngletBons(self.nb,self.db,on_change=on_change)
        self.clients=OngletClients(self.nb,self.db)
        self.livrs  =OngletLivreurs(self.nb,self.db)
        self.inc    =OngletIncidents(self.nb,self.db,on_change=on_incident_change)
        self.stats  =OngletStatistiques(self.nb,self.db)
        self.carte  =OngletCarte(self.nb,self.db)
        self.rapports=OngletRapports(self.nb,self.db,self.bak,status_cb=self._set_status)
        self.config  =OngletConfig(self.nb,self.db,self.bak,self._set_status, update_header_cb=self._update_header)

        tabs=[
            (self.dash,   "  🏠 Accueil  "),
            (self.bons,   "  📦 Bons  "),
            (self.clients,"  🏪 Clients  "),
            (self.livrs,  "  👥 Livreurs  "),
            (self.inc,    "  🚨 Incidents  "),
            (self.stats,  "  📊 Statistiques  "),
            (self.carte,  "  🗺️ Carte  "),
            (self.rapports,"  📄 Rapports  "),
            (self.config,  "  ⚙️ Configuration  "),
        ]
        for frame,label in tabs: self.nb.add(frame,text=label)

        # Barre statut
        self._status=tk.Label(self,text="✅ Ponda v2.4 — Prêt",bg=C["primaire"],fg="white",
                               font=("Helvetica",8),anchor="w",padx=12)
        self._status.pack(fill="x",side="bottom")

        self._refresh_badge()

    def _set_status(self,msg):
        self._status.config(text=msg)

    def _refresh_badge(self):
        if not hasattr(self, 'inc'):
            return
        nb_inc=self.db.count_incidents_ouverts()
        label=f"  🚨 Incidents ({nb_inc})  " if nb_inc else "  🚨 Incidents  "
        self.nb.tab(self.inc,text=label)

    def _check_alertes(self):
        """Vérification périodique des retards, mise à jour uniquement des alertes."""
        retards = len(self.db.get_bons_retard())
        if retards:
            self._set_status(f"⚠️  {retards} bon(s) en retard non livrés — vérifiez le Dashboard")
            if retards > self._last_retard:
                play_beep()   # Alerte sonore seulement s'il y a augmentation
            self._last_retard = retards
        else:
            self._last_retard = 0
        # Mise à jour uniquement des alertes dans le dashboard, pas un refresh complet
        if hasattr(self, 'dash'):
            self.dash.update_alertes()
        self.after(120_000, self._check_alertes)

    def on_closing(self):
        self.bak.backup_now()
        self.bak.stop()
        self.destroy()

# ═══════════════════════════════════════════════════════════════════════════════
if __name__=="__main__":
    app=AppPonda()
    app.protocol("WM_DELETE_WINDOW",app.on_closing)
    app.mainloop()