#!/usr/bin/env python3
"""Sales Opportunity Hub - gestione opportunità di vendita per rete commerciale.

Backend Flask + SQLite. Avvio:  python app.py          (primo avvio crea l'admin)
                               python app.py --demo   (dati dimostrativi)
"""
import csv
import io
import json
import mimetypes
import os
import re
import secrets
import sqlite3
import sys
import threading
import time
import unicodedata
import uuid
from datetime import date, datetime, timedelta
from functools import wraps

sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor"))  # librerie incluse: non serve installare nulla

from flask import Flask, abort, g, jsonify, request, send_file, send_from_directory, session
from openpyxl import Workbook, load_workbook
from werkzeug.security import check_password_hash, generate_password_hash

import dupcheck
import geo_load
import geo_seed

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("Europe/Rome")
except Exception:  # pragma: no cover
    TZ = None

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("SOH_DATA", os.path.join(BASE, "data"))
DB_PATH = os.environ.get("SOH_DB", os.path.join(DATA_DIR, "hub.db"))
os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
if sys.stdout is None or sys.stderr is None:  # avvio senza finestra (pythonw): messaggi nel file server.log
    _log = open(os.path.join(DATA_DIR, "server.log"), "a", buffering=1, encoding="utf-8")
    sys.stdout = sys.stdout or _log
    sys.stderr = sys.stderr or _log

# ----------------------------------------------------------------- costanti
AREE = ["Triveneto e Lombardia Est", "Centro", "Lombardia Nord e Piemonte Liguria", "Sud", "Milano"]

ROLES = ["admin", "direttore", "area_manager", "district", "team_leader", "agente", "recuperatore", "ufficio_disdette"]
ROLE_LABEL = {
    "admin": "Amministratore", "direttore": "Direttore di area", "area_manager": "Area manager", "district": "District",
    "team_leader": "Team Leader", "agente": "Agente", "recuperatore": "Recuperatore disdette", "ufficio_disdette": "Ufficio Disdette",
}
ROLE_RANK = {"admin": -1, "direttore": 0, "area_manager": 1, "ufficio_disdette": 1, "district": 2, "team_leader": 3, "agente": 4, "recuperatore": 4}

FUNZIONI = [
    ("dashboard", "Dashboard e grafici"),
    ("opportunita", "Vedere le opportunità"),
    ("esiti", "Rendicontare esiti e note"),
    ("agenda", "Agenda appuntamenti"),
    ("assegna", "Riassegnare opportunità"),
    ("campagne", "Campagne e import liste Excel"),
    ("notifiche", "Notifiche"),
    ("email", "Caricare opportunità da email (lead e disdette)"),
    ("proposte", "Proporre opportunità (assistente e inserimento manuale)"),
    ("approva", "Vagliare le proposte degli agenti"),
    ("export", "Esportare la rendicontazione"),
    ("utenti", "Gestione utenti e permessi"),
    ("organigramma", "Organigramma della rete"),
    ("aree", "Gestione aree e archivio geografico"),
    ("impostazioni", "Impostazioni (SLA, limiti)"),
    ("customerbase", "Customer base: caricare e aggiornare l'elenco completo dei clienti"),
    ("verifica", "Verifica clienti già presenti (customer base e altre liste)"),
    ("messaggi", "Ricevere messaggi e allegati dai responsabili"),
    ("messaggi_invia", "Inviare messaggi e allegati alla propria rete"),
    ("documenti", "Documenti personali e consultazione documenti comuni"),
    ("documenti_comuni", "Pubblicare documenti nell'area comune"),
    ("preventivi", "Fare preventivi ai clienti (PDF) e inviarli"),
]
FUNZ_KEYS = [f for f, _ in FUNZIONI]
BASE_AGENT = {"dashboard", "opportunita", "esiti", "agenda", "notifiche", "email", "proposte", "verifica"}
DEFAULTS = {
    "admin": set(FUNZ_KEYS),
    "direttore": {"dashboard", "opportunita", "esiti", "agenda", "assegna", "campagne", "notifiche", "email", "proposte", "approva", "export", "utenti", "organigramma", "impostazioni", "customerbase", "verifica"},
    "area_manager": {"dashboard", "opportunita", "esiti", "agenda", "assegna", "notifiche", "email", "proposte", "approva", "export", "organigramma", "verifica"},
    "district": {"dashboard", "opportunita", "agenda", "assegna", "notifiche", "email", "proposte", "export", "organigramma", "verifica"},
    "team_leader": {"dashboard", "opportunita", "agenda", "notifiche", "email", "proposte", "organigramma", "verifica"},
    "agente": set(BASE_AGENT),
    "recuperatore": set(BASE_AGENT),
    "ufficio_disdette": {"dashboard", "opportunita", "notifiche", "export", "organigramma"},
}

for _r in DEFAULTS:
    if _r != "admin":
        DEFAULTS[_r] |= {"messaggi", "documenti"}
for _r in ("direttore", "area_manager", "district", "team_leader"):
    DEFAULTS[_r] |= {"messaggi_invia", "documenti_comuni"}
for _r in ("direttore", "area_manager", "district", "team_leader", "agente"):
    DEFAULTS[_r] |= {"preventivi"}

STATI = ["da_gestire", "in_lavorazione", "preventivo", "appuntamento", "da_riprogrammare", "contratto", "ko"]
ESITI_CHIAMATA = {"parlato": "Ho parlato con il cliente", "non_risponde": "Non risponde", "da_richiamare": "Da richiamare",
                  "numero_errato": "Numero errato / non raggiungibile", "rifiuta": "Il cliente rifiuta"}
MOTIVI_NON_REC = {"scadenza_naturale": "Non recuperata: portata a naturale scadenza",
                  "con_penali": "Non recuperata: con applicazione penali",
                  "senza_penali": "Non recuperata: senza applicazione penali",
                  "non_lavorabile": "Non lavorabile"}
CATEGORIE = ["customer_base", "prospect", "disdetta"]
SOTTOTIPI = ["lead", "nuova_nata", "lead_prospect"]
REGOLE = ["colonna_agente", "territorio", "round_robin", "non_assegnare"]
SLA_TIPI = [("customer_base", "Customer base"), ("prospect", "Prospect"), ("lead", "Lead"), ("nuova_nata", "Nuova nata"),
            ("lead_prospect", "Lead prospect"), ("disdetta", "Disdette")]
SLA_KEYS = ["sla_" + k for k, _ in SLA_TIPI]
SETTINGS_DEFAULT = {"sla_ore": "48", "max_appuntamenti_giorno": "3", "regola_default": "territorio"}
SETTINGS_DEFAULT.update({k: "" for k in SLA_KEYS})  # vuoto = vale il tempo generale (sla_ore)
SLOT_ORARI = ["09:30", "11:00", "14:30", "16:00", "17:30", "19:00"]

SCHEMA = """
CREATE TABLE IF NOT EXISTS areas(id INTEGER PRIMARY KEY, nome TEXT UNIQUE NOT NULL);
CREATE TABLE IF NOT EXISTS users(
  id INTEGER PRIMARY KEY, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
  nome TEXT NOT NULL, ruolo TEXT NOT NULL, area_id INTEGER REFERENCES areas(id),
  parent_id INTEGER REFERENCES users(id), territori TEXT DEFAULT '', attivo INTEGER DEFAULT 1,
  must_change INTEGER DEFAULT 0, creato TEXT);
CREATE TABLE IF NOT EXISTS user_permissions(
  user_id INTEGER NOT NULL, funzione TEXT NOT NULL, consentito INTEGER NOT NULL,
  PRIMARY KEY(user_id, funzione));
CREATE TABLE IF NOT EXISTS campaigns(
  id INTEGER PRIMARY KEY, nome TEXT NOT NULL, categoria TEXT NOT NULL, sottotipo TEXT DEFAULT '',
  area_id INTEGER, creato_da INTEGER, creato TEXT, regola TEXT, totale INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS opportunities(
  id INTEGER PRIMARY KEY, campaign_id INTEGER, area_id INTEGER, cliente TEXT NOT NULL,
  telefono TEXT DEFAULT '', localita TEXT DEFAULT '', indirizzo TEXT DEFAULT '', note TEXT DEFAULT '',
  agente_id INTEGER, stato TEXT DEFAULT 'da_gestire', esito_nota TEXT DEFAULT '',
  assegnata TEXT, scadenza TEXT, appuntamento TEXT, gestita TEXT,
  n_riprogrammazioni INTEGER DEFAULT 0, sla_notificato INTEGER DEFAULT 0, creata TEXT);
CREATE TABLE IF NOT EXISTS proposals(
  id INTEGER PRIMARY KEY, area_id INTEGER, proponente_id INTEGER NOT NULL, cliente TEXT NOT NULL, telefono TEXT DEFAULT '',
  localita TEXT DEFAULT '', provincia TEXT DEFAULT '', indirizzo TEXT DEFAULT '', note TEXT DEFAULT '', testo TEXT DEFAULT '',
  origine TEXT DEFAULT 'manuale', stato TEXT DEFAULT 'in_attesa', deciso_da INTEGER, deciso TEXT, motivo TEXT DEFAULT '',
  opp_id INTEGER, creata TEXT);
CREATE INDEX IF NOT EXISTS ix_prop ON proposals(area_id, stato);
CREATE TABLE IF NOT EXISTS customer_base(
  id INTEGER PRIMARY KEY, area_id INTEGER NOT NULL, nome TEXT NOT NULL, ident TEXT DEFAULT '', telefono TEXT DEFAULT '',
  localita TEXT DEFAULT '', indirizzo TEXT DEFAULT '', chiave TEXT NOT NULL, batch_id INTEGER);
CREATE INDEX IF NOT EXISTS ix_cb ON customer_base(area_id, chiave);
CREATE TABLE IF NOT EXISTS cb_batches(
  id INTEGER PRIMARY KEY, area_id INTEGER NOT NULL, user_id INTEGER, quando TEXT, file TEXT DEFAULT '', modo TEXT DEFAULT '',
  totale INTEGER DEFAULT 0, nuovi INTEGER DEFAULT 0, rimossi INTEGER DEFAULT 0, aggiornati INTEGER DEFAULT 0, invariati INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS messages(
  id INTEGER PRIMARY KEY, mittente_id INTEGER NOT NULL, area_id INTEGER, titolo TEXT NOT NULL, testo TEXT DEFAULT '', creato TEXT);
CREATE TABLE IF NOT EXISTS message_rcpt(
  message_id INTEGER NOT NULL, user_id INTEGER NOT NULL, letto TEXT, PRIMARY KEY(message_id, user_id));
CREATE INDEX IF NOT EXISTS ix_rcpt_user ON message_rcpt(user_id, letto);
CREATE TABLE IF NOT EXISTS files(
  id INTEGER PRIMARY KEY, ambito TEXT NOT NULL, owner_id INTEGER, area_id INTEGER, message_id INTEGER, cartella TEXT DEFAULT '',
  nome TEXT NOT NULL, mime TEXT DEFAULT '', dimensione INTEGER DEFAULT 0, percorso TEXT NOT NULL, creato TEXT);
CREATE INDEX IF NOT EXISTS ix_files ON files(ambito, owner_id, area_id);
CREATE TABLE IF NOT EXISTS preventivi(
  id INTEGER PRIMARY KEY, opp_id INTEGER NOT NULL, numero TEXT DEFAULT '', autore_id INTEGER, creato TEXT, data_richiesta TEXT,
  cliente TEXT, indirizzo TEXT DEFAULT '', email TEXT DEFAULT '', righe TEXT DEFAULT '[]', totale REAL DEFAULT 0,
  inviato TEXT, inviato_a TEXT DEFAULT '');
CREATE INDEX IF NOT EXISTS ix_prev_opp ON preventivi(opp_id);
CREATE TABLE IF NOT EXISTS cb_staging(
  token TEXT PRIMARY KEY, area_id INTEGER, user_id INTEGER, dati TEXT, file TEXT DEFAULT '', quando TEXT);
CREATE TABLE IF NOT EXISTS events(
  id INTEGER PRIMARY KEY, opp_id INTEGER, user_id INTEGER, azione TEXT, nota TEXT DEFAULT '', quando TEXT);
CREATE TABLE IF NOT EXISTS notifications(
  id INTEGER PRIMARY KEY, user_id INTEGER, messaggio TEXT, opp_id INTEGER, tipo TEXT DEFAULT 'info',
  quando TEXT, letta INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS settings(chiave TEXT PRIMARY KEY, valore TEXT);
CREATE TABLE IF NOT EXISTS geo_regioni(id INTEGER PRIMARY KEY, nome TEXT UNIQUE NOT NULL);
CREATE TABLE IF NOT EXISTS geo_province(
  id INTEGER PRIMARY KEY, nome TEXT NOT NULL, sigla TEXT DEFAULT '', regione_id INTEGER NOT NULL REFERENCES geo_regioni(id),
  UNIQUE(regione_id, nome));
CREATE TABLE IF NOT EXISTS geo_comuni(
  id INTEGER PRIMARY KEY, nome TEXT NOT NULL, provincia_id INTEGER NOT NULL REFERENCES geo_province(id), lat REAL, lon REAL,
  UNIQUE(provincia_id, nome));
CREATE TABLE IF NOT EXISTS user_geo(
  user_id INTEGER NOT NULL, tipo TEXT NOT NULL, geo_id INTEGER NOT NULL, PRIMARY KEY(user_id, tipo, geo_id));
CREATE INDEX IF NOT EXISTS ix_geo_prov ON geo_province(regione_id);
CREATE INDEX IF NOT EXISTS ix_geo_com ON geo_comuni(provincia_id);
CREATE INDEX IF NOT EXISTS ix_opp_agente ON opportunities(agente_id, stato);
CREATE INDEX IF NOT EXISTS ix_opp_area ON opportunities(area_id, stato);
CREATE INDEX IF NOT EXISTS ix_opp_camp ON opportunities(campaign_id);
CREATE INDEX IF NOT EXISTS ix_ev_opp ON events(opp_id);
CREATE INDEX IF NOT EXISTS ix_notif_user ON notifications(user_id, letta);
"""

app = Flask(__name__, static_folder=os.path.join(BASE, "static"), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 70 * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.environ.get("SOH_SECURE_COOKIE", "0") == "1"
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=12)


def _secret():
    p = os.path.join(DATA_DIR, "secret.key")
    if not os.path.exists(p):
        with open(p, "w") as f:
            f.write(secrets.token_hex(32))
        os.chmod(p, 0o600)
    with open(p) as f:
        return f.read().strip()


app.secret_key = os.environ.get("SOH_SECRET") or _secret()


# ----------------------------------------------------------------- utilità
def now():
    return (datetime.now(TZ) if TZ else datetime.now()).replace(tzinfo=None, microsecond=0)


def ts(dt=None):
    return (dt or now()).strftime("%Y-%m-%d %H:%M:%S")


def connect():
    c = sqlite3.connect(DB_PATH, timeout=30)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    c.execute("PRAGMA journal_mode=WAL")
    return c


def db():
    if "db" not in g:
        g.db = connect()
    return g.db


@app.teardown_appcontext
def _close(_e):
    c = g.pop("db", None)
    if c is not None:
        c.close()


def get_setting(conn, key):
    r = conn.execute("SELECT valore FROM settings WHERE chiave=?", (key,)).fetchone()
    return r["valore"] if r else SETTINGS_DEFAULT[key]


def sla_tipo(categoria, sottotipo):
    if categoria == "disdetta":
        return "disdetta"
    if sottotipo in ("lead", "nuova_nata", "lead_prospect"):
        return sottotipo
    return categoria if categoria in ("customer_base", "prospect") else "prospect"


def sla_ore(conn, categoria, sottotipo=""):
    """Ore di SLA per il tipo di opportunità (regola specifica, altrimenti quella generale)."""
    v = get_setting(conn, "sla_" + sla_tipo(categoria, sottotipo))
    try:
        return max(1, int(v)) if str(v).strip() else int(get_setting(conn, "sla_ore"))
    except ValueError:
        return int(get_setting(conn, "sla_ore"))


def init_db():
    conn = connect()
    conn.executescript(SCHEMA)
    for a in AREE:
        conn.execute("INSERT OR IGNORE INTO areas(nome) VALUES(?)", (a,))
    for k, v in SETTINGS_DEFAULT.items():
        conn.execute("INSERT OR IGNORE INTO settings(chiave,valore) VALUES(?,?)", (k, v))
    geo_seed.seed(conn)
    if "comune_id" not in [r[1] for r in conn.execute("PRAGMA table_info(opportunities)")]:
        conn.execute("ALTER TABLE opportunities ADD COLUMN comune_id INTEGER")
    if "rif_esterno" not in [r[1] for r in conn.execute("PRAGMA table_info(opportunities)")]:
        conn.execute("ALTER TABLE opportunities ADD COLUMN rif_esterno TEXT DEFAULT ''")
    if "cal_token" not in [r[1] for r in conn.execute("PRAGMA table_info(users)")]:
        conn.execute("ALTER TABLE users ADD COLUMN cal_token TEXT")
    cols = [r[1] for r in conn.execute("PRAGMA table_info(opportunities)")]
    for nome, ddl in (("disdetta_esito", "TEXT"), ("disdetta_motivo", "TEXT"), ("valore", "REAL"), ("chiusa", "TEXT"),
                      ("n_chiamate", "INTEGER DEFAULT 0"), ("canone", "REAL"), ("ultima_chiamata", "TEXT")):
        if nome not in cols:
            conn.execute("ALTER TABLE opportunities ADD COLUMN %s %s" % (nome, ddl))
    # chiusure già esistenti: data di chiusura dallo storico; le disdette chiuse diventano recuperata / non recuperata
    conn.execute("UPDATE opportunities SET chiusa=COALESCE((SELECT MAX(quando) FROM events e WHERE e.opp_id=opportunities.id "
                 "AND e.azione IN ('contratto','ko')), gestita, creata) WHERE chiusa IS NULL AND stato IN ('contratto','ko')")
    conn.execute("UPDATE opportunities SET disdetta_esito=CASE stato WHEN 'contratto' THEN 'recuperata' ELSE 'non_recuperata' END "
                 "WHERE disdetta_esito IS NULL AND stato IN ('contratto','ko') AND campaign_id IN (SELECT id FROM campaigns WHERE categoria='disdetta')")
    conn.execute("CREATE INDEX IF NOT EXISTS ix_opp_comune ON opportunities(comune_id)")
    if "dup_info" not in cols:
        conn.execute("ALTER TABLE opportunities ADD COLUMN dup_info TEXT DEFAULT ''")
    if "con_preventivo" not in cols:
        conn.execute("ALTER TABLE opportunities ADD COLUMN con_preventivo INTEGER DEFAULT 0")
    for tab in ("opportunities", "proposals"):
        tc = [r[1] for r in conn.execute("PRAGMA table_info(%s)" % tab)]
        for nome in ("lat", "lon"):
            if nome not in tc:
                conn.execute("ALTER TABLE %s ADD COLUMN %s REAL" % (tab, nome))
    if "dup_info" not in [r[1] for r in conn.execute("PRAGMA table_info(proposals)")]:
        conn.execute("ALTER TABLE proposals ADD COLUMN dup_info TEXT DEFAULT ''")
    if conn.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
        conn.execute(
            "INSERT INTO users(username,password_hash,nome,ruolo,must_change,creato) VALUES(?,?,?,?,1,?)",
            ("admin", generate_password_hash("admin"), "Amministratore", "admin", ts()))
        print(">> Creato utente 'admin' con password 'admin' (verrà richiesto il cambio al primo accesso)")
    conn.commit()
    conn.close()


def body():
    return request.get_json(silent=True) or {}


def err(msg, code=400):
    return jsonify({"error": msg}), code


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


# ----------------------------------------------------------------- utenti e permessi
def current_user():
    if "_u" in g:
        return g._u
    uid = session.get("uid")
    g._u = None
    if uid:
        g._u = db().execute("SELECT * FROM users WHERE id=? AND attivo=1", (uid,)).fetchone()
    return g._u


def perms_of(conn, u):
    if u["ruolo"] == "admin":
        return set(FUNZ_KEYS)
    p = set(DEFAULTS[u["ruolo"]])
    for r in conn.execute("SELECT funzione, consentito FROM user_permissions WHERE user_id=?", (u["id"],)):
        if r["funzione"] in FUNZ_KEYS:
            (p.add if r["consentito"] else p.discard)(r["funzione"])
    return p


def need(funzione=None):
    def deco(fn):
        @wraps(fn)
        def wrap(*a, **kw):
            u = current_user()
            if not u:
                return err("Non autenticato", 401)
            if u["must_change"] and request.endpoint not in ("api_password", "api_me", "api_logout"):
                return err("Cambio password richiesto", 403)
            if funzione and funzione not in perms_of(db(), u):
                return err("Funzione non consentita per il tuo profilo", 403)
            return fn(*a, **kw)
        return wrap
    return deco


@app.before_request
def _csrf():
    if request.path.startswith("/api/") and request.method in ("POST", "PUT", "DELETE", "PATCH"):
        if request.headers.get("X-Requested-With") != "SOH":
            return err("Richiesta non valida", 400)


def subtree_ids(uid):
    rows = db().execute(
        "WITH RECURSIVE t(id) AS (SELECT ? UNION SELECT u.id FROM users u JOIN t ON u.parent_id=t.id) SELECT id FROM t",
        (uid,)).fetchall()
    return [r["id"] for r in rows]


def visible_user_ids(u):
    if u["ruolo"] == "admin":
        return [r["id"] for r in db().execute("SELECT id FROM users")]
    if u["ruolo"] == "direttore":
        return [r["id"] for r in db().execute("SELECT id FROM users WHERE area_id=?", (u["area_id"],))]
    if u["ruolo"] == "ufficio_disdette":
        return [r["id"] for r in db().execute("SELECT id FROM users WHERE area_id=? AND (ruolo='recuperatore' OR id=?)", (u["area_id"], u["id"]))]
    if u["ruolo"] == "area_manager":  # la sua rete + chi lavora opportunità nelle sue zone
        ids = set(subtree_ids(u["id"]))
        ids |= {r[0] for r in db().execute("SELECT DISTINCT o.agente_id FROM opportunities o WHERE o.area_id=? AND o.agente_id IS NOT NULL AND "
                                            + ZONE_AM, (u["area_id"],) + (u["id"],) * 3)}
        return sorted(ids)
    return subtree_ids(u["id"])


# comuni delle zone assegnate all'Area manager: province senza comuni specifici = tutti i comuni della provincia (3 parametri: id utente)
ZONE_AM = ("o.comune_id IN (SELECT id FROM geo_comuni WHERE provincia_id IN (SELECT geo_id FROM user_geo WHERE user_id=? AND tipo='provincia' "
           "AND geo_id NOT IN (SELECT provincia_id FROM geo_comuni WHERE id IN (SELECT geo_id FROM user_geo WHERE user_id=? AND tipo='comune'))) "
           "UNION SELECT geo_id FROM user_geo WHERE user_id=? AND tipo='comune')")
SOLO_DISDETTE = "o.campaign_id IN (SELECT id FROM campaigns WHERE categoria='disdetta')"


def scope_where(u, agent_id=None, area_id=None, categoria=None):
    """Clausola SQL di visibilità sulle opportunità (alias o, campaigns alias c)."""
    conds, ps = [], []
    if u["ruolo"] == "admin":
        if area_id:
            conds.append("o.area_id=?")
            ps.append(area_id)
    elif u["ruolo"] == "direttore":
        conds.append("o.area_id=?")
        ps.append(u["area_id"])
    elif u["ruolo"] == "area_manager":  # come il direttore, ma per le zone assegnate (più la sua rete)
        ids = subtree_ids(u["id"])
        conds.append("o.area_id=? AND (%s OR o.agente_id IN (%s))" % (ZONE_AM, ",".join("?" * len(ids))))
        ps += [u["area_id"], u["id"], u["id"], u["id"]] + ids
    elif u["ruolo"] == "ufficio_disdette":  # tutte le disdette dell'area, qualunque recuperatore
        conds.append("o.area_id=? AND " + SOLO_DISDETTE)
        ps.append(u["area_id"])
    else:
        ids = subtree_ids(u["id"])
        conds.append("o.agente_id IN (%s)" % ",".join("?" * len(ids)))
        ps += ids
    if agent_id:
        if agent_id not in visible_user_ids(u):
            abort(403)
        sub = subtree_ids(agent_id)
        conds.append("o.agente_id IN (%s)" % ",".join("?" * len(sub)))
        ps += sub
    if categoria in CATEGORIE:
        conds.append("c.categoria=?")
        ps.append(categoria)
    return (" AND ".join(conds) or "1=1"), ps


def can_access(u, o):
    if u["ruolo"] == "admin":
        return True
    if u["ruolo"] == "direttore":
        return o["area_id"] == u["area_id"]
    if u["ruolo"] == "area_manager":
        if o["area_id"] != u["area_id"]:
            return False
        if o["agente_id"] in subtree_ids(u["id"]):
            return True
        return bool(db().execute("SELECT 1 FROM opportunities o WHERE o.id=? AND " + ZONE_AM, (o["id"], u["id"], u["id"], u["id"])).fetchone())
    if u["ruolo"] == "ufficio_disdette":
        return o["area_id"] == u["area_id"] and bool(db().execute(
            "SELECT 1 FROM campaigns WHERE id=? AND categoria='disdetta'", (o["campaign_id"],)).fetchone())
    return o["agente_id"] in subtree_ids(u["id"])


def ancestors(conn, uid):
    out, seen = [], set()
    cur = conn.execute("SELECT parent_id FROM users WHERE id=?", (uid,)).fetchone()
    while cur and cur["parent_id"] and cur["parent_id"] not in seen:
        pid = cur["parent_id"]
        seen.add(pid)
        out.append(pid)
        cur = conn.execute("SELECT parent_id FROM users WHERE id=?", (pid,)).fetchone()
    return out


def notify(conn, user_ids, msg, tipo="info", opp_id=None):
    for uid in dict.fromkeys(user_ids):
        conn.execute("INSERT INTO notifications(user_id,messaggio,opp_id,tipo,quando) VALUES(?,?,?,?,?)",
                     (uid, msg, opp_id, tipo, ts()))


def cc_for_assignment(conn, agent_id):
    """Agente destinatario + District e Team Leader in conoscenza."""
    ids = [agent_id]
    for pid in ancestors(conn, agent_id):
        r = conn.execute("SELECT ruolo FROM users WHERE id=? AND attivo=1", (pid,)).fetchone()
        if r and r["ruolo"] in ("area_manager", "district", "team_leader"):
            ids.append(pid)
    return ids


def log_event(conn, oid, uid, azione, nota=""):
    conn.execute("INSERT INTO events(opp_id,user_id,azione,nota,quando) VALUES(?,?,?,?,?)", (oid, uid, azione, nota, ts()))


# ----------------------------------------------------------------- SLA
def check_sla(conn):
    """Notifica a tutta la catena le attività non lavorate entro l'SLA (una sola volta)."""
    rows = conn.execute(
        "SELECT o.id,o.cliente,o.agente_id,o.area_id,u.nome AS agente,k.categoria,k.sottotipo FROM opportunities o "
        "JOIN users u ON u.id=o.agente_id LEFT JOIN campaigns k ON k.id=o.campaign_id WHERE o.stato='da_gestire' AND o.sla_notificato=0 "
        "AND o.scadenza IS NOT NULL AND o.scadenza<?", (ts(),)).fetchall()
    if not rows:
        return 0
    by_agent = {}
    for r in rows:
        by_agent.setdefault(r["agente_id"], []).append(r)
    for aid, lst in by_agent.items():
        dest = [aid] + ancestors(conn, aid)
        dest += [r["id"] for r in conn.execute(
            "SELECT id FROM users WHERE ruolo='direttore' AND attivo=1 AND area_id=?", (lst[0]["area_id"],))]
        n = len(lst)
        msg = "SLA %dh scaduto: %s ha %d opportunità non lavorate (es. %s)" % (sla_ore(conn, lst[0]["categoria"], lst[0]["sottotipo"]), lst[0]["agente"], n, lst[0]["cliente"])
        notify(conn, dest, msg, "sla", lst[0]["id"])
        conn.executemany("UPDATE opportunities SET sla_notificato=1 WHERE id=?", [(r["id"],) for r in lst])
    conn.commit()
    return len(rows)


def _sla_loop():
    while True:
        try:
            c = connect()
            check_sla(c)
            c.close()
        except Exception as e:  # pragma: no cover
            print("SLA check error:", e)
        time.sleep(300)


# ----------------------------------------------------------------- auth
_fails = {}


@app.post("/api/login")
def api_login():
    d = body()
    uname = (d.get("username") or "").strip().lower()
    key = (request.remote_addr, uname)
    recent = [t for t in _fails.get(key, []) if t > time.time() - 600]
    if len(recent) >= 8:
        return err("Troppi tentativi. Riprova tra qualche minuto.", 429)
    u = db().execute("SELECT * FROM users WHERE lower(username)=? AND attivo=1", (uname,)).fetchone()
    if not u or not check_password_hash(u["password_hash"], d.get("password") or ""):
        recent.append(time.time())
        _fails[key] = recent
        return err("Credenziali non valide", 401)
    _fails.pop(key, None)
    session.clear()
    session["uid"] = u["id"]
    session.permanent = True
    return api_me()


@app.post("/api/logout")
def api_logout():
    session.clear()
    return jsonify({"ok": True})


@app.get("/api/me")
def api_me():
    u = current_user()
    if not u:
        return err("Non autenticato", 401)
    c = db()
    area = c.execute("SELECT nome FROM areas WHERE id=?", (u["area_id"],)).fetchone()
    nl = c.execute("SELECT COUNT(*) FROM notifications WHERE user_id=? AND letta=0", (u["id"],)).fetchone()[0]
    return jsonify({
        "user": {"id": u["id"], "username": u["username"], "nome": u["nome"], "ruolo": u["ruolo"],
                 "ruolo_label": ROLE_LABEL[u["ruolo"]], "area_id": u["area_id"],
                 "area": area["nome"] if area else None, "must_change": bool(u["must_change"])},
        "permessi": sorted(perms_of(c, u)), "non_lette": nl})


@app.post("/api/password")
def api_password():
    u = current_user()
    if not u:
        return err("Non autenticato", 401)
    d = body()
    if not check_password_hash(u["password_hash"], d.get("attuale") or ""):
        return err("Password attuale errata", 400)
    nuova = d.get("nuova") or ""
    if len(nuova) < 8 or nuova.lower() == "admin":
        return err("La nuova password deve avere almeno 8 caratteri", 400)
    db().execute("UPDATE users SET password_hash=?, must_change=0 WHERE id=?", (generate_password_hash(nuova), u["id"]))
    db().commit()
    return jsonify({"ok": True})


@app.get("/api/meta")
@need()
def api_meta():
    c = db()
    return jsonify({
        "aree": [dict(r) for r in c.execute("SELECT id,nome FROM areas ORDER BY id")],
        "ruoli": [{"id": r, "label": ROLE_LABEL[r]} for r in ROLES],
        "funzioni": [{"id": k, "label": v} for k, v in FUNZIONI],
        "default_ruolo": {r: sorted(p) for r, p in DEFAULTS.items()},
        "stati": STATI, "categorie": CATEGORIE, "sottotipi": SOTTOTIPI, "regole": REGOLE,
        "settings": {k: get_setting(c, k) for k in SETTINGS_DEFAULT},
        "sla_tipi": [{"id": k, "label": l} for k, l in SLA_TIPI],
        "locale": request.host.split(":")[0] in ("127.0.0.1", "localhost"),
        "tile_url": os.environ.get("SOH_TILE_URL", "https://tile.openstreetmap.org/{z}/{x}/{y}.png"),
        "tile_attr": os.environ.get("SOH_TILE_ATTR", "© OpenStreetMap")})


@app.get("/api/agenti")
@need()
def api_agenti():
    """Utenti visibili (per filtri e assegnazioni)."""
    u = current_user()
    ids = visible_user_ids(u)
    rows = db().execute(
        "SELECT u.id,u.nome,u.ruolo,u.area_id,u.territori,p.nome AS riferimento FROM users u "
        "LEFT JOIN users p ON p.id=u.parent_id WHERE u.attivo=1 AND u.ruolo IN ('area_manager','district','team_leader','agente','recuperatore') "
        "AND u.id IN (%s) ORDER BY u.nome" % ",".join("?" * len(ids)), ids).fetchall()
    return jsonify([dict(r) for r in rows])


# ----------------------------------------------------------------- opportunità
BASE_SELECT = (
    "SELECT o.*, c.nome AS campagna, c.categoria, c.sottotipo, a.nome AS agente, ar.nome AS area "
    "FROM opportunities o LEFT JOIN campaigns c ON c.id=o.campaign_id "
    "LEFT JOIN users a ON a.id=o.agente_id LEFT JOIN areas ar ON ar.id=o.area_id")


def opp_dict(r):
    d = dict(r)
    n = ts()
    d["scaduta"] = bool(r["stato"] == "da_gestire" and r["scadenza"] and r["scadenza"] < n)
    d["ore_residue"] = None
    if r["stato"] == "da_gestire" and r["scadenza"]:
        delta = datetime.strptime(r["scadenza"], "%Y-%m-%d %H:%M:%S") - now()
        d["ore_residue"] = round(delta.total_seconds() / 3600, 1)
    return d


def filters_from_request(u):
    q = request.args
    agent_id = int(q["agent_id"]) if q.get("agent_id", "").isdigit() else None
    area_id = int(q["area_id"]) if q.get("area_id", "").isdigit() else None
    where, ps = scope_where(u, agent_id, area_id, q.get("categoria"))
    if q.get("stato") in STATI:
        where += " AND o.stato=?"
        ps.append(q["stato"])
    if q.get("campaign_id", "").isdigit():
        where += " AND o.campaign_id=?"
        ps.append(int(q["campaign_id"]))
    if q.get("scadute") == "1":
        where += " AND o.stato='da_gestire' AND o.scadenza IS NOT NULL AND o.scadenza<?"
        ps.append(ts())
    if q.get("non_assegnate") == "1":
        where += " AND o.agente_id IS NULL"
    if q.get("q"):
        like = "%" + q["q"].strip().lower() + "%"
        where += " AND (lower(o.cliente) LIKE ? OR lower(o.localita) LIKE ? OR o.telefono LIKE ?)"
        ps += [like, like, like]
    return where, ps


@app.get("/api/opportunities")
@need("opportunita")
def api_opps():
    u = current_user()
    where, ps = filters_from_request(u)
    limit = min(max(request.args.get("limit", 300, type=int) or 300, 1), 1000)
    rows = db().execute(
        BASE_SELECT + " WHERE " + where + " ORDER BY (o.stato='da_gestire') DESC, o.scadenza IS NULL, o.scadenza, o.id LIMIT ?",
        ps + [limit]).fetchall()
    tot = db().execute("SELECT COUNT(*) FROM opportunities o LEFT JOIN campaigns c ON c.id=o.campaign_id WHERE " + where, ps).fetchone()[0]
    return jsonify({"totale": tot, "items": [opp_dict(r) for r in rows]})


def load_opp(oid, u):
    o = db().execute(BASE_SELECT + " WHERE o.id=?", (oid,)).fetchone()
    if not o or not can_access(u, o):
        abort(404)
    return o


@app.get("/api/opportunities/<int:oid>")
@need("opportunita")
def api_opp(oid):
    u = current_user()
    o = load_opp(oid, u)
    ev = db().execute(
        "SELECT e.azione,e.nota,e.quando,u.nome AS utente FROM events e LEFT JOIN users u ON u.id=e.user_id "
        "WHERE e.opp_id=? ORDER BY e.id DESC", (oid,)).fetchall()
    d = opp_dict(o)
    d["eventi"] = [dict(e) for e in ev]
    return jsonify(d)


def _valore(v):
    """Importo in euro scritto all'italiana (1.250,50), all'inglese (1250.50) o come numero."""
    if isinstance(v, (int, float)):
        x = float(v)
    else:
        t = str(v or "").replace("€", "").replace(" ", "").strip()
        if "," in t:
            t = t.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"\d{1,3}(\.\d{3})+", t):
            t = t.replace(".", "")
        try:
            x = float(t)
        except ValueError:
            return None
    return round(x, 2) if 0 <= x < 1e9 else None


@app.post("/api/opportunities/<int:oid>/esito")
@need("esiti")
def api_esito(oid):
    u = current_user()
    o = load_opp(oid, u)
    d = body()
    es = d.get("esito")
    nota = (d.get("nota") or "").strip()[:2000]
    disdetta = o["categoria"] == "disdetta"
    if o["stato"] in ("contratto", "ko"):
        return err("Opportunità già chiusa")
    c = db()
    if disdetta and es in ("recuperata", "non_recuperata"):
        if not nota:
            return err("Scrivi una nota con i dettagli della lavorazione")
        can = _valore(d.get("canone"))
        if can is None:
            return err("Conferma il canone originario del contratto (importo in euro)")
        if es == "recuperata":
            val = _valore(d.get("valore"))
            if val is None:
                return err("Indica il valore della disdetta recuperata (importo in euro)")
            c.execute("UPDATE opportunities SET stato='contratto', esito_nota=?, disdetta_esito='recuperata', disdetta_motivo=NULL, "
                      "valore=?, canone=?, chiusa=?, gestita=COALESCE(gestita,?) WHERE id=?", (nota, val, can, ts(), ts(), oid))
            log_event(c, oid, u["id"], "disdetta recuperata", "Valore recuperato %.2f € (canone originario %.2f €). %s" % (val, can, nota))
        else:
            mot = d.get("motivo")
            if mot not in MOTIVI_NON_REC:
                return err("Scegli il motivo per cui la disdetta non è stata recuperata")
            c.execute("UPDATE opportunities SET stato='ko', esito_nota=?, disdetta_esito='non_recuperata', disdetta_motivo=?, "
                      "valore=NULL, canone=?, chiusa=?, gestita=COALESCE(gestita,?) WHERE id=?", (nota, mot, can, ts(), ts(), oid))
            log_event(c, oid, u["id"], "disdetta non recuperata", "%s (canone originario %.2f €). %s" % (MOTIVI_NON_REC[mot], can, nota))
        c.commit()
        return jsonify({"ok": True})
    if disdetta and es in ("contratto", "ko"):
        return err("Per le disdette usa «Disdetta recuperata» o «Disdetta non recuperata»")
    if es not in ("contratto", "ko", "da_riprogrammare"):
        return err("Esito non valido")
    if es == "da_riprogrammare":
        c.execute("UPDATE opportunities SET stato='da_riprogrammare', appuntamento=NULL, esito_nota=?, "
                  "n_riprogrammazioni=n_riprogrammazioni+1, gestita=COALESCE(gestita,?) WHERE id=?", (nota, ts(), oid))
    else:
        c.execute("UPDATE opportunities SET stato=?, esito_nota=?, chiusa=?, gestita=COALESCE(gestita,?) WHERE id=?", (es, nota, ts(), ts(), oid))
    log_event(c, oid, u["id"], es, nota)
    c.commit()
    return jsonify({"ok": True})


@app.post("/api/opportunities/<int:oid>/chiamata")
@need("esiti")
def api_chiamata(oid):
    """Registro delle chiamate al cliente (prima dell'appuntamento o della chiusura)."""
    u = current_user()
    o = load_opp(oid, u)
    d = body()
    es = d.get("esito")
    nota = (d.get("nota") or "").strip()[:2000]
    if es not in ESITI_CHIAMATA:
        return err("Scegli l'esito della chiamata")
    if o["stato"] in ("contratto", "ko"):
        return err("Opportunità già chiusa")
    c = db()
    nuovo = "in_lavorazione" if o["stato"] in ("da_gestire", "da_riprogrammare") else o["stato"]
    c.execute("UPDATE opportunities SET stato=?, n_chiamate=COALESCE(n_chiamate,0)+1, ultima_chiamata=?, gestita=COALESCE(gestita,?) WHERE id=?",
              (nuovo, ts(), ts(), oid))
    log_event(c, oid, u["id"], "chiamata", ESITI_CHIAMATA[es] + (". " + nota if nota else ""))
    c.commit()
    return jsonify({"ok": True, "stato": nuovo})


@app.post("/api/opportunities/<int:oid>/nota")
@need("esiti")
def api_nota(oid):
    u = current_user()
    load_opp(oid, u)
    nota = (body().get("nota") or "").strip()[:2000]
    if not nota:
        return err("Nota vuota")
    log_event(db(), oid, u["id"], "nota", nota)
    db().commit()
    return jsonify({"ok": True})


def _it_data(data):
    return datetime.strptime(data, "%Y-%m-%d").strftime("%d/%m/%Y")


def verifica_agenda(c, agente_id, data, ora, escludi=None):
    """Controlla l'agenda dell'agente per il giorno/ora scelti. Non blocca: restituisce gli avvisi (limite giornaliero superato,
    appuntamenti concomitanti = a meno di un'ora di distanza) e un'alternativa suggerita."""
    mx = int(get_setting(c, "max_appuntamenti_giorno"))
    nuovo = datetime.strptime("%s %s" % (data, ora), "%Y-%m-%d %H:%M")

    def del_giorno(g):
        return [{"id": r["id"], "cliente": r["cliente"], "localita": r["localita"] or "", "ora": r["appuntamento"][11:16]}
                for r in c.execute("SELECT id,cliente,localita,appuntamento FROM opportunities WHERE agente_id=? AND stato='appuntamento' "
                                   "AND substr(appuntamento,1,10)=? AND id<>? ORDER BY appuntamento", (agente_id, g, escludi or 0))]

    def vicini(giorno, quando):
        return [x for x in giorno if abs((datetime.strptime("%s %s" % (quando.strftime("%Y-%m-%d"), x["ora"]), "%Y-%m-%d %H:%M") - quando).total_seconds()) < 3600]

    def slot_libero(giorno, g, preferita):
        cand = ["%02d:%02d" % (h, m) for h in range(8, 20) for m in (0, 30)]
        base = datetime.strptime("%s %s" % (g, preferita), "%Y-%m-%d %H:%M")
        liberi = [x for x in cand if not vicini(giorno, datetime.strptime("%s %s" % (g, x), "%Y-%m-%d %H:%M"))]
        liberi.sort(key=lambda x: abs((datetime.strptime("%s %s" % (g, x), "%Y-%m-%d %H:%M") - base).total_seconds()))
        return liberi[0] if liberi else None

    giorno = del_giorno(data)
    conc = vicini(giorno, nuovo)
    sopra = len(giorno) >= mx
    sug = None
    if conc and not sopra:
        o_ = slot_libero(giorno, data, ora)
        if o_:
            sug = {"data": data, "ora": o_}
    elif sopra:
        d_ = nuovo.date()
        for _ in range(21):
            d_ += timedelta(days=1)
            if d_.weekday() >= 5:
                continue
            g_ = d_.isoformat()
            dg = del_giorno(g_)
            if len(dg) < mx:
                o_ = slot_libero(dg, g_, ora)
                if o_:
                    sug = {"data": g_, "ora": o_}
                    break
    parti = []
    if sopra:
        parti.append("Il %s ci sono già %d appuntamenti in agenda (limite consigliato: %d)." % (_it_data(data), len(giorno), mx))
    if conc:
        parti.append("È concomitante con: %s. Conviene spostare uno dei due appuntamenti." %
                     "; ".join("%s alle %s%s" % (x["cliente"], x["ora"], " (%s)" % x["localita"] if x["localita"] else "") for x in conc))
    return {"data": data, "ora": ora, "limite": mx, "n_giorno": len(giorno), "sopra_limite": sopra, "concomitanti": conc,
            "giorno": giorno, "suggerito": sug, "problema": bool(sopra or conc), "avviso": " ".join(parti)}


def set_appointment(c, o, data, ora, user, forza=False):
    """Restituisce None se fissato, un testo se c'è un errore che non si può superare, oppure un dict di avvisi
    (limite superato / appuntamenti concomitanti) quando non è stato confermato con forza=True."""
    try:
        dd = datetime.strptime(data, "%Y-%m-%d").date()
        datetime.strptime(ora, "%H:%M")
    except (ValueError, TypeError):
        return "Data o ora non valida"
    if dd < now().date():
        return "La data non può essere nel passato"
    if o["agente_id"] is None:
        return "Opportunità non assegnata a un agente"
    if o["stato"] in ("contratto", "ko"):
        return "Opportunità già chiusa"
    info = verifica_agenda(c, o["agente_id"], data, ora, o["id"])
    if info["problema"] and not forza:
        return info
    c.execute("UPDATE opportunities SET stato='appuntamento', appuntamento=?, gestita=COALESCE(gestita,?) WHERE id=?",
              ("%s %s:00" % (data, ora), ts(), o["id"]))
    log_event(c, o["id"], user["id"], "appuntamento", "%s %s" % (data, ora))
    if info["problema"]:
        log_event(c, o["id"], user["id"], "nota", "Appuntamento inserito dopo avviso: " + info["avviso"])
    return None


@app.post("/api/opportunities/<int:oid>/appuntamento")
@need("agenda")
def api_appuntamento(oid):
    u = current_user()
    o = load_opp(oid, u)
    d = body()
    e = set_appointment(db(), o, d.get("data"), d.get("ora"), u, bool(d.get("forza")))
    if e is None and d.get("verifica_preventivo"):
        log_event(db(), oid, u["id"], "nota", "Appuntamento per la verifica del preventivo con il cliente")
    if isinstance(e, dict):
        return jsonify({"error": e["avviso"], "conferma": e}), 409
    if e:
        return err(e, 409)
    db().commit()
    return jsonify({"ok": True})


# ----------------------------------------------------------------- preventivi (PDF nel formato dell'offerta tipo)
import preventivo as prev_mod


def _prev_dict(r, completo=False):
    righe = json.loads(r["righe"] or "[]")
    d = {"id": r["id"], "opp_id": r["opp_id"], "numero": r["numero"], "creato": r["creato"], "data_richiesta": r["data_richiesta"],
         "cliente": r["cliente"], "indirizzo": r["indirizzo"], "email": r["email"], "totale": r["totale"], "inviato": r["inviato"],
         "inviato_a": r["inviato_a"], "servizi": [prev_mod.SERVIZI[x["servizio"]] for x in righe if x.get("servizio") in prev_mod.SERVIZI]}
    if completo:
        d["righe"] = righe
    return d


def _load_prev(pid, u):
    r = db().execute("SELECT * FROM preventivi WHERE id=?", (pid,)).fetchone()
    if not r:
        abort(404)
    load_opp(r["opp_id"], u)      # stessi diritti di accesso dell'opportunità
    return r


@app.get("/api/preventivi/servizi")
@need("preventivi")
def api_prev_servizi():
    return jsonify({"servizi": [{"id": k, "nome": v} for k, v in prev_mod.SERVIZI.items()], "tariffa_extra": prev_mod.TARIFFA_EXTRA})


@app.get("/api/opportunities/<int:oid>/preventivi")
@need("preventivi")
def api_prev_lista(oid):
    load_opp(oid, current_user())
    return jsonify([_prev_dict(r, True) for r in db().execute("SELECT * FROM preventivi WHERE opp_id=? ORDER BY id DESC", (oid,))])


@app.post("/api/opportunities/<int:oid>/preventivi")
@need("preventivi")
def api_prev_crea(oid):
    u = current_user()
    o = load_opp(oid, u)
    c = db()
    d = body()
    if o["categoria"] == "disdetta":
        return err("Per le disdette non si fa il preventivo")
    if o["stato"] in ("contratto", "ko"):
        return err("Opportunità già chiusa")
    righe = d.get("righe")
    if not isinstance(righe, list) or not righe:
        return err("Aggiungi almeno un servizio da quotare")
    if len(righe) > 10:
        return err("Massimo 10 servizi per preventivo")
    pulite = []
    for r in righe:
        x, e = prev_mod.valida_riga(r)
        if e:
            return err(e)
        pulite.append(x)
    data = str(d.get("data_richiesta") or "")[:10]
    try:
        datetime.strptime(data, "%Y-%m-%d")
    except ValueError:
        return err("Indica la data in cui il cliente ha richiesto il preventivo")
    cliente = (d.get("cliente") or o["cliente"] or "").strip()[:200]
    if not cliente:
        return err("Manca il nome del cliente")
    ind = (d.get("indirizzo") or "").strip()[:300]
    email = (d.get("email") or "").strip()[:150]
    tot = round(sum(x["canone"] for x in pulite), 2)
    cur = c.execute("INSERT INTO preventivi(opp_id,autore_id,creato,data_richiesta,cliente,indirizzo,email,righe,totale) VALUES(?,?,?,?,?,?,?,?,?)",
                    (oid, u["id"], ts(), data, cliente, ind, email, json.dumps(pulite), tot))
    pid = cur.lastrowid
    num = "P%s-%04d" % (now().year, pid)
    c.execute("UPDATE preventivi SET numero=? WHERE id=?", (num, pid))
    c.execute("UPDATE opportunities SET stato='preventivo', appuntamento=NULL, con_preventivo=1, gestita=COALESCE(gestita,?) WHERE id=?", (ts(), oid))
    log_event(c, oid, u["id"], "preventivo", "%s: %s, canone mensile %s" % (num, ", ".join(prev_mod.SERVIZI[x["servizio"]].split(" (")[0] for x in pulite), prev_mod.fmt_euro(tot)))
    c.commit()
    return jsonify({"ok": True, "id": pid, "numero": num})


@app.get("/api/preventivi/<int:pid>.pdf")
@need("preventivi")
def api_prev_pdf(pid):
    r = _load_prev(pid, current_user())
    righe = json.loads(r["righe"] or "[]")
    pdf = prev_mod.genera_pdf({"cliente": r["cliente"], "indirizzo": r["indirizzo"], "data_richiesta": r["data_richiesta"], "righe": righe})
    nome = re.sub(r"[^A-Za-z0-9]+", "_", r["cliente"]).strip("_")[:40] or "cliente"
    return send_file(io.BytesIO(pdf), mimetype="application/pdf", as_attachment=not request.args.get("inline"),
                     download_name="Preventivo_%s_%s.pdf" % (r["numero"], nome))


@app.post("/api/preventivi/<int:pid>/invia")
@need("preventivi")
def api_prev_invia(pid):
    u = current_user()
    r = _load_prev(pid, u)
    email = (body().get("email") or "").strip()[:150]
    c = db()
    c.execute("UPDATE preventivi SET inviato=?, inviato_a=?, email=CASE WHEN ?<>'' THEN ? ELSE email END WHERE id=?", (ts(), email, email, email, pid))
    log_event(c, r["opp_id"], u["id"], "preventivo inviato", "%s%s" % (r["numero"], " a " + email if email else ""))
    c.commit()
    return jsonify({"ok": True})


@app.post("/api/opportunities/<int:oid>/assegna")
@need("assegna")
def api_assegna(oid):
    u = current_user()
    o = load_opp(oid, u)
    c = db()
    if o["stato"] in ("contratto", "ko"):
        return err("Opportunità già chiusa")
    aid = body().get("agent_id")
    t = c.execute("SELECT * FROM users WHERE id=? AND attivo=1", (aid,)).fetchone()
    want = ("recuperatore", "agente") if o["categoria"] == "disdetta" else ("agente",)
    if not t or t["id"] not in visible_user_ids(u) or t["ruolo"] not in want or t["area_id"] != o["area_id"]:
        return err("Destinatario non valido (serve %s della stessa area)" % " o ".join(ROLE_LABEL[w].lower() for w in want))
    sla = sla_ore(c, o["categoria"], o["sottotipo"])
    c.execute("UPDATE opportunities SET agente_id=?, stato='da_gestire', appuntamento=NULL, gestita=NULL, assegnata=?, "
              "scadenza=?, sla_notificato=0 WHERE id=?", (t["id"], ts(), ts(now() + timedelta(hours=sla)), oid))
    log_event(c, oid, u["id"], "assegnata", "a %s" % t["nome"])
    notify(c, cc_for_assignment(c, t["id"]), "Opportunità assegnata a %s: %s" % (t["nome"], o["cliente"]), "assegnazione", oid)
    c.commit()
    return jsonify({"ok": True})


# ----------------------------------------------------------------- agenda
def slot_libero(used):
    for s in SLOT_ORARI:
        if s not in used:
            return s
    return "20:00"


def build_proposal(c, agent_id):
    mx = int(get_setting(c, "max_appuntamenti_giorno"))
    opps = c.execute("SELECT id,cliente,localita FROM opportunities WHERE agente_id=? AND stato IN ('da_gestire','in_lavorazione','da_riprogrammare') "
                     "ORDER BY scadenza IS NULL, scadenza, id", (agent_id,)).fetchall()
    groups = {}
    for o in opps:
        groups.setdefault((o["localita"] or "Senza località").strip().title(), []).append(o)
    flat = [(k, o) for k, lst in sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0])) for o in lst][:60]
    booked = {}
    for r in c.execute("SELECT appuntamento FROM opportunities WHERE agente_id=? AND stato='appuntamento' AND appuntamento>=?",
                       (agent_id, ts(now())[:10])):
        booked.setdefault(r[0][:10], []).append(r[0][11:16])
    day = now().date() + timedelta(days=1)
    limit = now().date() + timedelta(days=60)
    out = []
    for loc, o in flat:
        while True:
            while day.weekday() >= 5:
                day += timedelta(days=1)
            used = booked.setdefault(day.isoformat(), [])
            if len(used) < mx:
                break
            day += timedelta(days=1)
        if day > limit:
            break
        ora = slot_libero(used)
        used.append(ora)
        out.append({"id": o["id"], "cliente": o["cliente"], "localita": loc, "data": day.isoformat(), "ora": ora})
    return out


@app.get("/api/agenda")
@need("agenda")
def api_agenda():
    u = current_user()
    agent_id = int(request.args["agent_id"]) if request.args.get("agent_id", "").isdigit() else None
    where, ps = scope_where(u, agent_id)
    today = ts(now())[:10]
    rows = db().execute(
        BASE_SELECT + " WHERE " + where + " AND (o.stato='appuntamento' OR o.stato='da_riprogrammare') "
        "ORDER BY o.appuntamento, o.localita", ps).fetchall()
    giorni, da_ric = {}, []
    for r in rows:
        d = opp_dict(r)
        if r["stato"] == "da_riprogrammare":
            d["motivo"] = "Da riprogrammare"
            da_ric.append(d)
        elif r["appuntamento"][:10] < today:
            d["motivo"] = "Appuntamento passato senza esito"
            da_ric.append(d)
        else:
            giorni.setdefault(r["appuntamento"][:10], []).append(d)
    out = []
    for g_, lst in sorted(giorni.items()):
        lst.sort(key=lambda x: ((x["localita"] or "").lower(), x["appuntamento"]))
        out.append({"data": g_, "items": lst})
    return jsonify({"giorni": out, "da_ricalendarizzare": da_ric,
                    "max_giorno": int(get_setting(db(), "max_appuntamenti_giorno"))})


@app.get("/api/agenda/proposta")
@need("agenda")
def api_proposta():
    u = current_user()
    aid = request.args.get("agent_id", type=int)
    if not aid or aid not in visible_user_ids(u):
        return err("Seleziona un agente", 400)
    return jsonify(build_proposal(db(), aid))


@app.post("/api/agenda/proposta/conferma")
@need("agenda")
def api_proposta_conferma():
    u = current_user()
    items = body().get("items") or []
    c = db()
    ok, errori = 0, []
    for it in items[:60]:
        o = c.execute(BASE_SELECT + " WHERE o.id=?", (it.get("id"),)).fetchone()
        if not o or not can_access(u, o):
            continue
        e = set_appointment(c, o, it.get("data"), it.get("ora"), u)
        if isinstance(e, dict):
            errori.append("%s: %s Fissalo dal suo dettaglio per confermare." % (o["cliente"], e["avviso"]))
        elif e:
            errori.append("%s: %s" % (o["cliente"], e))
        else:
            ok += 1
    c.commit()
    return jsonify({"fissati": ok, "errori": errori})


# ----------------------------------------------------------------- dashboard
def _csv_ints(name):
    return [int(x) for x in (request.args.get(name) or "").split(",") if x.strip().isdigit()]


def _data_arg(name):
    v = (request.args.get(name) or "").strip()
    if not v:
        return None
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
        abort(400)
    try:
        datetime.strptime(v, "%Y-%m-%d")
    except ValueError:
        abort(400)
    return v


def dash_filters(u):
    """Filtri della dashboard e della mappa: area, categoria, più agenti, più campagne, periodo (data di arrivo)."""
    q = request.args
    area_id = q.get("area_id", type=int)
    where, ps = scope_where(u, None, area_id, q.get("categoria"))
    ids = _csv_ints("agent_ids")
    if q.get("agent_id", type=int):
        ids.append(q.get("agent_id", type=int))
    sub = []
    if ids:
        vis = set(visible_user_ids(u))
        if any(i not in vis for i in ids):
            abort(403)
        sub = sorted({x for i in ids for x in subtree_ids(i)})
        where += " AND o.agente_id IN (%s)" % ",".join("?" * len(sub))
        ps += sub
    camp = _csv_ints("campaign_ids")
    if camp:
        where += " AND o.campaign_id IN (%s)" % ",".join("?" * len(camp))
        ps += camp
    dal, al = _data_arg("dal"), _data_arg("al")
    if dal and al and dal > al:
        dal, al = al, dal
    if dal:
        where += " AND substr(COALESCE(o.creata,o.assegnata),1,10)>=?"
        ps.append(dal)
    if al:
        where += " AND substr(COALESCE(o.creata,o.assegnata),1,10)<=?"
        ps.append(al)
    return where, ps, {"area_id": area_id, "agenti": sub, "dal": dal, "al": al}


@app.get("/api/dashboard")
@need("dashboard")
def api_dashboard():
    u = current_user()
    c = db()
    where, ps, fl = dash_filters(u)
    n = ts()
    k = c.execute(
        "SELECT COUNT(*) totale, SUM(o.stato='da_gestire') da_gestire, SUM(o.stato<>'da_gestire') gestite, "
        "SUM(o.stato='appuntamento') appuntamenti, SUM(o.stato='da_riprogrammare') da_riprogrammare, "
        "SUM(o.stato='contratto') contratti, SUM(o.stato='ko') ko, "
        "SUM(o.stato='da_gestire' AND o.scadenza IS NOT NULL AND o.scadenza<?) scadute, "
        "SUM(o.agente_id IS NULL) non_assegnate "
        "FROM opportunities o LEFT JOIN campaigns c ON c.id=o.campaign_id WHERE " + where, [n] + ps).fetchone()
    kpi = {key: (k[key] or 0) for key in k.keys()}
    chiuse = kpi["contratti"] + kpi["ko"]
    kpi["conversione"] = round(100 * kpi["contratti"] / chiuse, 1) if chiuse else None

    fine = datetime.strptime(fl["al"], "%Y-%m-%d").date() if fl["al"] else now().date()
    if fl["dal"]:
        inizio = max(datetime.strptime(fl["dal"], "%Y-%m-%d").date(), fine - timedelta(days=91))
    else:
        inizio = fine - timedelta(days=13)
    fine = max(fine, inizio)
    days = [(inizio + timedelta(days=i)).isoformat() for i in range((fine - inizio).days + 1)]
    g_map = {r[0]: r[1] for r in c.execute(
        "SELECT substr(o.gestita,1,10) d, COUNT(*) FROM opportunities o LEFT JOIN campaigns c ON c.id=o.campaign_id "
        "WHERE " + where + " AND o.gestita>=? AND substr(o.gestita,1,10)<=? GROUP BY d", ps + [days[0], days[-1]])}
    c_map = {r[0]: r[1] for r in c.execute(
        "SELECT substr(e.quando,1,10) d, COUNT(*) FROM events e JOIN opportunities o ON o.id=e.opp_id "
        "LEFT JOIN campaigns c ON c.id=o.campaign_id WHERE " + where + " AND e.azione='contratto' AND e.quando>=? "
        "AND substr(e.quando,1,10)<=? GROUP BY d", ps + [days[0], days[-1]])}
    trend = [{"data": d, "gestite": g_map.get(d, 0), "contratti": c_map.get(d, 0)} for d in days]

    cat = c.execute(
        "SELECT c.categoria, COUNT(*) totale, SUM(o.stato='da_gestire') da_gestire, SUM(o.stato='contratto') contratti, "
        "SUM(o.stato='ko') ko FROM opportunities o LEFT JOIN campaigns c ON c.id=o.campaign_id WHERE " + where +
        " GROUP BY c.categoria", ps).fetchall()
    team = c.execute(
        "SELECT a.id, a.nome, a.ruolo, p.nome AS riferimento, COUNT(o.id) assegnate, SUM(o.stato='da_gestire') da_gestire, "
        "SUM(o.stato<>'da_gestire') gestite, SUM(o.stato='contratto') contratti, SUM(o.stato='ko') ko, "
        "SUM(o.stato='da_gestire' AND o.scadenza IS NOT NULL AND o.scadenza<?) scadute "
        "FROM opportunities o JOIN users a ON a.id=o.agente_id LEFT JOIN users p ON p.id=a.parent_id "
        "LEFT JOIN campaigns c ON c.id=o.campaign_id WHERE " + where + " GROUP BY a.id ORDER BY a.nome", [n] + ps).fetchall()
    w2, p2 = scope_where(u, None, fl["area_id"])
    camps = c.execute(
        "SELECT c.id,c.nome,c.categoria,c.sottotipo,COUNT(o.id) totale FROM campaigns c JOIN opportunities o ON o.campaign_id=c.id "
        "WHERE " + w2 + " GROUP BY c.id ORDER BY c.categoria, c.id DESC", p2).fetchall()
    return jsonify({
        "campagne": [dict(r) for r in camps],
        "periodo": {"dal": days[0], "al": days[-1], "filtrato": bool(fl["dal"] or fl["al"])},
        "kpi": kpi, "trend": trend,
        "categorie": [{k_: (r[k_] or 0) if k_ != "categoria" else r[k_] for k_ in r.keys()} for r in cat],
        "team": [{k_: (r[k_] or 0) if isinstance(r[k_], (int, float, type(None))) else r[k_] for k_ in r.keys()} for r in team]})


@app.get("/api/dashboard/elenco")
@need("dashboard")
def api_dashboard_elenco():
    """Elenco delle opportunità dietro un numero della dashboard (stessi filtri della dashboard)."""
    u = current_user()
    where, ps, _fl = dash_filters(u)
    f = request.args.get("filtro", "totale")
    n = ts()
    extra = {"totale": ("", []), "gestite": (" AND o.stato<>'da_gestire'", []), "da_gestire": (" AND o.stato='da_gestire'", []),
             "scadute": (" AND o.stato='da_gestire' AND o.scadenza IS NOT NULL AND o.scadenza<?", [n]),
             "contratti": (" AND o.stato='contratto'", []), "ko": (" AND o.stato='ko'", []),
             "appuntamenti": (" AND o.stato='appuntamento'", []), "non_assegnate": (" AND o.agente_id IS NULL", [])}
    if f not in extra:
        return err("Filtro non valido")
    cond, pe = extra[f]
    cat = request.args.get("cat")
    if cat in CATEGORIE:
        cond += " AND c.categoria=?"
        pe = pe + [cat]
    ag = request.args.get("agente", type=int)
    if ag:
        cond += " AND o.agente_id=?"
        pe = pe + [ag]
    rows = db().execute(BASE_SELECT + " WHERE " + where + cond + " ORDER BY (o.stato='da_gestire') DESC, o.scadenza IS NULL, o.scadenza, o.id LIMIT 501",
                        ps + pe).fetchall()
    return jsonify({"totale": min(len(rows), 500), "troncato": len(rows) > 500, "items": [opp_dict(r) for r in rows[:500]]})


# ----------------------------------------------------------------- campagne / import
ALIASES = {
    "cliente": ["cliente", "nome", "ragione sociale", "nominativo", "denominazione", "nome cliente"],
    "telefono": ["telefono", "tel", "cellulare", "recapito", "numero"],
    "localita": ["localita", "comune", "citta", "city", "zona"],
    "indirizzo": ["indirizzo", "via"],
    "note": ["note", "nota", "annotazioni"],
    "agente": ["agente", "username agente", "codice agente", "assegnato a", "assegnatario"],
    "piva": ["piva", "p.iva", "p iva", "partita iva", "partita_iva", "p. iva", "vat", "cf", "codice fiscale", "c.f.", "cf/piva", "piva/cf", "cod fiscale"],
}


def _cell(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def read_rows(f):
    name = (f.filename or "").lower()
    data = f.read()
    if name.endswith(".csv"):
        text = data.decode("utf-8-sig", errors="replace")
        delim = ";" if text.count(";") > text.count(",") else ","
        return [[c.strip() for c in r] for r in csv.reader(io.StringIO(text), delimiter=delim)]
    if name.endswith((".xlsx", ".xlsm")):
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        return [[_cell(c) for c in r] for r in wb.active.iter_rows(values_only=True)]
    raise ValueError("Formato non supportato: usa un file .xlsx o .csv")


@app.get("/api/template.xlsx")
@need("campagne")
def api_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "Lista"
    ws.append(["Cliente", "Telefono", "Località", "Indirizzo", "Note", "Agente"])
    ws.append(["Mario Rossi Srl", "3331234567", "Verona", "Via Roma 1", "Rinnovo in scadenza", ""])
    ws.append(["Anna Bianchi", "0421123456", "Treviso", "Via Garibaldi 8", "Lead da sito", ""])
    for col, w in zip("ABCDEF", (28, 16, 18, 28, 32, 20)):
        ws.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="modello_lista_opportunita.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/api/campaigns")
@need("opportunita")
def api_campaigns():
    u = current_user()
    where, ps = scope_where(u)
    rows = db().execute(
        "SELECT c.id,c.nome,c.categoria,c.sottotipo,c.creato,c.regola,ar.nome AS area, COUNT(o.id) totale, "
        "SUM(o.stato='da_gestire') da_gestire, SUM(o.stato IN ('in_lavorazione','preventivo','appuntamento','da_riprogrammare')) in_corso, "
        "SUM(o.stato='contratto') contratti, SUM(o.stato='ko') ko FROM campaigns c "
        "JOIN opportunities o ON o.campaign_id=c.id LEFT JOIN areas ar ON ar.id=c.area_id "
        "WHERE " + where + " GROUP BY c.id ORDER BY c.id DESC", ps).fetchall()
    return jsonify([{k: (r[k] if k in ("nome", "categoria", "sottotipo", "creato", "regola", "area") else (r[k] or 0))
                     for k in r.keys()} for r in rows])


@app.post("/api/campaigns/import")
@need("campagne")
def api_import():
    u = current_user()
    c = db()
    f = request.files.get("file")
    nome = (request.form.get("nome") or "").strip()[:120]
    categoria = request.form.get("categoria")
    sottotipo = request.form.get("sottotipo") or ""
    regola = request.form.get("regola") or get_setting(c, "regola_default")
    if not f or not nome:
        return err("Indica il nome della campagna e il file")
    if categoria not in CATEGORIE or regola not in REGOLE:
        return err("Categoria o regola non valida")
    if categoria == "prospect" and sottotipo not in SOTTOTIPI:
        return err("Per i prospect indica il tipo: lead, nuova nata o lead prospect")
    if categoria != "prospect":
        sottotipo = ""
    area_id = u["area_id"] if u["ruolo"] != "admin" else request.form.get("area_id", type=int)
    if not area_id:
        return err("Seleziona l'area")
    try:
        rows = read_rows(f)
    except ValueError as e:
        return err(str(e))
    except Exception:
        return err("File non leggibile")
    rows = [r for r in rows if any(x for x in r)]
    if len(rows) < 2:
        return err("Il file non contiene righe di dati")
    col = _col_map(rows)
    if "cliente" not in col:
        return err("Colonna 'Cliente' non trovata. Scarica il modello per il formato corretto.")
    if len(rows) - 1 > 5000:
        return err("Massimo 5000 righe per import")

    # controllo duplicati: customer base, opportunità di altre campagne, proposte in attesa, doppioni nel file
    voci = []
    for n_riga, r in enumerate(rows[1:], start=2):
        def gv(k):
            i = col.get(k)
            return r[i].strip() if i is not None and i < len(r) else ""
        voci.append({"riga": n_riga, "cliente": gv("cliente"), "telefono": gv("telefono"), "piva": gv("piva"), "indirizzo": gv("indirizzo"), "localita": gv("localita")})
    fonti_ctrl = ("opp", "prop") if categoria in ("customer_base", "disdetta") else ("cb", "opp", "prop")
    ctrl = dict(zip([v["riga"] for v in voci], _controlla_lista(_indice(c, area_id, fonti=fonti_ctrl, uid=u["id"]), voci)))
    forti = lambda m: dupcheck.migliore(m) in ("certo", "probabile")
    if request.form.get("anteprima") == "1":
        dup = []
        for v in voci:
            m = ctrl.get(v["riga"]) or []
            if m:
                dup.append({"riga": v["riga"], "cliente": v["cliente"], "telefono": v["telefono"], "livello": dupcheck.migliore(m),
                            "dettaglio": dupcheck.riassunto(m), "corrispondenze": _m_json(m)})
        return jsonify({"anteprima": True, "righe": len([v for v in voci if v["cliente"]]), "duplicati": len(dup),
                        "forti": len([d for d in dup if d["livello"] in ("certo", "probabile")]),
                        "elenco": dup[:300], "controllo_customer_base": "cb" in fonti_ctrl})
    salta_dup = request.form.get("duplicati") == "salta"

    ruolo_dest = "recuperatore" if categoria == "disdetta" else "agente"
    pool = c.execute("SELECT id,username,nome,territori FROM users WHERE attivo=1 AND ruolo=? AND area_id=? ORDER BY id",
                     (ruolo_dest, area_id)).fetchall()
    carico = {p["id"]: c.execute("SELECT COUNT(*) FROM opportunities WHERE agente_id=? AND stato='da_gestire'", (p["id"],)).fetchone()[0]
              for p in pool}
    by_key = {}
    for p in pool:
        by_key[norm(p["username"])] = p["id"]
        by_key[norm(p["nome"])] = p["id"]
    terr = {p["id"]: {norm(t) for t in (p["territori"] or "").split(",") if t.strip()} | geo_names_for(c, p["id"]) for p in pool}

    def meno_carico(ids):
        return min(ids, key=lambda i: (carico[i], i))

    sla = sla_ore(c, categoria, sottotipo)
    t0 = now()
    cur = c.execute("INSERT INTO campaigns(nome,categoria,sottotipo,area_id,creato_da,creato,regola) VALUES(?,?,?,?,?,?,?)",
                    (nome, categoria, sottotipo, area_id, u["id"], ts(), regola))
    cid = cur.lastrowid
    res = {"creati": 0, "assegnati": 0, "non_assegnati": 0, "per_agente": {}, "avvisi": [], "duplicati": 0, "saltati": 0}
    nomi = {p["id"]: p["nome"] for p in pool}
    for n_riga, r in enumerate(rows[1:], start=2):
        def g_(k):
            i = col.get(k)
            return r[i].strip() if i is not None and i < len(r) else ""
        cliente = g_("cliente")
        if not cliente:
            res["avvisi"].append("Riga %d senza cliente: saltata" % n_riga)
            continue
        mdup = ctrl.get(n_riga) or []
        if salta_dup and forti(mdup):
            res["saltati"] += 1
            continue
        if mdup:
            res["duplicati"] += 1
        loc = g_("localita")
        aid = None
        if regola != "non_assegnare" and pool:
            if regola == "colonna_agente" and g_("agente"):
                aid = by_key.get(norm(g_("agente")))
                if aid is None:
                    res["avvisi"].append("Riga %d: agente '%s' non trovato, assegnata per carico" % (n_riga, g_("agente")))
            if aid is None and regola == "territorio" and loc:
                match = [i for i in terr if norm(loc) in terr[i]]
                if match:
                    aid = meno_carico(match)
            if aid is None:
                aid = meno_carico(list(carico))
        if aid:
            carico[aid] += 1
            res["assegnati"] += 1
            res["per_agente"][nomi[aid]] = res["per_agente"].get(nomi[aid], 0) + 1
        else:
            res["non_assegnati"] += 1
        c.execute(
            "INSERT INTO opportunities(campaign_id,area_id,cliente,telefono,localita,indirizzo,note,agente_id,stato,assegnata,scadenza,creata,dup_info) "
            "VALUES(?,?,?,?,?,?,?,?,'da_gestire',?,?,?,?)",
            (cid, area_id, cliente[:200], g_("telefono")[:50], loc[:100], g_("indirizzo")[:200],
             (g_("note") + ("\nP.IVA/CF: " + g_("piva") if g_("piva") and not dupcheck.piva_da_testo(g_("note")) else ""))[:1000], aid,
             ts(t0) if aid else None, ts(t0 + timedelta(hours=sla)) if aid else None, ts(t0), dupcheck.riassunto(mdup)[:600]))
        res["creati"] += 1
    if res["creati"] == 0:
        c.rollback()
        return err("Nessuna riga valida nel file" if not res["saltati"] else
                   "Tutte le %d righe risultano già presenti in altre liste: nessuna campagna creata" % res["saltati"])
    c.execute("UPDATE campaigns SET totale=? WHERE id=?", (res["creati"], cid))
    # notifica agente destinatario con District e Team Leader in conoscenza
    for aid_name, n in res["per_agente"].items():
        aid = next(i for i, nm in nomi.items() if nm == aid_name)
        notify(c, cc_for_assignment(c, aid), "Campagna «%s»: %d opportunità assegnate a %s" % (nome, n, aid_name), "assegnazione")
    c.commit()
    resolve_missing(c)
    res["campagna_id"] = cid
    return jsonify(res)



# ----------------------------------------------------------------- verifica duplicati e customer base
def _dmy(s):
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return ""


def _area_per(u):
    if u["ruolo"] == "admin":
        a = request.values.get("area_id", type=int)
        if not a and request.is_json:
            a = body().get("area_id")
        return int(a) if a else None
    return u["area_id"]


_NM_RE = re.compile(r"(?:\u27e6|\\u27e6)(.*?)(?:\u27e7|\\u27e7)")
RUOLI_VEDONO_NOMI = ("admin", "direttore", "area_manager", "district")


def _nm(nome):
    """Nome di un venditore dentro un avviso di duplicato: lo vedono solo direttore, area manager, district (e admin)."""
    return "\u27e6%s\u27e7" % (nome or "")


def _redigi(testo, ruolo):
    if not testo:
        return testo
    if ruolo in RUOLI_VEDONO_NOMI:
        return _NM_RE.sub(lambda m: m.group(1), testo)
    return _NM_RE.sub("un altro venditore", testo)


def _indice(c, area_id, fonti=("cb", "opp", "prop"), uid=None):
    """Indice di tutto ciò che nell'area è già noto: customer base, opportunità (qualsiasi campagna), proposte in attesa."""
    ix = dupcheck.Indice()
    if not area_id:
        return ix
    if "cb" in fonti:
        ult = c.execute("SELECT quando FROM cb_batches WHERE area_id=? ORDER BY id DESC LIMIT 1", (area_id,)).fetchone()
        info = "Già cliente (customer base%s)" % (" aggiornata il " + _dmy(ult["quando"]) if ult else "")
        for r in c.execute("SELECT nome,ident,telefono,indirizzo,localita FROM customer_base WHERE area_id=?", (area_id,)):
            ix.add("cb", r["nome"], r["ident"], r["telefono"], info, ind=r["indirizzo"], loc=r["localita"])
    if "opp" in fonti:
        for r in c.execute("SELECT o.id,o.cliente,o.telefono,o.indirizzo,o.localita,o.note,o.stato,o.agente_id,k.nome AS campagna,k.categoria,u.nome AS agente "
                           "FROM opportunities o LEFT JOIN campaigns k ON k.id=o.campaign_id LEFT JOIN users u ON u.id=o.agente_id "
                           "WHERE o.area_id=?", (area_id,)):
            aperta = r["stato"] in ("da_gestire", "in_lavorazione", "preventivo", "appuntamento", "da_riprogrammare")
            stato = r["stato"].replace("_", " ")
            if aperta and r["agente"] and r["categoria"] != "customer_base":
                info = "IN LAVORAZIONE da %s (%s «%s», stato %s)" % (_nm(r["agente"]), "prospect" if r["categoria"] == "prospect" else "campagna", r["campagna"] or "?", stato)
            elif aperta:
                info = "In campagna «%s», non ancora assegnata (stato %s)" % (r["campagna"] or "?", stato)
            else:
                info = "Già trattato%s in campagna «%s»: %s" % (" da " + _nm(r["agente"]) if r["agente"] else "", r["campagna"] or "?", stato)
            ix.add("opp", r["cliente"], dupcheck.piva_da_testo(r["note"]), r["telefono"], info,
                   {"id": r["id"], "agente_id": r["agente_id"], "lavorazione": bool(aperta and r["agente_id"] and r["categoria"] != "customer_base")},
                   ind=r["indirizzo"], loc=r["localita"])
    if "prop" in fonti:
        for r in c.execute("SELECT p.id,p.cliente,p.telefono,p.indirizzo,p.localita,p.note,p.proponente_id,p.creata,u.nome AS proponente FROM proposals p "
                           "JOIN users u ON u.id=p.proponente_id WHERE p.area_id=? AND p.stato='in_attesa'", (area_id,)):
            ix.add("prop", r["cliente"], dupcheck.piva_da_testo(r["note"]), r["telefono"],
                   "Già proposto da %s il %s (in attesa del Direttore)" % (_nm(r["proponente"]), _dmy(r["creata"])), {"id": r["id"], "proponente_id": r["proponente_id"]}, ind=r["indirizzo"], loc=r["localita"])
    return ix


def _col_map(rows):
    hdr = [norm(h) for h in rows[0]]
    col = {}
    for key, al in ALIASES.items():
        for i, h in enumerate(hdr):
            if h in al:
                col[key] = i
                break
    return col


def _leggi_lista(f, max_righe=5000):
    """File Excel/CSV -> elenco di dict (riga, cliente, telefono, localita, indirizzo, piva)."""
    rows = [r for r in read_rows(f) if any(x for x in r)]
    if len(rows) < 2:
        raise ValueError("Il file non contiene righe di dati")
    col = _col_map(rows)
    if "cliente" not in col:
        raise ValueError("Colonna 'Cliente' non trovata (serve una colonna con il nome del cliente, es. «Cliente» o «Ragione sociale»)")
    if len(rows) - 1 > max_righe:
        raise ValueError("Massimo %d righe per file" % max_righe)
    out = []
    for n, r in enumerate(rows[1:], start=2):
        def g_(k):
            i = col.get(k)
            return r[i].strip() if i is not None and i < len(r) else ""
        out.append({"riga": n, "cliente": g_("cliente"), "telefono": g_("telefono"), "localita": g_("localita"),
                    "indirizzo": g_("indirizzo"), "piva": g_("piva")})
    return out


def _righe_testo(t, max_righe=2000):
    """Righe incollate: «Nome; telefono; partita IVA» (anche solo il nome)."""
    out = []
    for n, ln in enumerate(str(t or "").splitlines(), start=1):
        ln = ln.strip()
        if not ln:
            continue
        if re.search(r"[;\t|]", ln):
            parti = [p.strip() for p in re.split(r"[;\t|]", ln)]
        elif "," in ln and any(len(re.sub(r"\D", "", x)) >= 6 for x in ln.split(",")[1:]):
            parti = [p.strip() for p in ln.split(",")]
        else:
            parti = [ln]
        nome, tel, piva = parti[0], "", ""
        for p in parti[1:]:
            if dupcheck.norm_ident(p) and len(re.sub(r"[^A-Za-z0-9]", "", p)) >= 11:
                piva = p
            elif re.sub(r"\D", "", p) and len(re.sub(r"\D", "", p)) >= 6:
                tel = p
        out.append({"riga": n, "cliente": nome, "telefono": tel, "localita": "", "indirizzo": "", "piva": piva})
        if len(out) > max_righe:
            raise ValueError("Massimo %d righe per volta" % max_righe)
    return out


def _controlla_lista(ix, voci, doppi_nel_file=True):
    """Per ogni voce cerca i duplicati nell'indice e dentro la stessa lista. Aggiunge le voci all'indice."""
    ris = []
    for v in voci:
        m = ix.cerca(v["cliente"], v["piva"], v["telefono"], ind=v.get("indirizzo", ""), loc=v.get("localita", "")) if v["cliente"] else []
        if doppi_nel_file and v["cliente"]:
            ix.add("file", v["cliente"], v["piva"], v["telefono"], "Doppio nel file (riga %d)" % v["riga"], {"riga": v["riga"]}, ind=v.get("indirizzo", ""), loc=v.get("localita", ""))
        ris.append(m)
    return ris


def _m_json(m):
    return [{"livello": x["livello"], "motivo": x["motivo"], "fonte": x["fonte"], "info": x["info"], "nome": x["nome"],
             "lavorazione": bool((x.get("meta") or {}).get("lavorazione"))} for x in m]


def _check_voci(c, area_id, voci, uid):
    ix = _indice(c, area_id, uid=uid)
    ms = _controlla_lista(ix, voci)
    righe, n_dup = [], 0
    for v, m in zip(voci, ms):
        liv = dupcheck.migliore(m)
        if liv:
            n_dup += 1
        lav = any((x.get("meta") or {}).get("lavorazione") and x["livello"] != "simile" for x in m)
        forse = (not lav) and any((x.get("meta") or {}).get("lavorazione") for x in m)
        righe.append({"riga": v["riga"], "cliente": v["cliente"], "telefono": v["telefono"], "piva": v["piva"], "livello": liv,
                      "in_lavorazione": lav, "forse_lavorazione": forse, "dettaglio": dupcheck.riassunto(m), "corrispondenze": _m_json(m)})
    return righe, n_dup


def _voci_da_richiesta():
    f = request.files.get("file")
    if f and f.filename:
        return _leggi_lista(f, 2000)
    t = (request.form.get("testo") or (body().get("testo") if request.is_json else "") or "")
    if len(t.strip()) < 2:
        raise ValueError("Carica un file Excel/CSV oppure incolla l'elenco dei clienti da verificare")
    return _righe_testo(t)


@app.post("/api/verifica")
@need("verifica")
def api_verifica():
    u = current_user()
    area = _area_per(u)
    if not area:
        return err("Seleziona l'area" if u["ruolo"] == "admin" else "Il tuo profilo non appartiene a nessuna area")
    try:
        voci = _voci_da_richiesta()
    except ValueError as e:
        return err(str(e))
    except Exception:
        return err("File non leggibile")
    righe, n_dup = _check_voci(db(), area, voci, u["id"])
    return jsonify({"totale": len(righe), "duplicati": n_dup, "in_lavorazione": sum(1 for r in righe if r["in_lavorazione"]), "righe": righe,
                    "customer_base": db().execute("SELECT COUNT(*) FROM customer_base WHERE area_id=?", (area,)).fetchone()[0]})


@app.post("/api/verifica.xlsx")
@need("verifica")
def api_verifica_xlsx():
    u = current_user()
    area = _area_per(u)
    if not area:
        return err("Seleziona l'area")
    try:
        voci = _voci_da_richiesta()
    except ValueError as e:
        return err(str(e))
    righe, _ = _check_voci(db(), area, voci, u["id"])
    wb = Workbook()
    ws = wb.active
    ws.title = "Verifica"
    ws.append(["Riga", "Cliente", "Telefono", "P.IVA / CF", "Esito", "Dettaglio"])
    for r in righe:
        ws.append([r["riga"], r["cliente"], r["telefono"], r["piva"],
                   ("IN LAVORAZIONE - " if r["in_lavorazione"] else "") + (dupcheck.ETICHETTA[r["livello"]] if r["livello"] else "Nuovo"), _redigi(r["dettaglio"], u["ruolo"])])
    for col, w in zip("ABCDEF", (7, 34, 16, 18, 26, 90)):
        ws.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="verifica_clienti.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# --- customer base (elenco completo clienti dell'area)
def _cb_chiave(nome, ident, tel):
    idn = dupcheck.norm_ident(ident)
    if idn:
        return "I:" + idn
    t = sorted(dupcheck.norm_tels(tel))
    return "N:%s|%s" % (dupcheck.norm_nome(nome), t[0] if t else "")


@app.get("/api/customerbase")
@need("customerbase")
def api_cb():
    u = current_user()
    area = _area_per(u)
    if not area:
        return err("Seleziona l'area")
    c = db()
    tot = c.execute("SELECT COUNT(*) FROM customer_base WHERE area_id=?", (area,)).fetchone()[0]
    storia = c.execute("SELECT b.*, u.nome AS utente FROM cb_batches b LEFT JOIN users u ON u.id=b.user_id WHERE b.area_id=? "
                       "ORDER BY b.id DESC LIMIT 12", (area,)).fetchall()
    return jsonify({"area_id": area, "totale": tot, "storia": [dict(r) for r in storia]})


@app.get("/api/customerbase/modello.xlsx")
@need("customerbase")
def api_cb_modello():
    wb = Workbook()
    ws = wb.active
    ws.title = "Customer base"
    ws.append(["Cliente", "P.IVA / Codice fiscale", "Telefono", "Località", "Indirizzo"])
    ws.append(["Mario Rossi Srl", "01234567890", "0421123456", "Treviso", "Via Roma 1"])
    for col, w in zip("ABCDE", (34, 22, 16, 18, 30)):
        ws.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="modello_customer_base.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/api/customerbase/esporta.xlsx")
@need("customerbase")
def api_cb_esporta():
    u = current_user()
    area = _area_per(u)
    if not area:
        return err("Seleziona l'area")
    wb = Workbook()
    ws = wb.active
    ws.title = "Customer base"
    ws.append(["Cliente", "P.IVA / Codice fiscale", "Telefono", "Località", "Indirizzo"])
    for r in db().execute("SELECT nome,ident,telefono,localita,indirizzo FROM customer_base WHERE area_id=? ORDER BY nome", (area,)):
        ws.append([r["nome"], r["ident"], r["telefono"], r["localita"], r["indirizzo"]])
    for col, w in zip("ABCDE", (34, 22, 16, 18, 30)):
        ws.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="customer_base_attuale.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/customerbase/anteprima")
@need("customerbase")
def api_cb_anteprima():
    """Legge il file e mostra cosa cambierebbe rispetto all'elenco attuale, senza modificare nulla."""
    u = current_user()
    area = _area_per(u)
    if not area:
        return err("Seleziona l'area")
    f = request.files.get("file")
    if not f or not f.filename:
        return err("Scegli il file della customer base (.xlsx o .csv)")
    try:
        voci = _leggi_lista(f, 200000)
    except ValueError as e:
        return err(str(e))
    except Exception:
        return err("File non leggibile")
    c = db()
    nuovo, doppi, senza_nome, deboli = {}, 0, 0, 0
    for v in voci:
        if not v["cliente"]:
            senza_nome += 1
            continue
        if not dupcheck.norm_ident(v["piva"]) and not dupcheck.norm_tels(v["telefono"]):
            deboli += 1
        k = _cb_chiave(v["cliente"], v["piva"], v["telefono"])
        if k in nuovo:
            doppi += 1
            continue
        nuovo[k] = {"nome": v["cliente"][:200], "ident": dupcheck.norm_ident(v["piva"]), "telefono": v["telefono"][:60],
                    "localita": v["localita"][:100], "indirizzo": v["indirizzo"][:200]}
    if not nuovo:
        return err("Nessun cliente valido nel file")
    att = {r["chiave"]: r for r in c.execute("SELECT * FROM customer_base WHERE area_id=?", (area,))}
    n_nuovi = [k for k in nuovo if k not in att]
    n_rim = [k for k in att if k not in nuovo]
    agg = [k for k in nuovo if k in att and (att[k]["nome"] != nuovo[k]["nome"] or att[k]["telefono"] != nuovo[k]["telefono"] or
                                            att[k]["localita"] != nuovo[k]["localita"] or att[k]["indirizzo"] != nuovo[k]["indirizzo"])]
    inv = len(nuovo) - len(n_nuovi) - len(agg)
    tok = secrets.token_urlsafe(16)
    c.execute("DELETE FROM cb_staging WHERE quando<?", (ts(now() - timedelta(days=1)),))
    c.execute("INSERT INTO cb_staging(token,area_id,user_id,dati,file,quando) VALUES(?,?,?,?,?,?)",
              (tok, area, u["id"], json.dumps(nuovo), (f.filename or "")[:120], ts()))
    c.commit()
    return jsonify({"token": tok, "file": f.filename, "totale": len(nuovo), "attuale": len(att), "nuovi": len(n_nuovi), "rimossi": len(n_rim),
                    "aggiornati": len(agg), "invariati": inv, "doppi_nel_file": doppi, "senza_nome": senza_nome, "senza_identificativo": deboli,
                    "esempio_nuovi": [nuovo[k]["nome"] for k in n_nuovi[:8]], "esempio_rimossi": [att[k]["nome"] for k in n_rim[:8]]})


@app.post("/api/customerbase/conferma")
@need("customerbase")
def api_cb_conferma():
    u = current_user()
    c = db()
    d = body()
    st = c.execute("SELECT * FROM cb_staging WHERE token=?", (str(d.get("token") or ""),)).fetchone()
    if not st or (u["ruolo"] != "admin" and st["user_id"] != u["id"]):
        return err("Anteprima scaduta: ricarica il file")
    modo = "aggiungi" if d.get("modo") == "aggiungi" else "sostituisci"
    area = st["area_id"]
    nuovo = json.loads(st["dati"])
    att = {r["chiave"]: r for r in c.execute("SELECT * FROM customer_base WHERE area_id=?", (area,))}
    cur = c.execute("INSERT INTO cb_batches(area_id,user_id,quando,file,modo) VALUES(?,?,?,?,?)", (area, u["id"], ts(), st["file"], modo))
    bid = cur.lastrowid
    n = a = inv = 0
    for k, v in nuovo.items():
        if k not in att:
            c.execute("INSERT INTO customer_base(area_id,nome,ident,telefono,localita,indirizzo,chiave,batch_id) VALUES(?,?,?,?,?,?,?,?)",
                      (area, v["nome"], v["ident"], v["telefono"], v["localita"], v["indirizzo"], k, bid))
            n += 1
        else:
            o = att[k]
            if (o["nome"], o["telefono"], o["localita"], o["indirizzo"]) != (v["nome"], v["telefono"], v["localita"], v["indirizzo"]):
                c.execute("UPDATE customer_base SET nome=?,telefono=?,localita=?,indirizzo=?,batch_id=? WHERE id=?",
                          (v["nome"], v["telefono"], v["localita"], v["indirizzo"], bid, o["id"]))
                a += 1
            else:
                inv += 1
    rim = 0
    if modo == "sostituisci":
        for k, o in att.items():
            if k not in nuovo:
                c.execute("DELETE FROM customer_base WHERE id=?", (o["id"],))
                rim += 1
    tot = c.execute("SELECT COUNT(*) FROM customer_base WHERE area_id=?", (area,)).fetchone()[0]
    c.execute("UPDATE cb_batches SET totale=?,nuovi=?,rimossi=?,aggiornati=?,invariati=? WHERE id=?", (tot, n, rim, a, inv, bid))
    c.execute("DELETE FROM cb_staging WHERE token=?", (st["token"],))
    ric = _ricontrolla_area(c, area)
    dest = [x for x in _direttori(c, area) if x != u["id"]]
    if dest and (ric["proposte"] or ric["opportunita"]):
        notify(c, dest, "Customer base aggiornata: %d proposte e %d opportunità prospect risultano ora già clienti" % (ric["proposte"], ric["opportunita"]), "info")
    log_event(c, None, u["id"], "customer base", "area %s: +%d -%d ~%d" % (area, n, rim, a))
    c.commit()
    return jsonify({"ok": True, "totale": tot, "nuovi": n, "rimossi": rim, "aggiornati": a, "invariati": inv, "ricontrollo": ric})


def _togli_cb(testo):
    return " | ".join(p for p in str(testo or "").split(" | ") if "Già cliente" not in p)


def _ricontrolla_area(c, area):
    """Dopo l'aggiornamento della customer base: rifà i controlli sulle proposte in attesa e sulle opportunità prospect ancora aperte."""
    ix = _indice(c, area)
    cb = _indice(c, area, fonti=("cb",))
    ris = {"proposte": 0, "opportunita": 0, "elenco": []}
    for p in c.execute("SELECT id,cliente,telefono,indirizzo,localita,note,dup_info FROM proposals WHERE area_id=? AND stato='in_attesa'", (area,)).fetchall():
        m = ix.cerca(p["cliente"], dupcheck.piva_da_testo(p["note"]), p["telefono"], ind=p["indirizzo"], loc=p["localita"],
                     salta=lambda v, i=p["id"]: v["fonte"] == "prop" and v["meta"].get("id") == i)
        nuovo = dupcheck.riassunto(m)[:600]
        if "Già cliente" in nuovo and "Già cliente" not in (p["dup_info"] or ""):
            ris["proposte"] += 1
            ris["elenco"].append(p["cliente"])
        c.execute("UPDATE proposals SET dup_info=? WHERE id=?", (nuovo, p["id"]))
    for o in c.execute("SELECT o.id,o.cliente,o.telefono,o.indirizzo,o.localita,o.note,o.dup_info FROM opportunities o JOIN campaigns k ON k.id=o.campaign_id "
                       "WHERE o.area_id=? AND k.categoria='prospect' AND o.stato IN ('da_gestire','in_lavorazione','preventivo','appuntamento','da_riprogrammare')",
                       (area,)).fetchall():
        m = [x for x in cb.cerca(o["cliente"], dupcheck.piva_da_testo(o["note"]), o["telefono"], ind=o["indirizzo"], loc=o["localita"]) if x["livello"] in ("certo", "probabile", "indirizzo")]
        base = _togli_cb(o["dup_info"])
        nuovo = " | ".join(x for x in (base, dupcheck.riassunto(m[:1])) if x)[:600]
        if m and "Già cliente" not in (o["dup_info"] or ""):
            ris["opportunita"] += 1
            ris["elenco"].append(o["cliente"])
        if nuovo != (o["dup_info"] or ""):
            c.execute("UPDATE opportunities SET dup_info=? WHERE id=?", (nuovo, o["id"]))
    ris["elenco"] = ris["elenco"][:30]
    return ris



# ----------------------------------------------------------------- messaggi con allegati e documenti
UPLOAD_DIR = os.path.join(DATA_DIR, "allegati")
MAX_FILE = 15 * 1024 * 1024
MAX_ALLEGATI = 10
QUOTA_PERSONALE = 300 * 1024 * 1024
EXT_VIETATE = {"exe", "bat", "cmd", "com", "scr", "js", "mjs", "vbs", "vbe", "ps1", "msi", "jar", "html", "htm", "xhtml", "svg", "hta", "dll",
               "lnk", "reg", "sh", "apk", "wsf", "php", "py"}
INLINE_OK = {"image/png", "image/jpeg", "image/gif", "image/webp", "application/pdf"}


def _nome_file(n):
    n = os.path.basename(str(n or "").replace("\\", "/"))
    n = re.sub(r"[\x00-\x1f<>:\"|?*]", "_", n).strip(" .")
    return (n or "file")[:150]


def _salva_upload(c, fs, ambito, owner_id, area_id, message_id=None, cartella=""):
    """Salva i file caricati. Restituisce (n_salvati, errori)."""
    ok, errori = 0, []
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    for f in fs:
        if not f or not f.filename:
            continue
        nome = _nome_file(f.filename)
        ext = nome.rsplit(".", 1)[-1].lower() if "." in nome else ""
        if ext in EXT_VIETATE:
            errori.append("%s: tipo di file non consentito (.%s)" % (nome, ext))
            continue
        dati = f.read(MAX_FILE + 1)
        if len(dati) > MAX_FILE:
            errori.append("%s: supera i %d MB" % (nome, MAX_FILE // 1048576))
            continue
        if not dati:
            errori.append("%s: file vuoto" % nome)
            continue
        if ambito == "personale":
            usato = c.execute("SELECT COALESCE(SUM(dimensione),0) FROM files WHERE ambito='personale' AND owner_id=?", (owner_id,)).fetchone()[0]
            if usato + len(dati) > QUOTA_PERSONALE:
                errori.append("%s: spazio personale esaurito (%d MB)" % (nome, QUOTA_PERSONALE // 1048576))
                continue
        rel = os.path.join(uuid.uuid4().hex[:2], uuid.uuid4().hex)
        path = os.path.join(UPLOAD_DIR, rel)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as fh:
            fh.write(dati)
        mime = mimetypes.guess_type(nome)[0] or "application/octet-stream"
        c.execute("INSERT INTO files(ambito,owner_id,area_id,message_id,cartella,nome,mime,dimensione,percorso,creato) VALUES(?,?,?,?,?,?,?,?,?,?)",
                  (ambito, owner_id, area_id, message_id, cartella[:60], nome, mime, len(dati), rel.replace("\\", "/"), ts()))
        ok += 1
    return ok, errori


def _file_dict(r):
    return {"id": r["id"], "nome": r["nome"], "mime": r["mime"], "dimensione": r["dimensione"], "cartella": r["cartella"], "creato": r["creato"]}


def _puo_vedere_file(c, u, f):
    if u["ruolo"] == "admin" or f["owner_id"] == u["id"]:
        return True
    if f["ambito"] == "comune":
        return f["area_id"] is None or f["area_id"] == u["area_id"]
    if f["ambito"] == "messaggio":
        m = c.execute("SELECT mittente_id FROM messages WHERE id=?", (f["message_id"],)).fetchone()
        if m and m["mittente_id"] == u["id"]:
            return True
        return bool(c.execute("SELECT 1 FROM message_rcpt WHERE message_id=? AND user_id=?", (f["message_id"], u["id"])).fetchone())
    return False


def _elimina_file_fisico(c, f):
    altri = c.execute("SELECT COUNT(*) FROM files WHERE percorso=? AND id<>?", (f["percorso"], f["id"])).fetchone()[0]
    c.execute("DELETE FROM files WHERE id=?", (f["id"],))
    if not altri:
        try:
            os.remove(os.path.join(UPLOAD_DIR, f["percorso"]))
        except OSError:
            pass


@app.get("/api/file/<int:fid>")
@need()
def api_file(fid):
    u = current_user()
    c = db()
    f = c.execute("SELECT * FROM files WHERE id=?", (fid,)).fetchone()
    if not f or not _puo_vedere_file(c, u, f):
        abort(404)
    path = os.path.join(UPLOAD_DIR, f["percorso"])
    if not os.path.isfile(path):
        abort(404)
    inline = request.args.get("anteprima") == "1" and f["mime"] in INLINE_OK
    return send_file(path, mimetype=f["mime"] if inline else "application/octet-stream", as_attachment=not inline, download_name=f["nome"])


@app.post("/api/file/<int:fid>/salva")
@need("documenti")
def api_file_salva(fid):
    """Copia un allegato (o un documento comune) nei propri documenti."""
    u = current_user()
    c = db()
    f = c.execute("SELECT * FROM files WHERE id=?", (fid,)).fetchone()
    if not f or not _puo_vedere_file(c, u, f):
        abort(404)
    usato = c.execute("SELECT COALESCE(SUM(dimensione),0) FROM files WHERE ambito='personale' AND owner_id=?", (u["id"],)).fetchone()[0]
    if usato + f["dimensione"] > QUOTA_PERSONALE:
        return err("Spazio personale esaurito")
    c.execute("INSERT INTO files(ambito,owner_id,area_id,cartella,nome,mime,dimensione,percorso,creato) VALUES('personale',?,?,?,?,?,?,?,?)",
              (u["id"], u["area_id"], "", f["nome"], f["mime"], f["dimensione"], f["percorso"], ts()))
    c.commit()
    return jsonify({"ok": True})


# --- documenti personali e comuni
@app.get("/api/documenti")
@need("documenti")
def api_documenti():
    u = current_user()
    c = db()
    mie = c.execute("SELECT * FROM files WHERE ambito='personale' AND owner_id=? ORDER BY id DESC", (u["id"],)).fetchall()
    if u["ruolo"] == "admin":
        com = c.execute("SELECT f.*, a.nome AS area, u.nome AS autore FROM files f LEFT JOIN areas a ON a.id=f.area_id LEFT JOIN users u ON u.id=f.owner_id "
                        "WHERE f.ambito='comune' ORDER BY f.id DESC").fetchall()
    else:
        com = c.execute("SELECT f.*, a.nome AS area, u.nome AS autore FROM files f LEFT JOIN areas a ON a.id=f.area_id LEFT JOIN users u ON u.id=f.owner_id "
                        "WHERE f.ambito='comune' AND (f.area_id IS NULL OR f.area_id=?) ORDER BY f.id DESC", (u["area_id"],)).fetchall()
    usato = sum(r["dimensione"] for r in mie)
    perm = perms_of(c, u)
    return jsonify({"personali": [_file_dict(r) for r in mie],
                    "comuni": [dict(_file_dict(r), area=r["area"] or "Tutte le aree", autore=r["autore"] or "", mio=(r["owner_id"] == u["id"]),
                                    eliminabile=(u["ruolo"] == "admin" or r["owner_id"] == u["id"] or (u["ruolo"] == "direttore" and r["area_id"] == u["area_id"]))) for r in com],
                    "puo_pubblicare": "documenti_comuni" in perm, "usato": usato, "quota": QUOTA_PERSONALE, "max_file": MAX_FILE,
                    "aree": [dict(r) for r in c.execute("SELECT id,nome FROM areas ORDER BY id")] if u["ruolo"] == "admin" else []})


@app.post("/api/documenti")
@need("documenti")
def api_documenti_carica():
    u = current_user()
    c = db()
    ambito = "comune" if request.form.get("ambito") == "comune" else "personale"
    area = u["area_id"]
    if ambito == "comune":
        if "documenti_comuni" not in perms_of(c, u):
            return err("Non puoi pubblicare nell'area comune", 403)
        if u["ruolo"] == "admin":
            a = request.form.get("area_id", type=int)
            area = a or None
        elif not area:
            return err("Il tuo profilo non appartiene a nessuna area")
    fs = request.files.getlist("file")
    if not fs or not any(f.filename for f in fs):
        return err("Scegli almeno un file")
    ok, errori = _salva_upload(c, fs, ambito, u["id"], area, cartella=(request.form.get("cartella") or "").strip())
    c.commit()
    if not ok:
        return err("; ".join(errori) or "Nessun file salvato")
    if ambito == "comune" and ok:
        if area is None:
            dest = [r["id"] for r in c.execute("SELECT id FROM users WHERE attivo=1 AND id<>?", (u["id"],))]
        else:
            dest = [r["id"] for r in c.execute("SELECT id FROM users WHERE attivo=1 AND id<>? AND area_id=?", (u["id"], area))]
        notify(c, dest, "Nuovi documenti nell'area comune (%d) da %s" % (ok, u["nome"]), "messaggio")
        c.commit()
    return jsonify({"ok": True, "salvati": ok, "errori": errori})


@app.delete("/api/documenti/<int:fid>")
@need("documenti")
def api_documenti_elimina(fid):
    u = current_user()
    c = db()
    f = c.execute("SELECT * FROM files WHERE id=? AND ambito IN ('personale','comune')", (fid,)).fetchone()
    if not f:
        abort(404)
    ok = u["ruolo"] == "admin" or f["owner_id"] == u["id"] or (f["ambito"] == "comune" and u["ruolo"] == "direttore" and f["area_id"] == u["area_id"])
    if not ok:
        abort(403)
    _elimina_file_fisico(c, f)
    c.commit()
    return jsonify({"ok": True})


# --- messaggi
def _dest_ammessi(c, u):
    ids = [i for i in visible_user_ids(u) if i != u["id"]]
    if not ids:
        return []
    return c.execute("SELECT id,nome,ruolo,area_id,parent_id FROM users WHERE attivo=1 AND ruolo<>'admin' AND id IN (%s) ORDER BY nome" % ",".join("?" * len(ids)), ids).fetchall()


@app.get("/api/messaggi/destinatari")
@need("messaggi_invia")
def api_msg_dest():
    c = db()
    u = current_user()
    nomi = {r["id"]: r["nome"] for r in c.execute("SELECT id,nome FROM users")}
    return jsonify([{"id": r["id"], "nome": r["nome"], "ruolo": r["ruolo"], "ruolo_label": ROLE_LABEL[r["ruolo"]], "riferimento": nomi.get(r["parent_id"], "")}
                    for r in _dest_ammessi(c, u)])


@app.post("/api/messaggi")
@need("messaggi_invia")
def api_msg_invia():
    u = current_user()
    c = db()
    titolo = (request.form.get("titolo") or "").strip()[:150]
    testo = (request.form.get("testo") or "").strip()[:8000]
    if not titolo:
        return err("Scrivi l'oggetto del messaggio")
    fs = [f for f in request.files.getlist("file") if f and f.filename]
    if not testo and not fs:
        return err("Scrivi il testo o allega almeno un file")
    if len(fs) > MAX_ALLEGATI:
        return err("Massimo %d allegati per messaggio" % MAX_ALLEGATI)
    amm = _dest_ammessi(c, u)
    modo = request.form.get("modo") or "tutti"
    if modo == "scelti":
        try:
            scelti = {int(x) for x in json.loads(request.form.get("ids") or "[]")}
        except (ValueError, TypeError):
            scelti = set()
        dest = [r["id"] for r in amm if r["id"] in scelti]
    elif modo in ("agente", "recuperatore", "team_leader", "district", "area_manager"):
        dest = [r["id"] for r in amm if r["ruolo"] == modo]
    else:
        dest = [r["id"] for r in amm if r["ruolo"] in ("agente", "recuperatore", "team_leader", "district", "area_manager")]
    if not dest:
        return err("Nessun destinatario: scegli a chi inviare il messaggio")
    cur = c.execute("INSERT INTO messages(mittente_id,area_id,titolo,testo,creato) VALUES(?,?,?,?,?)", (u["id"], u["area_id"], titolo, testo, ts()))
    mid = cur.lastrowid
    ok, errori = _salva_upload(c, fs, "messaggio", u["id"], u["area_id"], message_id=mid)
    if fs and not ok:
        c.rollback()
        return err("Allegati non salvati: " + "; ".join(errori))
    for d in dest:
        c.execute("INSERT INTO message_rcpt(message_id,user_id) VALUES(?,?)", (mid, d))
    notify(c, dest, "📣 %s: %s%s" % (u["nome"], titolo, " (con %d allegati)" % ok if ok else ""), "messaggio")
    c.commit()
    return jsonify({"ok": True, "id": mid, "destinatari": len(dest), "allegati": ok, "errori": errori})


@app.get("/api/messaggi")
@need("messaggi")
def api_msg_lista():
    u = current_user()
    c = db()
    if request.args.get("cartella") == "inviati":
        rows = c.execute("SELECT m.id,m.titolo,m.testo,m.creato,(SELECT COUNT(*) FROM message_rcpt r WHERE r.message_id=m.id) AS n_dest, "
                         "(SELECT COUNT(*) FROM message_rcpt r WHERE r.message_id=m.id AND r.letto IS NOT NULL) AS n_letti, "
                         "(SELECT COUNT(*) FROM files f WHERE f.message_id=m.id) AS n_all FROM messages m WHERE m.mittente_id=? ORDER BY m.id DESC LIMIT 200", (u["id"],)).fetchall()
        return jsonify([{"id": r["id"], "titolo": r["titolo"], "anteprima": r["testo"][:120], "creato": r["creato"], "n_dest": r["n_dest"],
                         "n_letti": r["n_letti"], "n_allegati": r["n_all"]} for r in rows])
    rows = c.execute("SELECT m.id,m.titolo,m.testo,m.creato,u.nome AS mittente,r.letto, (SELECT COUNT(*) FROM files f WHERE f.message_id=m.id) AS n_all "
                     "FROM message_rcpt r JOIN messages m ON m.id=r.message_id JOIN users u ON u.id=m.mittente_id WHERE r.user_id=? ORDER BY m.id DESC LIMIT 200",
                     (u["id"],)).fetchall()
    return jsonify([{"id": r["id"], "titolo": r["titolo"], "anteprima": r["testo"][:120], "creato": r["creato"], "mittente": r["mittente"],
                     "letto": bool(r["letto"]), "n_allegati": r["n_all"]} for r in rows])


@app.get("/api/messaggi/<int:mid>")
@need("messaggi")
def api_msg_dettaglio(mid):
    u = current_user()
    c = db()
    m = c.execute("SELECT m.*, u.nome AS mittente FROM messages m JOIN users u ON u.id=m.mittente_id WHERE m.id=?", (mid,)).fetchone()
    if not m:
        abort(404)
    rc = c.execute("SELECT * FROM message_rcpt WHERE message_id=? AND user_id=?", (mid, u["id"])).fetchone()
    mio = m["mittente_id"] == u["id"]
    if not (rc or mio or u["ruolo"] == "admin"):
        abort(404)
    if rc and not rc["letto"]:
        c.execute("UPDATE message_rcpt SET letto=? WHERE message_id=? AND user_id=?", (ts(), mid, u["id"]))
        c.commit()
    out = {"id": m["id"], "titolo": m["titolo"], "testo": m["testo"], "creato": m["creato"], "mittente": m["mittente"], "mio": mio,
           "allegati": [_file_dict(f) for f in c.execute("SELECT * FROM files WHERE message_id=? ORDER BY id", (mid,))]}
    if mio or u["ruolo"] == "admin":
        out["destinatari"] = [{"nome": r["nome"], "ruolo": ROLE_LABEL[r["ruolo"]], "letto": r["letto"]} for r in c.execute(
            "SELECT u.nome,u.ruolo,r.letto FROM message_rcpt r JOIN users u ON u.id=r.user_id WHERE r.message_id=? ORDER BY u.nome", (mid,))]
    return jsonify(out)


@app.delete("/api/messaggi/<int:mid>")
@need("messaggi")
def api_msg_elimina(mid):
    u = current_user()
    c = db()
    m = c.execute("SELECT * FROM messages WHERE id=?", (mid,)).fetchone()
    if not m:
        abort(404)
    if m["mittente_id"] != u["id"] and u["ruolo"] != "admin":
        abort(403)
    for f in c.execute("SELECT * FROM files WHERE message_id=?", (mid,)).fetchall():
        _elimina_file_fisico(c, f)
    c.execute("DELETE FROM message_rcpt WHERE message_id=?", (mid,))
    c.execute("DELETE FROM messages WHERE id=?", (mid,))
    c.commit()
    return jsonify({"ok": True})


# ----------------------------------------------------------------- notifiche
@app.get("/api/notifications")
@need("notifiche")
def api_notifs():
    u = current_user()
    check_sla(db())
    rows = db().execute("SELECT id,messaggio,opp_id,tipo,quando,letta FROM notifications WHERE user_id=? "
                        "ORDER BY id DESC LIMIT 100", (u["id"],)).fetchall()
    return jsonify([dict(r) for r in rows])


@app.post("/api/notifications/read")
@need("notifiche")
def api_notifs_read():
    u = current_user()
    ids = body().get("ids")
    if ids:
        for i in ids:
            db().execute("UPDATE notifications SET letta=1 WHERE id=? AND user_id=?", (i, u["id"]))
    else:
        db().execute("UPDATE notifications SET letta=1 WHERE user_id=?", (u["id"],))
    db().commit()
    return jsonify({"ok": True})


# ----------------------------------------------------------------- utenti e permessi
USERNAME_RE = re.compile(r"^[a-z0-9._-]{3,40}$")


def user_row_dict(c, r):
    ov = {p["funzione"]: p["consentito"] for p in c.execute("SELECT funzione,consentito FROM user_permissions WHERE user_id=?", (r["id"],))}
    return {"id": r["id"], "username": r["username"], "nome": r["nome"], "ruolo": r["ruolo"],
            "ruolo_label": ROLE_LABEL[r["ruolo"]], "area_id": r["area_id"], "area": r["area"],
            "parent_id": r["parent_id"], "riferimento": r["riferimento"], "territori": r["territori"] or "",
            "attivo": bool(r["attivo"]), "override": ov, "effettivi": sorted(perms_of(c, r))}


@app.get("/api/users")
@need("utenti")
def api_users():
    u = current_user()
    c = db()
    q = ("SELECT u.*, ar.nome AS area, p.nome AS riferimento FROM users u LEFT JOIN areas ar ON ar.id=u.area_id "
         "LEFT JOIN users p ON p.id=u.parent_id")
    rows = c.execute(q + ("" if u["ruolo"] == "admin" else " WHERE u.area_id=? AND u.ruolo<>'admin'"),
                     () if u["ruolo"] == "admin" else (u["area_id"],)).fetchall()
    rows = sorted(rows, key=lambda r: (ROLE_RANK[r["ruolo"]], r["nome"]))
    zone = geo_summary(c)
    out = []
    for r in rows:
        d = user_row_dict(c, r)
        d["zone"] = zone.get(r["id"], {"province": [], "comuni": 0})
        out.append(d)
    return jsonify(out)


def validate_user(actor, d, existing=None):
    c = db()
    ruolo = d.get("ruolo", existing["ruolo"] if existing else None)
    if ruolo not in ROLES:
        return None, "Ruolo non valido"
    if ruolo == "admin" and actor["ruolo"] != "admin":
        return None, "Solo un amministratore può creare amministratori"
    nome = (d.get("nome", existing["nome"] if existing else "") or "").strip()
    if not nome:
        return None, "Il nome è obbligatorio"
    if ruolo == "admin":
        area_id, parent_id = None, None
    else:
        area_id = actor["area_id"] if actor["ruolo"] != "admin" else d.get("area_id", existing["area_id"] if existing else None)
        if not area_id or not c.execute("SELECT 1 FROM areas WHERE id=?", (area_id,)).fetchone():
            return None, "Seleziona un'area"
        parent_id = d.get("parent_id", existing["parent_id"] if existing else None) or None
        if ruolo == "direttore":
            parent_id = None
        elif ruolo in ("area_manager", "ufficio_disdette") and not parent_id:
            parent_id = None
        elif parent_id:
            p = c.execute("SELECT * FROM users WHERE id=? AND attivo=1", (parent_id,)).fetchone()
            if not p or p["area_id"] != area_id or ROLE_RANK[p["ruolo"]] >= ROLE_RANK[ruolo] or p["ruolo"] == "admin":
                return None, "Il referente deve essere un superiore della stessa area"
            if existing and parent_id in subtree_ids(existing["id"]):
                return None, "Gerarchia circolare non consentita"
        if ruolo in ("recuperatore", "ufficio_disdette", "area_manager") and parent_id:
            p = c.execute("SELECT ruolo FROM users WHERE id=?", (parent_id,)).fetchone()
            if p["ruolo"] != "direttore":
                return None, "Recuperatori, Ufficio Disdette e Area manager rispondono solo al direttore di area"
    return {"ruolo": ruolo, "nome": nome[:100], "area_id": area_id, "parent_id": parent_id,
            "territori": (d.get("territori", existing["territori"] if existing else "") or "")[:500]}, None


@app.post("/api/users")
@need("utenti")
def api_user_create():
    actor = current_user()
    d = body()
    uname = (d.get("username") or "").strip().lower()
    if not USERNAME_RE.match(uname):
        return err("Username: 3-40 caratteri tra lettere minuscole, numeri, . _ -")
    pw = d.get("password") or ""
    if len(pw) < 8:
        return err("La password deve avere almeno 8 caratteri")
    if db().execute("SELECT 1 FROM users WHERE username=?", (uname,)).fetchone():
        return err("Username già in uso")
    clean, e = validate_user(actor, d)
    if e:
        return err(e)
    cur = db().execute("INSERT INTO users(username,password_hash,nome,ruolo,area_id,parent_id,territori,must_change,creato) "
                       "VALUES(?,?,?,?,?,?,?,1,?)", (uname, generate_password_hash(pw), clean["nome"], clean["ruolo"],
                                                    clean["area_id"], clean["parent_id"], clean["territori"], ts()))
    db().commit()
    return jsonify({"id": cur.lastrowid})


@app.put("/api/users/<int:uid>")
@need("utenti")
def api_user_update(uid):
    actor = current_user()
    c = db()
    t = c.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    if not t or (actor["ruolo"] != "admin" and (t["area_id"] != actor["area_id"] or t["ruolo"] == "admin")):
        abort(404)
    d = body()
    clean, e = validate_user(actor, d, t)
    if e:
        return err(e)
    attivo = 1 if d.get("attivo", bool(t["attivo"])) else 0
    if uid == actor["id"] and (not attivo or clean["ruolo"] != t["ruolo"]):
        return err("Non puoi disattivarti o cambiare il tuo ruolo")
    if t["ruolo"] == "admin" and (not attivo or clean["ruolo"] != "admin"):
        if c.execute("SELECT COUNT(*) FROM users WHERE ruolo='admin' AND attivo=1 AND id<>?", (uid,)).fetchone()[0] == 0:
            return err("Deve restare almeno un amministratore attivo")
    c.execute("UPDATE users SET nome=?, ruolo=?, area_id=?, parent_id=?, territori=?, attivo=? WHERE id=?",
              (clean["nome"], clean["ruolo"], clean["area_id"], clean["parent_id"], clean["territori"], attivo, uid))
    if isinstance(d.get("permessi"), dict) and clean["ruolo"] != "admin":
        for f, v in d["permessi"].items():
            if f not in FUNZ_KEYS:
                continue
            if v is None:
                c.execute("DELETE FROM user_permissions WHERE user_id=? AND funzione=?", (uid, f))
            else:
                c.execute("INSERT OR REPLACE INTO user_permissions(user_id,funzione,consentito) VALUES(?,?,?)", (uid, f, 1 if v else 0))
    c.commit()
    return jsonify({"ok": True})


@app.post("/api/users/<int:uid>/password")
@need("utenti")
def api_user_password(uid):
    actor = current_user()
    t = db().execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    if not t or (actor["ruolo"] != "admin" and (t["area_id"] != actor["area_id"] or t["ruolo"] == "admin")):
        abort(404)
    pw = body().get("password") or ""
    if len(pw) < 8:
        return err("La password deve avere almeno 8 caratteri")
    db().execute("UPDATE users SET password_hash=?, must_change=1 WHERE id=?", (generate_password_hash(pw), uid))
    db().commit()
    return jsonify({"ok": True})


# ----------------------------------------------------------------- impostazioni / export
@app.get("/api/settings")
@need()
def api_settings_get():
    return jsonify({k: get_setting(db(), k) for k in SETTINGS_DEFAULT})


@app.put("/api/settings")
@need("impostazioni")
def api_settings_put():
    u = current_user()
    d = body()
    c = db()
    if any(k in d for k in ["sla_ore"] + SLA_KEYS) and u["ruolo"] != "admin":
        return err("Le regole SLA può modificarle solo l'amministratore", 403)
    try:
        mx = int(d.get("max_appuntamenti_giorno", get_setting(c, "max_appuntamenti_giorno")))
        sla = int(d.get("sla_ore", get_setting(c, "sla_ore")))
        spec = {}
        for k in SLA_KEYS:
            if k in d:
                v = str(d[k]).strip()
                spec[k] = "" if v == "" else str(int(v))
    except (TypeError, ValueError):
        return err("Valori numerici non validi")
    reg = d.get("regola_default", get_setting(c, "regola_default"))
    if not (1 <= sla <= 720 and 1 <= mx <= 10 and reg in REGOLE and all(v == "" or 1 <= int(v) <= 720 for v in spec.values())):
        return err("Valori fuori dai limiti consentiti (SLA da 1 a 720 ore)")
    for k, v in [("sla_ore", sla), ("max_appuntamenti_giorno", mx), ("regola_default", reg)] + list(spec.items()):
        c.execute("INSERT OR REPLACE INTO settings(chiave,valore) VALUES(?,?)", (k, str(v)))
    c.commit()
    return jsonify({"ok": True})


@app.get("/api/backup")
@need("impostazioni")
def api_backup():
    """Copia completa dei dati (solo amministratore): file .db scaricabile."""
    u = current_user()
    if u["ruolo"] != "admin":
        abort(403)
    import tempfile
    fd, tmp = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    src = connect()
    dst = sqlite3.connect(tmp)
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()
    with open(tmp, "rb") as f:
        data = f.read()
    try:
        os.remove(tmp)
    except OSError:
        pass
    nome = "sicuritalia-backup-%s.db" % now().strftime("%Y%m%d-%H%M")
    resp = app.response_class(data, mimetype="application/octet-stream")
    resp.headers["Content-Disposition"] = 'attachment; filename="%s"' % nome
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.get("/api/export.xlsx")
@need("export")
def api_export():
    u = current_user()
    where, ps = filters_from_request(u)
    rows = db().execute(BASE_SELECT + " WHERE " + where + " ORDER BY o.id", ps).fetchall()
    wb = Workbook()
    ws = wb.active
    ws.title = "Rendicontazione"
    ws.append(["ID", "Campagna", "Categoria", "Tipo", "Cliente", "Telefono", "Località", "Area", "Agente", "Stato",
               "Note esito", "Assegnata", "Scadenza SLA", "Prima gestione", "Appuntamento", "Riprogrammazioni",
               "Esito disdetta", "Motivo non recuperata", "Canone originario", "Valore recuperato", "Chiusa il", "Chiamate"])
    for r in rows:
        ws.append([r["id"], r["campagna"], r["categoria"], r["sottotipo"], r["cliente"], r["telefono"], r["localita"],
                   r["area"], r["agente"], r["stato"], r["esito_nota"], r["assegnata"], r["scadenza"], r["gestita"],
                   r["appuntamento"], r["n_riprogrammazioni"], r["disdetta_esito"] or "",
                   MOTIVI_NON_REC.get(r["disdetta_motivo"], ""), r["canone"], r["valore"], r["chiusa"], r["n_chiamate"] or 0])
    for row in ws.iter_rows(min_row=2):  # i testi importati che iniziano con "=" restano testo (no formule in Excel)
        for cell in row:
            if isinstance(cell.value, str) and cell.value[:1] == "=":
                cell.data_type = "s"
    ws.freeze_panes = "A2"
    for i, w in enumerate((8, 28, 14, 14, 28, 16, 18, 24, 22, 16, 40, 19, 19, 19, 19, 10, 16, 36, 14, 14, 19, 9), start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="rendicontazione_%s.xlsx" % now().strftime("%Y%m%d"),
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")



# ----------------------------------------------------------------- statistica disdette (trimestri solari)
def _stat_disdette(u, anno, agent_id=None, area_id=None):
    """Le disdette sono attribuite al trimestre solare in cui sono state assegnate. Il trimestre ha un mese solare in più
    (quello successivo alla fine) per completare la lavorazione: la statistica è definitiva dal giorno dopo."""
    where, ps = scope_where(u, agent_id, area_id, "disdetta")
    base = ("FROM opportunities o LEFT JOIN campaigns c ON c.id=o.campaign_id WHERE " + where +
            " AND substr(COALESCE(o.assegnata,o.creata),1,4)=?")
    rows = db().execute("SELECT o.disdetta_esito es, o.disdetta_motivo mot, o.valore val, o.canone can, COALESCE(o.assegnata,o.creata) ar, o.stato st " + base,
                        ps + [str(anno)]).fetchall()

    def pct(n, t):
        return round(100.0 * n / t, 1) if t else None

    def blocco(rs):
        rs_es = [r for r in rs if r["es"]]
        rec = [r for r in rs_es if r["es"] == "recuperata"]
        nl = [r for r in rs_es if r["es"] == "non_recuperata" and r["mot"] == "non_lavorabile"]
        non = [r for r in rs_es if r["es"] == "non_recuperata" and r["mot"] != "non_lavorabile"]
        lav = len(rec) + len(non)  # le «non lavorabili» sono escluse da tutte le percentuali
        motivi = []
        for k, lab in list(MOTIVI_NON_REC.items()) + [("", "Motivo non indicato (chiusure precedenti)")]:
            if k == "non_lavorabile":
                motivi.append({"id": k, "label": lab + " (escluse dal calcolo)", "n": len(nl), "pct_non_rec": None, "pct_totale": None, "escluso": True})
                continue
            n = sum(1 for r in non if (r["mot"] or "") == k)
            if n or k:
                motivi.append({"id": k or "nd", "label": lab, "n": n, "pct_non_rec": pct(n, len(non)), "pct_totale": pct(n, lav)})
        return {"ricevute": len(rs), "aperte": len(rs) - len(rs_es), "esitate": len(rs_es), "non_lavorabili": len(nl), "lavorate": lav,
                "recuperate": len(rec), "pct_recuperate": pct(len(rec), lav),
                "valore": round(sum(r["val"] or 0 for r in rec), 2),
                "canone_esitate": round(sum(r["can"] or 0 for r in rec + non), 2),
                "canone_recuperate": round(sum(r["can"] or 0 for r in rec), 2),
                "pct_valore": pct(sum(r["val"] or 0 for r in rec), sum(r["can"] or 0 for r in rec)),
                "non_recuperate": len(non), "pct_non_recuperate": pct(len(non), lav), "motivi": motivi}

    out = []
    oggi = now().date()
    for q in (1, 2, 3, 4):
        b = blocco([r for r in rows if (int(r["ar"][5:7]) + 2) // 3 == q])
        fine_mese = 3 * q + 1  # mese successivo alla fine del trimestre
        ultimo = (date(anno + (1 if fine_mese == 13 else 0), (fine_mese % 12) + 1, 1) - timedelta(days=1))
        b.update({"trimestre": q, "label": "%d° trimestre" % q, "mesi": ("gen–mar", "apr–giu", "lug–set", "ott–dic")[q - 1],
                  "definitiva": oggi > ultimo, "definitiva_dal": (ultimo + timedelta(days=1)).isoformat(),
                  "lavorazione_fino": ultimo.isoformat()})
        out.append(b)
    tot = blocco(rows)
    tot.update({"trimestre": 0, "label": "Totale %d" % anno, "mesi": "anno solare", "definitiva": all(t["definitiva"] for t in out)})
    anni = sorted({r[0] for r in db().execute("SELECT substr(COALESCE(o.assegnata,o.creata),1,4) " + base.replace(" AND substr(COALESCE(o.assegnata,o.creata),1,4)=?", ""), ps)} | {str(now().year)}, reverse=True)
    return {"anno": anno, "anni": anni, "trimestri": out, "totale": tot}


def _solo_disdette_ruoli():
    if current_user()["ruolo"] not in ("admin", "direttore", "area_manager", "recuperatore", "ufficio_disdette"):
        abort(403)


def _anno_arg():
    try:
        a = int(request.args.get("anno") or now().year)
    except ValueError:
        a = now().year
    return a if 2000 <= a <= 2100 else now().year


@app.get("/api/statistiche/disdette")
@need("dashboard")
def api_stat_disdette():
    u = current_user()
    _solo_disdette_ruoli()
    ag = request.args.get("agent_id", type=int)
    ar = request.args.get("area_id", type=int) if u["ruolo"] == "admin" else None
    return jsonify(_stat_disdette(u, _anno_arg(), ag, ar))


@app.get("/api/statistiche/disdette/dettaglio")
@need("dashboard")
def api_stat_disdette_dettaglio():
    """Elenco delle disdette che compongono un numero della statistica (trimestre e voce cliccati)."""
    u = current_user()
    _solo_disdette_ruoli()
    anno = _anno_arg()
    q = request.args.get("trimestre", 0, type=int)
    filtro = request.args.get("filtro", "esitate")
    motivo = request.args.get("motivo", "")
    ag = request.args.get("agent_id", type=int)
    ar = request.args.get("area_id", type=int) if u["ruolo"] == "admin" else None
    where, ps = scope_where(u, ag, ar, "disdetta")
    rows = db().execute(BASE_SELECT + " WHERE " + where + " AND substr(COALESCE(o.assegnata,o.creata),1,4)=? ORDER BY o.chiusa, o.id", ps + [str(anno)]).fetchall()
    out = []
    for r in rows:
        ar_ = r["assegnata"] or r["creata"] or ""
        if q and (int(ar_[5:7]) + 2) // 3 != q:
            continue
        es, mot = r["disdetta_esito"], r["disdetta_motivo"]
        if filtro == "ricevute":
            ok = True
        elif filtro == "aperte":
            ok = not es
        elif filtro == "esitate":
            ok = bool(es)
        elif filtro == "recuperate":
            ok = es == "recuperata"
        elif filtro == "non_lavorabili":
            ok = es == "non_recuperata" and mot == "non_lavorabile"
        elif filtro == "non_recuperate":
            ok = es == "non_recuperata" and mot != "non_lavorabile" and (not motivo or (mot or "nd") == motivo)
        else:
            return err("Filtro non valido")
        if ok:
            out.append({"id": r["id"], "cliente": r["cliente"], "localita": r["localita"], "recuperatore": r["agente"], "assegnata": ar_,
                        "chiusa": r["chiusa"], "stato": r["stato"], "esito": es, "motivo": MOTIVI_NON_REC.get(mot, "" if not es or es == "recuperata" else "Motivo non indicato"),
                        "canone": r["canone"], "valore": r["valore"], "nota": r["esito_nota"], "chiamate": r["n_chiamate"] or 0})
    return jsonify({"anno": anno, "trimestre": q, "filtro": filtro, "totale": len(out), "items": out[:500],
                    "somma_valore": round(sum(x["valore"] or 0 for x in out), 2), "somma_canone": round(sum(x["canone"] or 0 for x in out), 2)})


@app.get("/api/statistiche/disdette.xlsx")
@need("export")
def api_stat_disdette_xlsx():
    u = current_user()
    _solo_disdette_ruoli()
    ag = request.args.get("agent_id", type=int)
    ar = request.args.get("area_id", type=int) if u["ruolo"] == "admin" else None
    st = _stat_disdette(u, _anno_arg(), ag, ar)
    wb = Workbook()
    ws = wb.active
    ws.title = "Disdette %d" % st["anno"]
    ws.append(["Periodo", "Mesi", "Stato statistica", "Disdette ricevute", "Ancora in lavorazione", "Disdette esitate", "Di cui non lavorabili (escluse dalle %)", "Recuperate", "% recuperate", "Valore recuperato (€)", "Canone originario recuperate (€)", "Valore su canone originario (%)", "Canone originario esitate (€)", "Non recuperate", "% non recuperate"]
              + [v for v in MOTIVI_NON_REC.values()] + ["Motivo non indicato"])
    for t in st["trimestri"] + [st["totale"]]:
        mp = {m["id"]: m["n"] for m in t["motivi"]}
        ws.append([t["label"], t["mesi"], "definitiva" if t["definitiva"] else "provvisoria", t["ricevute"], t["aperte"], t["esitate"], t["non_lavorabili"], t["recuperate"], t["pct_recuperate"], t["valore"], t["canone_recuperate"], t["pct_valore"], t["canone_esitate"], t["non_recuperate"],
                   t["pct_non_recuperate"]] + [mp.get(k, 0) for k in MOTIVI_NON_REC] + [mp.get("nd", 0)])
    ws.freeze_panes = "A2"
    for i, w in enumerate((16, 14, 16, 16, 20, 16, 22, 12, 13, 20, 20, 20, 20, 15, 16, 34, 34, 34, 18, 22), start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="disdette_trimestri_%d.xlsx" % st["anno"],
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ----------------------------------------------------------------- geografia
def ids_arg(name):
    return [int(x) for x in (request.args.get(name) or "").split(",") if x.strip().isdigit()]


def comune_index(conn):
    idx = {}
    for r in conn.execute("SELECT id,nome,provincia_id FROM geo_comuni"):
        idx.setdefault(geo_seed.key(r["nome"]), []).append((r["id"], r["provincia_id"]))
    return idx


def user_prov_ids(conn, uid):
    ids = {r[0] for r in conn.execute("SELECT geo_id FROM user_geo WHERE user_id=? AND tipo='provincia'", (uid,))}
    ids |= {r[0] for r in conn.execute(
        "SELECT DISTINCT c.provincia_id FROM user_geo g JOIN geo_comuni c ON c.id=g.geo_id WHERE g.user_id=? AND g.tipo='comune'", (uid,))}
    return ids


def pick_comune(idx, localita, hint_prov):
    txt = re.sub(r"\(.*?\)", " ", localita or "").strip()
    cands = [txt]
    parts = txt.split()
    if len(parts) > 1 and len(parts[-1]) == 2:
        cands.append(" ".join(parts[:-1]))
    for t in cands:
        found = idx.get(geo_seed.key(t))
        if found:
            for cid, pid in found:
                if pid in hint_prov:
                    return cid
            return found[0][0]
    return None


_CAP = re.compile(r"\b\d{5}\b")


def _num(v):
    try:
        x = float(v)
        return x if x == x else None
    except (TypeError, ValueError):
        return None


def comune_vicino(conn, lat, lon):
    """Comune dell'archivio più vicino alle coordinate (None se lontano più di ~25 km)."""
    try:
        lat, lon = float(lat), float(lon)
    except (TypeError, ValueError):
        return None
    for d in (0.1, 0.25):
        rows = conn.execute("SELECT id,lat,lon FROM geo_comuni WHERE lat BETWEEN ? AND ? AND lon BETWEEN ? AND ?",
                            (lat - d, lat + d, lon - d, lon + d)).fetchall()
        if rows:
            best = min(rows, key=lambda r: (r["lat"] - lat) ** 2 + ((r["lon"] - lon) * 0.74) ** 2)
            if ((best["lat"] - lat) * 111) ** 2 + ((best["lon"] - lon) * 82) ** 2 <= 25 ** 2:
                return best["id"]
    return None


def trova_comune(conn, idx, localita, indirizzo, hint_prov, lat=None, lon=None):
    """Comune di un'opportunità: dal campo località, poi dall'indirizzo (anche «Via X 5, 31100 Treviso TV»), infine dalle coordinate."""
    cid = pick_comune(idx, localita, hint_prov) if localita else None
    if cid:
        return cid
    testo = " ".join(x for x in (localita, indirizzo) if x)
    if testo.strip():
        pezzi = [x.strip() for x in re.split(r"[,;\-–/]", testo) if x.strip()]
        for pz in reversed(pezzi):
            pz2 = _CAP.sub(" ", pz).strip()
            cid = pick_comune(idx, pz2, hint_prov)
            if cid:
                return cid
        toks = re.findall(r"[A-Za-zÀ-ÿ'’]+", _CAP.sub(" ", testo))
        for n in (3, 2, 1):
            for i in range(len(toks) - n, -1, -1):
                w = " ".join(toks[i:i + n])
                if n == 1 and (len(w) < 4 or w.lower() in IT_STOP):
                    continue
                found = idx.get(geo_seed.key(w))
                if found:
                    for c_, pid in found:
                        if pid in hint_prov:
                            return c_
                    if n > 1 or w[:1].isupper():
                        return found[0][0]
    if lat is not None and lon is not None:
        return comune_vicino(conn, lat, lon)
    return None


def resolve_missing(conn, limit=20000):
    """Collega le opportunità ai comuni dell'archivio geografico (da località, indirizzo o coordinate)."""
    rows = conn.execute("SELECT id,localita,indirizzo,agente_id,lat,lon FROM opportunities WHERE comune_id IS NULL "
                        "AND (localita<>'' OR indirizzo<>'' OR lat IS NOT NULL) LIMIT ?", (limit,)).fetchall()
    if not rows:
        return 0
    idx, hints, n = comune_index(conn), {}, 0
    for r in rows:
        a = r["agente_id"]
        if a not in hints:
            hints[a] = user_prov_ids(conn, a) if a else set()
        cid = trova_comune(conn, idx, r["localita"] or "", r["indirizzo"] or "", hints[a], r["lat"], r["lon"])
        if cid:
            conn.execute("UPDATE opportunities SET comune_id=? WHERE id=?", (cid, r["id"]))
            n += 1
    if n:
        conn.commit()
    return n


def geo_names_for(conn, uid):
    """Nomi (normalizzati) dei comuni di lavoro di un utente; una provincia senza comuni scelti vale per intero."""
    sel = conn.execute("SELECT tipo,geo_id FROM user_geo WHERE user_id=?", (uid,)).fetchall()
    com = [r["geo_id"] for r in sel if r["tipo"] == "comune"]
    prov = [r["geo_id"] for r in sel if r["tipo"] == "provincia"]
    names = set()
    covered = set()
    if com:
        for r in conn.execute("SELECT nome,provincia_id FROM geo_comuni WHERE id IN (%s)" % ",".join("?" * len(com)), com):
            names.add(norm(r["nome"]))
            covered.add(r["provincia_id"])
    full = [p for p in prov if p not in covered]
    if full:
        for r in conn.execute("SELECT nome FROM geo_comuni WHERE provincia_id IN (%s)" % ",".join("?" * len(full)), full):
            names.add(norm(r["nome"]))
    return names


def geo_summary(conn):
    out = {}
    for r in conn.execute(
            "SELECT g.user_id uid, p.sigla, p.nome FROM user_geo g JOIN geo_province p ON p.id=g.geo_id WHERE g.tipo='provincia' "
            "UNION SELECT g.user_id, p.sigla, p.nome FROM user_geo g JOIN geo_comuni c ON c.id=g.geo_id "
            "JOIN geo_province p ON p.id=c.provincia_id WHERE g.tipo='comune'"):
        out.setdefault(r["uid"], {"province": [], "comuni": 0})["province"].append(r["sigla"] or r["nome"])
    for r in conn.execute("SELECT user_id uid, COUNT(*) n FROM user_geo WHERE tipo='comune' GROUP BY user_id"):
        out.setdefault(r["uid"], {"province": [], "comuni": 0})["comuni"] = r["n"]
    for v in out.values():
        v["province"] = sorted(set(v["province"]))
    return out


def set_geo_stato(conn, ok, msg):
    conn.execute("INSERT OR REPLACE INTO settings(chiave,valore) VALUES('geo_stato',?)",
                 ("%s|%s|%s" % ("ok" if ok else "errore", ts(), msg),))
    conn.commit()


def geo_stato(conn):
    n = conn.execute("SELECT COUNT(*) FROM geo_comuni").fetchone()[0]
    r = conn.execute("SELECT valore FROM settings WHERE chiave='geo_stato'").fetchone()
    parts = (r["valore"].split("|", 2) if r else ["", "", ""])
    return {"comuni": n, "province": conn.execute("SELECT COUNT(*) FROM geo_province").fetchone()[0],
            "regioni": conn.execute("SELECT COUNT(*) FROM geo_regioni").fetchone()[0],
            "completo": n >= 3000, "esito": parts[0], "quando": parts[1], "messaggio": parts[2]}


@app.get("/api/geo/regioni")
@need()
def api_geo_regioni():
    return jsonify([dict(r) for r in db().execute("SELECT id,nome FROM geo_regioni ORDER BY nome")])


@app.get("/api/geo/province")
@need()
def api_geo_province():
    ids = ids_arg("regioni")
    if not ids:
        return jsonify([])
    rows = db().execute("SELECT id,nome,sigla,regione_id FROM geo_province WHERE regione_id IN (%s) ORDER BY nome" % ",".join("?" * len(ids)), ids)
    return jsonify([dict(r) for r in rows])


@app.get("/api/geo/comuni")
@need()
def api_geo_comuni():
    ids = ids_arg("province")
    if not ids:
        return jsonify([])
    rows = db().execute("SELECT id,nome,provincia_id FROM geo_comuni WHERE provincia_id IN (%s) ORDER BY nome" % ",".join("?" * len(ids)), ids)
    return jsonify([dict(r) for r in rows])


@app.get("/api/geo/stato")
@need()
def api_geo_stato():
    return jsonify(geo_stato(db()))


def _geo_import(data, filename):
    rows = geo_load.parse_any(filename, data)
    if not rows:
        raise ValueError("Il file non contiene comuni riconoscibili")
    res = geo_load.import_rows(db(), rows)
    set_geo_stato(db(), True, "%d comuni letti dal file (%d nuovi)" % (len(rows), res["nuovi"]))
    return res


@app.post("/api/geo/aggiorna")
@need("aree")
def api_geo_aggiorna():
    try:
        res = _geo_import(geo_load.download(), "IT.zip")
    except ValueError as e:
        return err(str(e))
    except Exception as e:
        set_geo_stato(db(), False, "Download non riuscito: %s" % str(e)[:150])
        return err("Download non riuscito (il computer o la rete potrebbero bloccare internet). "
                   "Scarica l'elenco a mano e caricalo con «Carica un file». Dettaglio: %s" % str(e)[:120], 502)
    resolve_missing(db())
    return jsonify(dict(res, **geo_stato(db())))


@app.post("/api/geo/carica")
@need("aree")
def api_geo_carica():
    f = request.files.get("file")
    if not f:
        return err("Seleziona un file")
    try:
        res = _geo_import(f.read(), f.filename)
    except ValueError as e:
        return err(str(e))
    except Exception:
        return err("File non leggibile")
    resolve_missing(db())
    return jsonify(dict(res, **geo_stato(db())))


def _geo_autoload():
    """Al primo avvio prova a scaricare l'elenco completo dei comuni (se internet è disponibile)."""
    if os.environ.get("SOH_GEO_AUTO", "1") == "0":
        return
    c = None
    try:
        c = connect()
        if c.execute("SELECT COUNT(*) FROM geo_comuni").fetchone()[0] >= 3000:
            return
        rows = geo_load.parse_any("IT.zip", geo_load.download())
        res = geo_load.import_rows(c, rows)
        set_geo_stato(c, True, "Scaricati %d comuni (%d nuovi)" % (len(rows), res["nuovi"]))
        resolve_missing(c)
    except Exception as e:
        try:
            if c:
                set_geo_stato(c, False, "Download automatico non riuscito: %s" % str(e)[:150])
        except Exception:
            pass
    finally:
        if c:
            c.close()


def target_user(actor, uid):
    t = db().execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    if not t or (actor["ruolo"] != "admin" and (t["area_id"] != actor["area_id"] or t["ruolo"] == "admin")):
        abort(404)
    return t


@app.get("/api/users/<int:uid>/geo")
@need("utenti")
def api_user_geo_get(uid):
    target_user(current_user(), uid)
    out = {"regioni": [], "province": [], "comuni": []}
    plurale = {"regione": "regioni", "provincia": "province", "comune": "comuni"}
    for r in db().execute("SELECT tipo,geo_id FROM user_geo WHERE user_id=?", (uid,)):
        if r["tipo"] in plurale:
            out[plurale[r["tipo"]]].append(r["geo_id"])
    return jsonify(out)


@app.put("/api/users/<int:uid>/geo")
@need("utenti")
def api_user_geo_put(uid):
    actor = current_user()
    target_user(actor, uid)
    d = body()
    c = db()

    def ints(k):
        return {int(x) for x in (d.get(k) or []) if str(x).isdigit()}
    reg = ints("regioni") & {r[0] for r in c.execute("SELECT id FROM geo_regioni")}
    prov = set()
    if reg:
        prov = ints("province") & {r[0] for r in c.execute(
            "SELECT id FROM geo_province WHERE regione_id IN (%s)" % ",".join("?" * len(reg)), list(reg))}
    com = set()
    if prov:
        com = ints("comuni") & {r[0] for r in c.execute(
            "SELECT id FROM geo_comuni WHERE provincia_id IN (%s)" % ",".join("?" * len(prov)), list(prov))}
    c.execute("DELETE FROM user_geo WHERE user_id=?", (uid,))
    for tipo, vals in (("regione", reg), ("provincia", prov), ("comune", com)):
        c.executemany("INSERT INTO user_geo(user_id,tipo,geo_id) VALUES(?,?,?)", [(uid, tipo, v) for v in vals])
    c.commit()
    return jsonify({"ok": True, "regioni": len(reg), "province": len(prov), "comuni": len(com)})


# ----------------------------------------------------------------- aree
@app.get("/api/areas")
@need("aree")
def api_areas():
    rows = db().execute(
        "SELECT a.id,a.nome,(SELECT COUNT(*) FROM users u WHERE u.area_id=a.id) utenti,"
        "(SELECT COUNT(*) FROM opportunities o WHERE o.area_id=a.id) opportunita,"
        "(SELECT COUNT(*) FROM campaigns c WHERE c.area_id=a.id) campagne FROM areas a ORDER BY a.id").fetchall()
    return jsonify([dict(r) for r in rows])


def _admin_only():
    if current_user()["ruolo"] != "admin":
        abort(403)


@app.post("/api/areas")
@need("aree")
def api_area_create():
    _admin_only()
    nome = (body().get("nome") or "").strip()[:100]
    if len(nome) < 2:
        return err("Indica il nome dell'area")
    if db().execute("SELECT 1 FROM areas WHERE lower(nome)=lower(?)", (nome,)).fetchone():
        return err("Esiste già un'area con questo nome")
    cur = db().execute("INSERT INTO areas(nome) VALUES(?)", (nome,))
    db().commit()
    return jsonify({"id": cur.lastrowid})


@app.put("/api/areas/<int:aid>")
@need("aree")
def api_area_rename(aid):
    _admin_only()
    nome = (body().get("nome") or "").strip()[:100]
    if len(nome) < 2:
        return err("Indica il nome dell'area")
    if db().execute("SELECT 1 FROM areas WHERE lower(nome)=lower(?) AND id<>?", (nome, aid)).fetchone():
        return err("Esiste già un'area con questo nome")
    db().execute("UPDATE areas SET nome=? WHERE id=?", (nome, aid))
    db().commit()
    return jsonify({"ok": True})


@app.delete("/api/areas/<int:aid>")
@need("aree")
def api_area_delete(aid):
    _admin_only()
    c = db()
    n = c.execute("SELECT (SELECT COUNT(*) FROM users WHERE area_id=?)+(SELECT COUNT(*) FROM opportunities WHERE area_id=?)"
                  "+(SELECT COUNT(*) FROM campaigns WHERE area_id=?)", (aid, aid, aid)).fetchone()[0]
    if n:
        return err("L'area ha utenti, campagne o opportunità: non può essere eliminata")
    c.execute("DELETE FROM areas WHERE id=?", (aid,))
    c.commit()
    return jsonify({"ok": True})


# ----------------------------------------------------------------- organigramma
@app.get("/api/organigramma")
@need("organigramma")
def api_organigramma():
    u = current_user()
    c = db()
    ids = [r["id"] for r in c.execute("SELECT id FROM users WHERE ruolo<>'admin'")] if u["ruolo"] == "admin" else visible_user_ids(u)
    if not ids:
        return jsonify([])
    rows = c.execute(
        "SELECT u.id,u.nome,u.username,u.ruolo,u.parent_id,u.area_id,u.attivo,ar.nome AS area FROM users u "
        "LEFT JOIN areas ar ON ar.id=u.area_id WHERE u.id IN (%s) AND u.ruolo<>'admin' ORDER BY u.nome" % ",".join("?" * len(ids)), ids).fetchall()
    st = {r["agente_id"]: r for r in c.execute(
        "SELECT agente_id, SUM(stato NOT IN ('contratto','ko')) aperte, SUM(stato='contratto') contratti, SUM(stato='ko') ko "
        "FROM opportunities WHERE agente_id IS NOT NULL GROUP BY agente_id")}
    zone = geo_summary(c)
    out = []
    for r in rows:
        s_ = st.get(r["id"])
        out.append({"id": r["id"], "nome": r["nome"], "username": r["username"], "ruolo": r["ruolo"], "ruolo_label": ROLE_LABEL[r["ruolo"]],
                    "parent_id": r["parent_id"], "area_id": r["area_id"], "area": r["area"], "attivo": bool(r["attivo"]),
                    "aperte": (s_["aperte"] or 0) if s_ else 0, "contratti": (s_["contratti"] or 0) if s_ else 0,
                    "ko": (s_["ko"] or 0) if s_ else 0, "zone": zone.get(r["id"], {"province": [], "comuni": 0})})
    return jsonify(out)


# ----------------------------------------------------------------- mappa
def zone_comuni(c, uids):
    """Comuni (con coordinate) in cui lavorano gli utenti. Una provincia scelta senza comuni specifici vale per tutti i suoi comuni."""
    if not uids:
        return []
    ph = ",".join("?" * len(uids))
    prov, com = {}, {}
    for r in c.execute("SELECT user_id,tipo,geo_id FROM user_geo WHERE tipo IN ('provincia','comune') AND user_id IN (%s)" % ph, uids):
        (prov if r["tipo"] == "provincia" else com).setdefault(r["user_id"], set()).add(r["geo_id"])
    com_ids = {i for v in com.values() for i in v}
    com_prov = {}
    if com_ids:
        for r in c.execute("SELECT id,provincia_id FROM geo_comuni WHERE id IN (%s)" % ",".join("?" * len(com_ids)), list(com_ids)):
            com_prov[r["id"]] = r["provincia_id"]
    prov_tutte = set()
    for uid, ps_ in prov.items():
        scelte = {com_prov.get(i) for i in com.get(uid, set())}
        prov_tutte |= {p_ for p_ in ps_ if p_ not in scelte}
    out = {}
    if com_ids:
        for r in c.execute("SELECT id,nome,lat,lon FROM geo_comuni WHERE lat IS NOT NULL AND id IN (%s)" % ",".join("?" * len(com_ids)), list(com_ids)):
            out[r["id"]] = {"nome": r["nome"], "lat": r["lat"], "lon": r["lon"]}
    if prov_tutte:
        for r in c.execute("SELECT id,nome,lat,lon FROM geo_comuni WHERE lat IS NOT NULL AND provincia_id IN (%s) LIMIT 5000" % ",".join("?" * len(prov_tutte)), list(prov_tutte)):
            out[r["id"]] = {"nome": r["nome"], "lat": r["lat"], "lon": r["lon"]}
    return list(out.values())[:5000]


@app.get("/api/mappa")
@need("dashboard")
def api_mappa():
    u = current_user()
    c = db()
    resolve_missing(c)
    where, ps, fl = dash_filters(u)
    area_id, agent_ids = fl["area_id"], fl["agenti"]
    rows = c.execute(
        "SELECT gc.id, gc.nome, gc.lat, gc.lon, gp.sigla, SUM(o.stato='contratto') contratti, SUM(o.stato='ko') ko, "
        "SUM(o.stato NOT IN ('contratto','ko')) aperte FROM opportunities o LEFT JOIN campaigns c ON c.id=o.campaign_id "
        "JOIN geo_comuni gc ON gc.id=o.comune_id JOIN geo_province gp ON gp.id=gc.provincia_id "
        "WHERE " + where + " AND gc.lat IS NOT NULL GROUP BY gc.id", ps).fetchall()
    tot = c.execute("SELECT COUNT(*) FROM opportunities o LEFT JOIN campaigns c ON c.id=o.campaign_id WHERE " + where, ps).fetchone()[0]
    loc = sum((r["contratti"] or 0) + (r["ko"] or 0) + (r["aperte"] or 0) for r in rows)
    if agent_ids:
        uids = agent_ids
    elif u["ruolo"] == "admin" and not area_id:
        uids = [r[0] for r in c.execute("SELECT id FROM users")]
    elif u["ruolo"] == "admin":
        uids = [r[0] for r in c.execute("SELECT id FROM users WHERE area_id=?", (area_id,))]
    else:
        uids = visible_user_ids(u)
    zone = zone_comuni(c, uids)
    esempi = []
    if tot - loc > 0:
        for r in c.execute("SELECT o.cliente,o.localita,o.indirizzo FROM opportunities o LEFT JOIN campaigns c ON c.id=o.campaign_id "
                           "WHERE " + where + " AND (o.comune_id IS NULL OR o.comune_id NOT IN (SELECT id FROM geo_comuni WHERE lat IS NOT NULL)) LIMIT 8", ps):
            esempi.append("%s (%s)" % (r["cliente"], (r["localita"] or r["indirizzo"] or "nessuna località")[:50]))
    return jsonify({"esempi_non_localizzate": esempi, "punti": [{"id": r["id"], "comune": r["nome"], "sigla": r["sigla"], "lat": r["lat"], "lon": r["lon"],
                               "contratti": r["contratti"] or 0, "ko": r["ko"] or 0, "aperte": r["aperte"] or 0} for r in rows],
                    "zone": zone, "non_localizzate": tot - loc, "totale": tot})


# ----------------------------------------------------------------- statici
@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.errorhandler(404)
def _404(_e):
    if request.path.startswith("/api/"):
        return err("Non trovato", 404)
    if request.path.startswith("/cal/"):
        return app.response_class("Calendario non trovato: l'indirizzo non è valido o è stato sostituito.", 404, mimetype="text/plain")
    return send_from_directory(app.static_folder, "index.html")


@app.errorhandler(400)
def _400(_e):
    return err("Richiesta non valida (controlla le date)", 400)


@app.errorhandler(403)
def _403(_e):
    return err("Non consentito", 403)


@app.errorhandler(413)
def _413(_e):
    return err("File troppo grande (max 10 MB)", 413)


@app.after_request
def _headers(resp):
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("X-Frame-Options", "DENY")
    if request.path.startswith("/api/") and resp.mimetype == "application/json" and not resp.direct_passthrough:
        try:
            raw = resp.get_data(as_text=True)
            if "u27e6" in raw or "\u27e6" in raw:
                u_ = current_user()
                resp.set_data(_redigi(raw, u_["ruolo"] if u_ else ""))
        except Exception:
            pass
    resp.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")  # OpenStreetMap richiede il Referer
    if request.path.startswith("/api/"):
        resp.headers["Cache-Control"] = "no-store"
    elif request.path in ("/sw.js", "/") or request.path.endswith(".webmanifest"):
        resp.headers["Cache-Control"] = "no-cache"
    return resp


@app.get("/sw.js")
def service_worker():
    r = send_from_directory(app.static_folder, "sw.js", mimetype="text/javascript")
    r.headers["Service-Worker-Allowed"] = "/"
    return r


@app.get("/manifest.webmanifest")
def manifest():
    return send_from_directory(app.static_folder, "manifest.webmanifest", mimetype="application/manifest+json")


# ----------------------------------------------------------------- carica da email
import email_parse


def _assignable(c, u, tipo):
    ruoli = ("agente", "recuperatore") if tipo == "disdetta" else ("agente",)
    ids = visible_user_ids(u)
    if not ids:
        return []
    rows = c.execute("SELECT id,nome,ruolo,area_id FROM users WHERE attivo=1 AND ruolo IN (%s) AND id IN (%s) ORDER BY nome" %
                     (",".join("?" * len(ruoli)), ",".join("?" * len(ids))), list(ruoli) + ids).fetchall()
    return [dict(r) for r in rows]


def _suggest_agent(c, cand, parsed):
    if not cand:
        return None
    if parsed.get("tipo") == "disdetta" and parsed.get("agente_gestore"):
        toks = set(norm(parsed["agente_gestore"]).replace(".", " ").split())
        best = None
        for a in cand:
            at = set(norm(a["nome"]).replace(".", " ").split())
            sc = len(toks & at)
            if sc >= min(2, len(at)) and sc > 0 and (best is None or sc > best[0]):
                best = (sc, a["id"])
        if best:
            return best[1]
    loc = norm(parsed.get("localita"))
    if loc:
        hits = [a for a in cand if loc in geo_names_for(c, a["id"])]
        if hits:
            def carico(a):
                return c.execute("SELECT COUNT(*) FROM opportunities WHERE agente_id=? AND stato IN ('da_gestire','in_lavorazione','da_riprogrammare')",
                                 (a["id"],)).fetchone()[0]
            return min(hits, key=carico)["id"]
    return None


def _duplicate(c, tipo, rif):
    if not rif:
        return None
    r = c.execute("SELECT o.id,o.cliente,o.stato,o.creata FROM opportunities o LEFT JOIN campaigns k ON k.id=o.campaign_id "
                  "WHERE o.rif_esterno=? AND (k.categoria='disdetta')=? ORDER BY o.id DESC LIMIT 1", (rif, tipo == "disdetta")).fetchone()
    return dict(r) if r else None


@app.post("/api/email/leggi")
@need("email")
def api_email_leggi():
    u = current_user()
    c = db()
    try:
        if request.files.get("file"):
            f = request.files["file"]
            text = email_parse.extract_text(f.filename, f.read())
        else:
            text = email_parse._clean(str(body().get("testo") or ""))[:email_parse.MAX_CHARS]
            if len(text) < 20:
                return err("Incolla il testo della mail oppure scegli un file")
    except ValueError as e:
        return err(str(e))
    p = email_parse.parse(text)
    tipo = p["tipo"] or "lead"
    cand = _assignable(c, u, tipo)
    cand_l = _assignable(c, u, "lead")
    cand_d = _assignable(c, u, "disdetta")
    dup = _duplicate(c, tipo, p.get("rif_esterno")) if p["tipo"] else None
    return jsonify({"parsed": p, "riconosciuta": bool(p["tipo"]), "duplicato": dup,
                    "assegnatari": {"lead": cand_l, "disdetta": cand_d},
                    "suggerito": {"lead": _suggest_agent(c, cand_l, dict(p, tipo="lead")),
                                  "disdetta": _suggest_agent(c, cand_d, dict(p, tipo="disdetta"))},
                    "assegna_a_me": u["ruolo"] in ("agente", "recuperatore")})


def _email_campaign(c, u, area_id, tipo):
    mese = ("gennaio febbraio marzo aprile maggio giugno luglio agosto settembre ottobre novembre dicembre".split())[now().month - 1]
    if tipo == "disdetta":
        nome, cat, sot = "Disdette da email - %s %d" % (mese, now().year), "disdetta", ""
    else:
        nome, cat, sot = "Lead da email - %s %d" % (mese, now().year), "prospect", "lead"
    r = c.execute("SELECT id FROM campaigns WHERE nome=? AND area_id IS ?", (nome, area_id)).fetchone()
    if r:
        return r["id"]
    cur = c.execute("INSERT INTO campaigns(nome,categoria,sottotipo,area_id,creato_da,creato,regola) VALUES(?,?,?,?,?,?,?)",
                    (nome, cat, sot, area_id, u["id"], ts(), "non_assegnare"))
    return cur.lastrowid


@app.post("/api/email/crea")
@need("email")
def api_email_crea():
    u = current_user()
    c = db()
    d = body()
    tipo = d.get("tipo")
    if tipo not in ("lead", "disdetta"):
        return err("Scegli se è un Lead o una Disdetta")
    cliente = (d.get("cliente") or "").strip()[:200]
    if not cliente:
        return err("Indica il nome del cliente")
    if u["ruolo"] in ("agente", "recuperatore"):
        t = u
    else:
        aid = d.get("agent_id")
        t = next((a for a in _assignable(c, u, tipo) if a["id"] == aid), None)
        if not t:
            return err("Scegli l'agente a cui assegnare l'opportunità")
    if t["area_id"] is None:
        return err("L'utente scelto non appartiene a nessuna area")
    rif = (d.get("rif_esterno") or "").strip()[:60]
    dup = _duplicate(c, tipo, rif)
    if dup and not d.get("forza"):
        return jsonify({"error": "Questa mail risulta già caricata (%s, stato %s)." % (dup["cliente"], dup["stato"]), "duplicato": dup}), 409
    ap = (d.get("appuntamento") or "").strip() if tipo == "lead" else ""
    ap_dt = None
    if ap:
        try:
            ap_dt = datetime.strptime(ap.replace("T", " ")[:16], "%Y-%m-%d %H:%M")
        except ValueError:
            ap_dt = None
    if ap_dt and ap_dt.date() >= now().date() and not d.get("forza_app"):
        info = verifica_agenda(c, t["id"], ap_dt.strftime("%Y-%m-%d"), ap_dt.strftime("%H:%M"))
        if info["problema"]:
            return jsonify({"error": info["avviso"], "conferma_app": info, "agente": t["nome"]}), 409
    area = t["area_id"]
    loc = (d.get("localita") or "").strip()[:120]
    prov = (d.get("provincia") or "").strip().upper()[:2]
    cid = _email_campaign(c, u, area, tipo)
    canone = _valore(d.get("canone")) if tipo == "disdetta" and str(d.get("canone") or "").strip() else None
    idx = comune_index(c)
    comune = pick_comune(idx, loc, user_prov_ids(c, t["id"]) | {r[0] for r in c.execute(
        "SELECT id FROM geo_province WHERE sigla=?", (prov,))}) if loc else None
    sla = sla_ore(c, "disdetta" if tipo == "disdetta" else "prospect", "lead" if tipo == "lead" else "")
    cur = c.execute(
        "INSERT INTO opportunities(campaign_id,area_id,cliente,telefono,localita,indirizzo,note,agente_id,stato,assegnata,scadenza,creata,comune_id,rif_esterno) "
        "VALUES(?,?,?,?,?,?,?,?,'da_gestire',?,?,?,?,?)",
        (cid, area, cliente, (d.get("telefono") or "").strip()[:40], (loc + (" (%s)" % prov if prov and "(" not in loc else ""))[:140],
         (d.get("indirizzo") or "").strip()[:300], (d.get("note") or "").strip()[:4000], t["id"], ts(),
         ts(now() + timedelta(hours=sla)), ts(), comune, rif))
    oid = cur.lastrowid
    if canone is not None:
        c.execute("UPDATE opportunities SET canone=? WHERE id=?", (canone, oid))
    log_event(c, oid, u["id"], "creata", "da email (%s)" % tipo)
    log_event(c, oid, u["id"], "assegnata", "a %s" % t["nome"])
    warn, fissato = "", ""
    if tipo == "lead" and ap:
        if not ap_dt:
            warn = "Data appuntamento non valida: ignorata."
        else:
            o = c.execute("SELECT * FROM opportunities WHERE id=?", (oid,)).fetchone()
            e = set_appointment(c, o, ap_dt.strftime("%Y-%m-%d"), ap_dt.strftime("%H:%M"), u, True)
            if e:
                warn = "Appuntamento del %s non fissato: %s. Resta da gestire." % (ap_dt.strftime("%d/%m/%Y %H:%M"), e)
                log_event(c, oid, u["id"], "nota", warn)
            else:
                fissato = ap_dt.strftime("%d/%m/%Y alle %H:%M")
    notify(c, cc_for_assignment(c, t["id"]), "Nuova %s da email assegnata a %s: %s" % ("disdetta" if tipo == "disdetta" else "lead", t["nome"], cliente), "assegnazione", oid)
    c.commit()
    return jsonify({"ok": True, "id": oid, "assegnato_a": t["nome"], "avviso": warn, "fissato": fissato})


# ----------------------------------------------------------------- proposte degli agenti (assistente + manuale)
IT_STOP = set("a ad al alla alle allo agli ai da dal dalla di del della delle dei in il la le lo un una uno per con su tra fra e ed o ma che chi presso zona vicino sig signor signora sig.ra ditta azienda cliente societa srl spa snc sas tel telefono cell cellulare numero vorrebbe vuole interessato interessata preventivo antifurto allarme impianto richiesta lead nuovo nuova contatto contattare chiamare richiamare ha hanno è sono molto".split())
STRADE = r"(?:via|viale|v\.le|piazza|p\.zza|corso|c\.so|strada|str\.|largo|vicolo|localit[aà]|loc\.|borgo|calle|campo|fondamenta|salizada)"


def parse_chat(conn, u, text):
    """Assistente: ricava i dati di un'opportunità da una frase libera (regole italiane, senza servizi esterni)."""
    t = " ".join((text or "").replace("\r", " ").split())
    out = {"cliente": "", "telefono": "", "localita": "", "provincia": "", "indirizzo": "", "note": (text or "").strip()[:2000]}
    ph = email_parse._phones(t)
    out["telefono"] = ph[0] if ph else ""
    # indirizzo
    m = re.search(r"(?i)\b(" + STRADE + r"\s+[A-Za-zÀ-ÿ'’.\- ]{2,40}?(?:\s*,?\s*(?:n\.?|civico)?\s*\d{1,4}[A-Za-z]?)?)(?=[,;.]|\s+(?:a|in|di|a|tel|cell|telefono)\b|$)", t)
    if m:
        out["indirizzo"] = m.group(1).strip(" ,.")
    # località: cerca i comuni noti (n-grammi, dal più lungo)
    idx = comune_index(conn)
    mine = user_prov_ids(conn, u["id"])
    toks = re.findall(r"[A-Za-zÀ-ÿ'’]+", t)
    best = None
    for n in (3, 2, 1):
        for i in range(len(toks) - n + 1):
            w = toks[i:i + n]
            if n == 1 and (len(w[0]) < 4 or w[0].lower() in IT_STOP):
                continue
            found = idx.get(geo_seed.key(" ".join(w)))
            if not found:
                continue
            prep = i > 0 and toks[i - 1].lower() in ("a", "ad", "in", "di", "presso", "zona", "vicino", "comune")
            cap = w[0][:1].isupper()
            inzona = any(pid in mine for _c, pid in found)
            score = (n, prep, inzona, cap)
            if (prep or inzona or cap) and (best is None or score > best[0]):
                best = (score, " ".join(w), found)
        if best and best[0][0] == n:
            break
    if best:
        _sc, nome, found = best
        pid = next((p for _c, p in found if p in mine), found[0][1])
        r = conn.execute("SELECT sigla FROM geo_province WHERE id=?", (pid,)).fetchone()
        out["localita"] = nome.title()
        out["provincia"] = (r["sigla"] if r else "") or ""
    # cliente
    skip = set(w.lower() for w in (out["localita"] or "").split())
    PART = ("da", "di", "del", "della", "de", "dei", "degli", "e", "&", "dal", "dalla")
    SUFF = ("srl", "spa", "snc", "sas", "s.r.l.", "s.n.c.", "s.p.a.", "sapa")
    start = 0
    mm = re.search(r"(?i:\b(?:cliente|azienda|ditta|societ[aà]|sig\.?ra|sig\.?|signor[ae]?|dott\.?|ing\.?|geom\.?)\s+)", t)
    if mm:
        start = len(re.findall(r"[A-Za-zÀ-ÿ'’]+", t[:mm.end()]))
    name = []
    i = start
    while i < len(toks):
        w = toks[i]
        lw = w.lower()
        if lw in skip and name:
            break
        if (w[:1].isupper() or lw in SUFF) and (lw not in IT_STOP or name or mm) and len(w) > 1:
            name.append(w)
        elif lw in PART and name and i + 1 < len(toks) and toks[i + 1][:1].isupper() and toks[i + 1].lower() not in skip:
            name.append(w)
        elif name:
            break
        elif mm:
            break
        i += 1
    while name and name[-1].lower() in PART:
        name.pop()
    out["cliente"] = " ".join(name[:5]).strip(" .,")
    manca = [n for n, k in (("nome del cliente", "cliente"), ("telefono", "telefono"), ("località", "localita")) if not out[k]]
    return out, manca


def _direttori(conn, area_id):
    return [r["id"] for r in conn.execute("SELECT id FROM users WHERE ruolo='direttore' AND attivo=1 AND area_id=?", (area_id,))]


@app.post("/api/proposte/assistente")
@need("proposte")
def api_proposte_assistente():
    u = current_user()
    testo = str(body().get("testo") or "").strip()
    if len(testo) < 5:
        return err("Scrivi qualche riga sull'opportunità (cliente, telefono, località, cosa chiede)")
    dati, manca = parse_chat(db(), u, testo)
    if manca:
        msg = "Ho letto la richiesta. Mi manca: %s. Puoi completarli qui sotto prima di inviare." % ", ".join(manca)
    else:
        msg = "Ho capito tutto. Controlla i dati e invia la proposta al Direttore di area."
    return jsonify({"dati": dati, "mancano": manca, "messaggio": msg})


@app.post("/api/proposte")
@need("proposte")
def api_proposte_crea():
    u = current_user()
    c = db()
    d = body()
    cliente = (d.get("cliente") or "").strip()[:200]
    if not cliente:
        return err("Indica il nome del cliente")
    area = u["area_id"]
    if area is None:
        area = d.get("area_id") if u["ruolo"] == "admin" else None
    if not area:
        return err("Il tuo profilo non appartiene a nessuna area")
    tel = (d.get("telefono") or "").strip()[:40]
    if not tel:
        return err("Indica un telefono del cliente")
    f = lambda k, n: (d.get(k) or "").strip()[:n]
    m = _indice(c, area, uid=u["id"]).cerca(cliente, dupcheck.piva_da_testo(f("note", 3000)), tel, ind=f("indirizzo", 300), loc=f("localita", 120))
    cur = c.execute("INSERT INTO proposals(area_id,proponente_id,cliente,telefono,localita,provincia,indirizzo,note,testo,origine,creata,dup_info) "
                    "VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                    (area, u["id"], cliente, tel, f("localita", 120), f("provincia", 2).upper(), f("indirizzo", 300), f("note", 3000),
                     f("testo", 3000), "assistente" if d.get("origine") == "assistente" else "manuale", ts(), dupcheck.riassunto(m)[:600]))
    pid = cur.lastrowid
    dest = _direttori(c, area) or [r["id"] for r in c.execute("SELECT id FROM users WHERE ruolo='admin' AND attivo=1")]
    notify(c, dest + [a for a in ancestors(c, u["id"]) if a not in dest],
           "Nuova proposta di %s da vagliare: %s" % (u["nome"], cliente), "proposta")
    c.commit()
    return jsonify({"ok": True, "id": pid, "dup_info": dupcheck.riassunto(m)})


def _prop_scope(u):
    if u["ruolo"] == "admin":
        return "1=1", []
    if u["ruolo"] == "direttore":
        return "p.area_id=?", [u["area_id"]]
    if u["ruolo"] == "ufficio_disdette":
        return "1=0", []
    if u["ruolo"] == "area_manager":
        ids = visible_user_ids(u)
        return "p.proponente_id IN (%s)" % ",".join("?" * len(ids)), ids
    ids = subtree_ids(u["id"])
    return "p.proponente_id IN (%s)" % ",".join("?" * len(ids)), ids


@app.get("/api/proposte")
@need("proposte")
def api_proposte_lista():
    u = current_user()
    w, ps = _prop_scope(u)
    rows = db().execute("SELECT p.*, a.nome AS proponente, d.nome AS deciso_da_nome FROM proposals p JOIN users a ON a.id=p.proponente_id "
                        "LEFT JOIN users d ON d.id=p.deciso_da WHERE " + w + " ORDER BY (p.stato='in_attesa') DESC, p.id DESC LIMIT 300", ps).fetchall()
    return jsonify({"proposte": [dict(r) for r in rows], "puo_vagliare": "approva" in perms_of(db(), u)})


@app.post("/api/proposte/<int:pid>/decisione")
@need("approva")
def api_proposte_decisione(pid):
    u = current_user()
    c = db()
    p = c.execute("SELECT * FROM proposals WHERE id=?", (pid,)).fetchone()
    if not p or (u["ruolo"] != "admin" and p["area_id"] != u["area_id"]):
        abort(404)
    if p["stato"] != "in_attesa":
        return err("Proposta già vagliata")
    d = body()
    if d.get("azione") == "rifiuta":
        motivo = (d.get("motivo") or "").strip()[:500]
        c.execute("UPDATE proposals SET stato='rifiutata', deciso_da=?, deciso=?, motivo=? WHERE id=?", (u["id"], ts(), motivo, pid))
        notify(c, [p["proponente_id"]], "Proposta rifiutata: %s%s" % (p["cliente"], " — " + motivo if motivo else ""), "proposta")
        c.commit()
        return jsonify({"ok": True})
    if d.get("azione") != "approva":
        return err("Azione non valida")
    aid = d.get("agent_id") or p["proponente_id"]
    t = c.execute("SELECT * FROM users WHERE id=? AND attivo=1 AND ruolo IN ('agente','recuperatore') AND area_id=?", (aid, p["area_id"])).fetchone()
    if not t:
        return err("Scegli un agente della stessa area a cui assegnare l'opportunità")
    cid = _email_campaign(c, u, p["area_id"], "lead")
    sla = sla_ore(c, "prospect", "lead")
    idx = comune_index(c)
    prov_ids = {r[0] for r in c.execute("SELECT id FROM geo_province WHERE sigla=?", (p["provincia"],))} if p["provincia"] else set()
    comune = trova_comune(c, idx, p["localita"] or "", p["indirizzo"] or "", user_prov_ids(c, t["id"]) | prov_ids, p["lat"], p["lon"])
    cur = c.execute(
        "INSERT INTO opportunities(campaign_id,area_id,cliente,telefono,localita,indirizzo,note,agente_id,stato,assegnata,scadenza,creata,comune_id,rif_esterno,dup_info,lat,lon) "
        "VALUES(?,?,?,?,?,?,?,?,'da_gestire',?,?,?,?,?,?,?,?)",
        (cid, p["area_id"], p["cliente"], p["telefono"], (p["localita"] + (" (%s)" % p["provincia"] if p["provincia"] else ""))[:140],
         p["indirizzo"], ("Proposta di %s\n%s" % ((c.execute("SELECT nome FROM users WHERE id=?", (p["proponente_id"],)).fetchone() or {"nome": ""})["nome"], p["note"] or "")).strip()[:4000],
         t["id"], ts(), ts(now() + timedelta(hours=sla)), ts(), comune, "PROPOSTA %d" % pid, p["dup_info"] or "", p["lat"], p["lon"]))
    oid = cur.lastrowid
    log_event(c, oid, u["id"], "creata", "proposta approvata")
    log_event(c, oid, u["id"], "assegnata", "a %s" % t["nome"])
    c.execute("UPDATE proposals SET stato='approvata', deciso_da=?, deciso=?, opp_id=? WHERE id=?", (u["id"], ts(), oid, pid))
    notify(c, [p["proponente_id"]] + cc_for_assignment(c, t["id"]), "Proposta approvata e assegnata a %s: %s" % (t["nome"], p["cliente"]), "assegnazione", oid)
    c.commit()
    return jsonify({"ok": True, "opp_id": oid})


# ----------------------------------------------------------------- ricerca prospect (assistente)
import prospect

CERCA_RE = re.compile(r"(?i)\b(cerca|cercami|trova|trovami|elenca|elencami|lista|liste|elenco|genera|generami|prospect|possibili clienti|potenziali clienti)\b")


def _luogo(conn, u, testo):
    """Comune citato nel testo (con coordinate), preferendo le zone dell'agente."""
    idx = comune_index(conn)
    mine = user_prov_ids(conn, u["id"])
    toks = re.findall(r"[A-Za-zÀ-ÿ'’]+", testo)
    best = None
    for n in (3, 2, 1):
        for i in range(len(toks) - n + 1):
            w = " ".join(toks[i:i + n])
            if n == 1 and (len(w) < 4 or w.lower() in IT_STOP):
                continue
            found = idx.get(geo_seed.key(w))
            if found:
                prep = i > 0 and toks[i - 1].lower() in ("a", "ad", "in", "di", "presso", "zona", "vicino", "comune", "intorno", "attorno")
                inz = any(pid in mine for _c, pid in found)
                sc = (n, prep, inz, w[:1].isupper())
                if (prep or inz or w[:1].isupper()) and (best is None or sc > best[0]):
                    best = (sc, found)
        if best and best[0][0] == n:
            break
    if not best:
        return None
    found = best[1]
    cid = next((c for c, p in found if p in mine), found[0][0])
    r = conn.execute("SELECT c.id,c.nome,c.lat,c.lon,p.sigla FROM geo_comuni c JOIN geo_province p ON p.id=c.provincia_id WHERE c.id=?", (cid,)).fetchone()
    return dict(r) if r and r["lat"] is not None else None


@app.post("/api/prospect/cerca")
@need("proposte")
def api_prospect_cerca():
    u = current_user()
    c = db()
    testo = str(body().get("testo") or "").strip()[:500]
    label, sel = prospect.categoria_da_testo(testo)
    if not sel:
        return err("Che tipo di attività cerchi? Ad esempio: mobilifici, ristoranti, hotel, officine, farmacie, gioiellerie, palestre, studi medici, aziende, magazzini…")
    luogo = _luogo(c, u, testo)
    if not luogo:
        return err("In quale comune? Scrivilo nella richiesta, ad esempio «mobilifici a Treviso» (il comune deve essere nell'archivio geografico).")
    km = prospect.raggio_da_testo(testo)
    try:
        data = prospect.overpass(prospect.query(luogo["lat"], luogo["lon"], km, sel))
    except RuntimeError as e:
        return err(str(e), 502)
    righe = prospect.parse(data, luogo["lat"], luogo["lon"])[:80]
    area = u["area_id"]
    ix = _indice(c, area, uid=u["id"])
    for r in righe:
        m = ix.cerca(r["nome"], r.get("piva"), r.get("telefono"), ind=r.get("indirizzo", ""), loc=r.get("citta", ""))
        r["livello"] = dupcheck.migliore(m)
        r["presente"] = r["livello"] in ("certo", "probabile")
        r["controllo"] = dupcheck.riassunto(m)
    return jsonify({"categoria": label, "comune": luogo["nome"], "provincia": luogo["sigla"], "raggio_km": km, "risultati": righe,
                    "fonte": "OpenStreetMap (dati aperti): i recapiti e la P.IVA ci sono solo se qualcuno li ha inseriti nella mappa"})


# ----------------------------------------------------------------- assistente AI (Anthropic)
import ai
import places


def _ai_key(conn):
    k = (os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("SOH_ANTHROPIC_KEY") or "").strip()
    if k:
        return k, "server"
    r = conn.execute("SELECT valore FROM settings WHERE chiave='ai_key'").fetchone()
    return ((r["valore"] or "").strip(), "impostazioni") if r and (r["valore"] or "").strip() else (None, None)


def _google_key(conn):
    k = (os.environ.get("GOOGLE_PLACES_KEY") or os.environ.get("SOH_GOOGLE_KEY") or "").strip()
    if k:
        return k, "server"
    r = conn.execute("SELECT valore FROM settings WHERE chiave='google_key'").fetchone()
    return ((r["valore"] or "").strip(), "impostazioni") if r and (r["valore"] or "").strip() else (None, None)


@app.get("/api/google/stato")
@need()
def api_google_stato():
    k, fonte = _google_key(db())
    u = current_user()
    d = {"attivo": bool(k), "fonte": fonte, "oggi": places.usate_oggi(), "limite": places.MAX_GIORNO}
    if k and u["ruolo"] == "admin":
        d["finale"] = "…" + k[-4:]
    return jsonify(d)


@app.put("/api/google/chiave")
@need("impostazioni")
def api_google_chiave():
    if current_user()["ruolo"] != "admin":
        return err("Solo l'amministratore può impostare la chiave Google", 403)
    k = str(body().get("chiave") or "").strip()
    c = db()
    if not k:
        c.execute("DELETE FROM settings WHERE chiave='google_key'")
    else:
        if not re.fullmatch(r"AIza[0-9A-Za-z_\-]{30,60}", k):
            return err("La chiave non sembra valida: una chiave Google inizia con «AIza» (la trovi in Google Cloud → API e servizi → Credenziali).")
        c.execute("INSERT OR REPLACE INTO settings(chiave,valore) VALUES('google_key',?)", (k,))
    c.commit()
    return api_google_stato()


@app.post("/api/google/prova")
@need("impostazioni")
def api_google_prova():
    if current_user()["ruolo"] != "admin":
        abort(403)
    k, _ = _google_key(db())
    if not k:
        return err("Nessuna chiave impostata")
    try:
        places.prova(k)
    except places.PlacesError as e:
        return err(str(e), 502)
    return jsonify({"ok": True})


@app.get("/api/ai/stato")
@need()
def api_ai_stato():
    k, fonte = _ai_key(db())
    u = current_user()
    d = {"attivo": bool(k), "fonte": fonte, "modello": ai.MODELLO}
    if k and u["ruolo"] == "admin":
        d["finale"] = "…" + k[-4:]
    return jsonify(d)


@app.put("/api/ai/chiave")
@need("impostazioni")
def api_ai_chiave():
    if current_user()["ruolo"] != "admin":
        return err("Solo l'amministratore può impostare la chiave dell'assistente AI", 403)
    k = str(body().get("chiave") or "").strip()
    c = db()
    if not k:
        c.execute("DELETE FROM settings WHERE chiave='ai_key'")
    else:
        if not re.fullmatch(r"sk-ant-[A-Za-z0-9_\-]{20,300}", k):
            return err("La chiave non sembra valida: deve iniziare con «sk-ant-» (la trovi nella console Anthropic, sezione API keys).")
        c.execute("INSERT OR REPLACE INTO settings(chiave,valore) VALUES('ai_key',?)", (k,))
    c.commit()
    return api_ai_stato()


@app.post("/api/ai/prova")
@need("impostazioni")
def api_ai_prova():
    if current_user()["ruolo"] != "admin":
        abort(403)
    k, _ = _ai_key(db())
    if not k:
        return err("Nessuna chiave impostata")
    try:
        ai.interpreta(k, "ristoranti a Verona", None, [x[1] for x in prospect.CATEGORIE])
    except ai.AIError as e:
        return err(str(e), 502)
    return jsonify({"ok": True, "modello": ai.MODELLO})


def _luoghi_da_ai(c, u, luoghi):
    """Converte i luoghi proposti dall'AI in comuni/province reali dell'archivio geografico (valori sconosciuti scartati)."""
    out, mine = [], user_prov_ids(c, u["id"])
    idx = None
    prov = c.execute("SELECT id,nome,sigla FROM geo_province").fetchall()
    for L in (luoghi or [])[:3]:
        if not isinstance(L, dict):
            continue
        nome = str(L.get("nome") or "").strip()[:60]
        k = geo_seed.key(nome)
        if len(k) < 2:
            continue
        pm = [r for r in prov if geo_seed.key(r["nome"]) == k or (r["sigla"] or "").lower() == nome.lower()]
        pm.sort(key=lambda r: r["id"] not in mine)
        if L.get("tipo") == "provincia" and pm:
            out.append({"tipo": "provincia", "nome": pm[0]["nome"], "sigla": pm[0]["sigla"] or ""})
            continue
        idx = idx or comune_index(c)
        found = idx.get(k)
        if found:
            cid = next((cc for cc, pp in found if pp in mine), found[0][0])
            r = c.execute("SELECT c.id,c.nome,c.lat,c.lon,p.sigla FROM geo_comuni c JOIN geo_province p ON p.id=c.provincia_id WHERE c.id=?", (cid,)).fetchone()
            if r and r["lat"] is not None:
                out.append({"tipo": "comune", "nome": r["nome"], "sigla": r["sigla"] or "", "lat": r["lat"], "lon": r["lon"]})
                continue
        if pm:  # nome di provincia usato come comune
            out.append({"tipo": "provincia", "nome": pm[0]["nome"], "sigla": pm[0]["sigla"] or ""})
    return out


def _cerca_google(chiave, luoghi, etichetta, raggio):
    """Cerca con Google Places: una ricerca per luogo e per categoria (max 3x3), con un tetto di richieste per ricerca."""
    cat = [x.strip() for x in str(etichetta).split(",") if x.strip()][:3] or ["attività"]
    budget, righe, tronc = [places.MAX_RICHIESTE_RICERCA], [], False
    for L in luoghi:
        for ct in cat:
            if L["tipo"] == "provincia":
                rr, t = places.cerca(chiave, "%s in provincia di %s" % (ct, L["nome"]), budget=budget)
            else:
                rr, t = places.cerca(chiave, "%s a %s%s" % (ct, L["nome"], " (%s)" % L["sigla"] if L.get("sigla") else ""),
                                     L["lat"], L["lon"], raggio, budget=budget)
            for r in rr:
                r["prov"] = L["sigla"]
            righe += rr
            tronc = tronc or t
    return righe, tronc


def _cerca_prospect(c, u, luoghi, sel, etichetta, raggio, solo_tel=False, solo_piva=False, limite=300):
    righe, visti, troncato = [], set(), False
    nota = ""
    luoghi0 = list(luoghi)
    gk, _f = _google_key(c)
    fonte_g = False
    if gk:
        try:
            righe, troncato = _cerca_google(gk, luoghi, etichetta, raggio)
            fonte_g = True
            visti = {r["osm"] for r in righe}
            uniq, vv = [], set()
            for r in righe:
                if r["osm"] not in vv:
                    vv.add(r["osm"])
                    uniq.append(r)
            righe = uniq
            luoghi = []
            if solo_piva:
                nota = "Google non fornisce la partita IVA: il filtro «solo con P.IVA» non è stato applicato."
                solo_piva = False
        except places.PlacesError as e:
            nota = "Google Places non disponibile (%s) Ho usato OpenStreetMap." % str(e).rstrip(".") + "."
            righe, visti, troncato = [], set(), False
    for L in luoghi:
        if L["tipo"] == "provincia":
            try:
                data = prospect.overpass(prospect.query_provincia(L["sigla"], L["nome"], sel, limite), timeout=45)
                rr = prospect.parse(data, None, None)
            except RuntimeError:
                # zona troppo grande per il servizio gratuito: ripiego sul capoluogo con raggio ampio
                cap = c.execute("SELECT c.nome,c.lat,c.lon FROM geo_comuni c JOIN geo_province p ON p.id=c.provincia_id "
                                "WHERE p.nome=? AND c.nome=? AND c.lat IS NOT NULL", (L["nome"], L["nome"])).fetchone()
                if not cap:
                    raise
                data = prospect.overpass(prospect.query(cap["lat"], cap["lon"], 15, sel, limit=limite))
                rr = prospect.parse(data, cap["lat"], cap["lon"])
                nota = ("La provincia intera è troppo grande per il servizio di ricerca gratuito in questo momento: "
                        "ho cercato entro 15 km da %s. Per le altre zone chiedi un comune alla volta." % cap["nome"])
        else:
            data = prospect.overpass(prospect.query(L["lat"], L["lon"], raggio, sel, limit=limite))
            rr = prospect.parse(data, L["lat"], L["lon"])
        troncato = troncato or len((data or {}).get("elements", [])) >= limite
        for r in rr:
            key = r["osm"]
            if key in visti:
                continue
            visti.add(key)
            r["prov"] = L["sigla"]
            righe.append(r)
    if solo_tel:
        righe = [r for r in righe if r["telefono"]]
    if solo_piva:
        righe = [r for r in righe if r["piva"]]
    righe = righe[:500]
    ix = _indice(c, u["area_id"], uid=u["id"])
    for r in righe:
        m = ix.cerca(r["nome"], r.get("piva"), r.get("telefono"), ind=r.get("indirizzo", ""), loc=r.get("citta", ""))
        r["livello"] = dupcheck.migliore(m)
        r["presente"] = r["livello"] in ("certo", "probabile")
        r["controllo"] = dupcheck.riassunto(m)
    desc = ", ".join(("provincia di " if L["tipo"] == "provincia" else "") + L["nome"] for L in luoghi0)
    return {"tipo": "lista", "categoria": etichetta, "comune": desc, "provincia": luoghi0[0]["sigla"], "luoghi_desc": desc,
            "raggio_km": raggio, "risultati": righe, "troncato": troncato, "nota": nota,
            "solo_tel": solo_tel, "solo_piva": solo_piva,
            "fonte": ("Google Places: nome, indirizzo, telefono e sito; la partita IVA non è disponibile" if fonte_g else
                      "OpenStreetMap (dati aperti): i recapiti e la P.IVA ci sono solo se qualcuno li ha inseriti nella mappa")}


def _dati_puliti(dati, testo):
    d = {k: str((dati or {}).get(k) or "").strip()[:300 if k in ("indirizzo", "note") else 120] for k in
         ("cliente", "telefono", "localita", "provincia", "indirizzo", "note")}
    d["provincia"] = re.sub(r"[^A-Za-z]", "", d["provincia"])[:2].upper()
    cifre = re.sub(r"\D", "", testo)
    t = re.sub(r"\D", "", d["telefono"])
    if t and t not in cifre:  # il numero deve comparire davvero nel testo: niente numeri inventati
        d["telefono"] = ""
    return d


def _assistente_ai(c, u, testo, spec):
    tipo = spec.get("tipo")
    if tipo == "domanda" or spec.get("serve_chiarimento"):
        return {"tipo": "messaggio", "messaggio": str(spec.get("serve_chiarimento") or spec.get("risposta") or "Puoi riformulare la richiesta?")[:400]}
    if tipo == "opportunita_singola":
        dati = _dati_puliti(spec.get("dati"), testo)
        manca = [n for k, n in (("cliente", "nome del cliente"), ("telefono", "telefono"), ("localita", "località")) if not dati[k]]
        msg = ("Ho letto la richiesta. Mi manca: %s. Puoi completarli qui sotto prima di inviare." % ", ".join(manca)) if manca \
            else "Ho capito tutto. Controlla i dati e invia la proposta."
        return {"tipo": "opportunita", "dati": dati, "mancano": manca, "messaggio": msg, "ai": True}
    sel, etich = [], []
    for lab in spec.get("categorie") or []:
        for keys, label, s_ in prospect.CATEGORIE:
            if label == lab and label not in etich:
                sel += s_
                etich.append(label)
    tag_ok = []
    for t in spec.get("osm_tag") or []:
        t = str(t).strip().lower()
        if ai.TAG_OSM.match(t):
            k, v = t.split("=")
            sel.append((k, [v]))
            tag_ok.append(t)
    if not sel:
        return {"tipo": "messaggio", "messaggio": "Che tipo di attività cerchi? Ad esempio: ristoranti, hotel, mobilifici, officine, studi medici…"}
    luoghi = _luoghi_da_ai(c, u, spec.get("luoghi"))
    if not luoghi:
        return {"tipo": "messaggio", "messaggio": "In quale comune o provincia? Non trovo il luogo nell'archivio geografico: prova a scriverlo per esteso."}
    try:
        raggio = max(1, min(15, int(spec.get("raggio_km") or 5)))
    except (TypeError, ValueError):
        raggio = 5
    etichetta = ", ".join(etich) or str(spec.get("categoria_etichetta") or "Attività")[:60]
    r = _cerca_prospect(c, u, luoghi, sel, etichetta, raggio, bool(spec.get("solo_con_telefono")), bool(spec.get("solo_con_piva")))
    r["ai"] = True
    r["risposta"] = str(spec.get("risposta") or "")[:300]
    r["interpretazione"] = {"categorie": etich, "osm_tag": tag_ok, "categoria_etichetta": etichetta,
                            "luoghi": [{"nome": L["nome"], "tipo": L["tipo"]} for L in luoghi], "raggio_km": raggio,
                            "solo_con_telefono": r["solo_tel"], "solo_con_piva": r["solo_piva"]}
    return r


def _assistente_semplice(c, u, testo):
    """Senza AI: riconoscimento con regole fisse (parole chiave)."""
    if CERCA_RE.search(testo):
        label, sel = prospect.categoria_da_testo(testo)
        if not sel:
            return {"tipo": "messaggio", "messaggio": "Che tipo di attività cerchi? Ad esempio: mobilifici, ristoranti, hotel, officine, farmacie, gioiellerie, palestre, studi medici, aziende, magazzini…"}
        luogo = _luogo(c, u, testo)
        if not luogo:
            return {"tipo": "messaggio", "messaggio": "In quale comune? Scrivilo nella richiesta, ad esempio «mobilifici a Treviso»."}
        L = {"tipo": "comune", "nome": luogo["nome"], "sigla": luogo["sigla"] or "", "lat": luogo["lat"], "lon": luogo["lon"]}
        return _cerca_prospect(c, u, [L], sel, label, prospect.raggio_da_testo(testo))
    dati, manca = parse_chat(c, u, testo)
    msg = ("Ho letto la richiesta. Mi manca: %s. Puoi completarli qui sotto prima di inviare." % ", ".join(manca)) if manca \
        else "Ho capito tutto. Controlla i dati e invia la proposta al Direttore di area."
    return {"tipo": "opportunita", "dati": dati, "mancano": manca, "messaggio": msg, "ai": False}


@app.post("/api/assistente")
@need("proposte")
def api_assistente():
    u = current_user()
    c = db()
    d = body()
    testo = str(d.get("testo") or "").strip()[:600]
    if len(testo) < 3:
        return err("Scrivi cosa cerchi o descrivi l'opportunità")
    prev = d.get("prev") if isinstance(d.get("prev"), dict) else None
    key, _ = _ai_key(c)
    avviso = None
    try:
        if key:
            try:
                ai.controlla_limite(u["id"])
                spec = ai.interpreta(key, testo, prev, [x[1] for x in prospect.CATEGORIE])
                return jsonify(_assistente_ai(c, u, testo, spec))
            except ai.AIError as e:
                avviso = str(e)
        r = _assistente_semplice(c, u, testo)
    except RuntimeError as e:  # servizio di ricerca (Overpass) non raggiungibile
        return err(str(e), 502)
    if avviso:
        r["avviso"] = "Assistente AI non disponibile (%s) Ho usato la ricerca semplice." % avviso
    return jsonify(r)


@app.post("/api/proposte/lista")
@need("proposte")
def api_proposte_lista_bulk():
    """Invia al Direttore più prospect trovati dalla ricerca (una proposta ciascuno)."""
    u = current_user()
    c = db()
    d = body()
    area = u["area_id"] or (d.get("area_id") if u["ruolo"] == "admin" else None)
    if not area:
        return err("Il tuo profilo non appartiene a nessuna area")
    items = d.get("items") or []
    if not isinstance(items, list) or not items:
        return err("Seleziona almeno un prospect")
    n = 0
    n_dup = 0
    ix = _indice(c, area, uid=u["id"])
    idx_c = comune_index(c)
    for it in items[:500]:
        nome = str((it or {}).get("nome") or "").strip()[:200]
        if not nome:
            continue
        piva = prospect.piva_valida(it.get("piva"))
        mm = ix.cerca(nome, piva, str(it.get("telefono") or ""), ind=str(it.get("indirizzo") or ""), loc=str(it.get("citta") or ""))
        n_dup += 1 if mm else 0
        ix.add("prop", nome, piva or "", str(it.get("telefono") or ""), "Già proposto da %s il %s (in attesa del Direttore)" % (_nm(u["nome"]), _dmy(ts())), ind=str(it.get("indirizzo") or ""), loc=str(it.get("citta") or ""))
        note = "Prospect da ricerca (OpenStreetMap)%s%s%s" % ("\nP.IVA: " + piva if piva else "", "\nSito: " + str(it.get("sito"))[:200] if it.get("sito") else "",
                                                          "\nEmail: " + str(it.get("email"))[:120] if it.get("email") else "")
        la, lo = _num(it.get("lat")), _num(it.get("lon"))
        cid_ = trova_comune(c, idx_c, str(it.get("citta") or ""), str(it.get("indirizzo") or ""), set(), la, lo)
        nome_c, sig_c = "", ""
        if cid_:
            rr_ = c.execute("SELECT c.nome,p.sigla FROM geo_comuni c JOIN geo_province p ON p.id=c.provincia_id WHERE c.id=?", (cid_,)).fetchone()
            nome_c, sig_c = rr_["nome"], rr_["sigla"] or ""
        loc_ = str(it.get("citta") or nome_c or d.get("comune") or "")[:120]
        c.execute("INSERT INTO proposals(area_id,proponente_id,cliente,telefono,localita,provincia,indirizzo,note,testo,origine,creata,dup_info,lat,lon) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                  (area, u["id"], nome, str(it.get("telefono") or "")[:40], loc_,
                   str(it.get("prov") or sig_c or d.get("provincia") or "")[:2].upper(), str(it.get("indirizzo") or "")[:300], note[:3000], "", "ricerca", ts(),
                   dupcheck.riassunto(mm)[:600], la, lo))
        n += 1
    if not n:
        return err("Nessun prospect valido")
    dest = _direttori(c, area) or [r["id"] for r in c.execute("SELECT id FROM users WHERE ruolo='admin' AND attivo=1")]
    notify(c, dest + [a for a in ancestors(c, u["id"]) if a not in dest], "%s propone %d possibili prospect da vagliare" % (u["nome"], n), "proposta")
    c.commit()
    return jsonify({"ok": True, "inviati": n, "con_avvisi": n_dup})


@app.post("/api/spegni")
@need("impostazioni")
def api_spegni():
    """Spegne il programma (solo uso locale sul proprio PC, solo amministratore)."""
    if request.host.split(":")[0] not in ("127.0.0.1", "localhost") or current_user()["ruolo"] != "admin":
        abort(403)
    threading.Timer(0.8, lambda: os._exit(0)).start()
    return jsonify({"ok": True})


# ----------------------------------------------------------------- calendario (.ics) per Outlook & co.
def _ics_esc(t):
    return (str(t or "").replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,")
            .replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\\n"))


def _ics_fold(line):
    """Righe lunghe max 75 byte (RFC 5545), spezzando solo tra caratteri interi."""
    b, out = line.encode("utf-8"), []
    while len(b) > 75:
        cut = 75
        while (b[cut] & 0xC0) == 0x80:
            cut -= 1
        out.append(b[:cut].decode("utf-8"))
        b = b" " + b[cut:]
    out.append(b.decode("utf-8"))
    return "\r\n".join(out)


def _ics_dt(local_str):
    """'YYYY-MM-DD HH:MM:SS' (ora italiana) -> 'YYYYMMDDTHHMMSSZ' (UTC)."""
    d = datetime.strptime(local_str[:19], "%Y-%m-%d %H:%M:%S")
    if TZ:
        from datetime import timezone
        return d.replace(tzinfo=TZ).astimezone(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    return d.strftime("%Y%m%dT%H%M%S")


def _ics_event(o, kind, host="sicuritalia-hub"):
    """kind: 'app' = appuntamento (1 ora), 'sla' = scadenza SLA (15 minuti)."""
    start = o["appuntamento"] if kind == "app" else o["scadenza"]
    d0 = datetime.strptime(start[:19], "%Y-%m-%d %H:%M:%S")
    end = ts(d0 + timedelta(hours=1 if kind == "app" else 0, minutes=0 if kind == "app" else 15))
    titolo = ("Appuntamento: " if kind == "app" else "Scadenza SLA: ") + (o["cliente"] or "")
    luogo = ", ".join(x for x in [(o["indirizzo"] or "").strip(), (o["localita"] or "").strip()] if x)
    desc = []
    if o["telefono"]:
        desc.append("Telefono: " + o["telefono"])
    if o["campagna"]:
        desc.append("Campagna: " + o["campagna"])
    if o["rif_esterno"]:
        desc.append("Riferimento: " + o["rif_esterno"])
    if o["note"]:
        desc.append("Note: " + o["note"])
    desc.append("Aperto da Sicuritalia Opportunity Hub, opportunità n. %d" % o["id"])
    lines = ["BEGIN:VEVENT", "UID:opp%d-%s@%s" % (o["id"], kind, host),
             "DTSTAMP:" + datetime.utcnow().strftime("%Y%m%dT%H%M%SZ"),
             "DTSTART:" + _ics_dt(start), "DTEND:" + _ics_dt(end),
             "SUMMARY:" + _ics_esc(titolo)]
    if luogo:
        lines.append("LOCATION:" + _ics_esc(luogo))
    lines.append("DESCRIPTION:" + _ics_esc("\n".join(desc)))
    lines.append("CATEGORIES:" + ("Appuntamento" if kind == "app" else "Scadenza SLA"))
    lines += ["BEGIN:VALARM", "ACTION:DISPLAY", "DESCRIPTION:" + _ics_esc(titolo),
              "TRIGGER:-PT30M" if kind == "app" else "TRIGGER:-PT60M", "END:VALARM", "END:VEVENT"]
    return lines


def _ics_wrap(eventi, nome):
    head = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Sicuritalia//Opportunity Hub//IT", "CALSCALE:GREGORIAN",
            "METHOD:PUBLISH", "X-WR-CALNAME:" + _ics_esc(nome), "X-WR-TIMEZONE:Europe/Rome",
            "REFRESH-INTERVAL;VALUE=DURATION:PT1H", "X-PUBLISHED-TTL:PT1H"]
    lines = head + [l for ev in eventi for l in ev] + ["END:VCALENDAR"]
    return "\r\n".join(_ics_fold(l) for l in lines) + "\r\n"


def _ics_response(testo, nome_file):
    r = app.response_class(testo, mimetype="text/calendar")
    r.headers["Content-Disposition"] = 'inline; filename="%s"' % nome_file
    r.headers["Cache-Control"] = "no-store"
    r.headers["X-Robots-Tag"] = "noindex, nofollow"
    return r


def _cal_token(u):
    c = db()
    t = c.execute("SELECT cal_token FROM users WHERE id=?", (u["id"],)).fetchone()["cal_token"]
    if not t:
        t = secrets.token_urlsafe(24)
        c.execute("UPDATE users SET cal_token=? WHERE id=?", (t, u["id"]))
        c.commit()
    return t


@app.get("/api/calendario/link")
@need("agenda")
def api_cal_link():
    t = _cal_token(current_user())
    base = request.host_url.rstrip("/")
    url = "%s/cal/%s.ics" % (base, t)
    return jsonify({"url": url, "https": url.startswith("https://"),
                    "locale": request.host.split(":")[0] in ("127.0.0.1", "localhost")})


@app.post("/api/calendario/rigenera")
@need("agenda")
def api_cal_rigenera():
    u = current_user()
    db().execute("UPDATE users SET cal_token=? WHERE id=?", (secrets.token_urlsafe(24), u["id"]))
    db().commit()
    return api_cal_link()


@app.get("/cal/<token>.ics")
def cal_feed(token):
    """Feed pubblico (senza login) protetto dal codice segreto nell'indirizzo."""
    if len(token) < 20:
        abort(404)
    c = db()
    u = c.execute("SELECT * FROM users WHERE cal_token=? AND attivo=1", (token,)).fetchone()
    if not u or "agenda" not in perms_of(c, u):
        abort(404)
    da = ts(now() - timedelta(days=30))
    rows = c.execute(
        BASE_SELECT + " WHERE o.agente_id=? AND ((o.stato='appuntamento' AND o.appuntamento>=?) "
        "OR (o.stato='da_gestire' AND o.scadenza IS NOT NULL AND o.scadenza>=?)) ORDER BY o.id",
        (u["id"], da, ts(now() - timedelta(days=1)))).fetchall()
    ev = []
    for o in rows:
        ev.append(_ics_event(o, "app" if o["stato"] == "appuntamento" else "sla"))
    return _ics_response(_ics_wrap(ev, "Sicuritalia - " + u["nome"]), "sicuritalia-agenda.ics")


@app.get("/api/opportunities/<int:oid>/ics")
@need("opportunita")
def api_opp_ics(oid):
    o = load_opp(oid, current_user())
    if o["stato"] != "appuntamento" or not o["appuntamento"]:
        return err("Questa opportunità non ha un appuntamento fissato", 400)
    return _ics_response(_ics_wrap([_ics_event(o, "app")], "Appuntamento"), "appuntamento-%d.ics" % oid)


@app.get("/healthz")
def healthz():
    return jsonify({"ok": True})


if os.environ.get("SOH_BEHIND_PROXY") == "1":  # dietro reverse proxy / piattaforma cloud (HTTPS terminato a monte)
    from werkzeug.middleware.proxy_fix import ProxyFix
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

init_db()

_sla_started = False


def start_background():
    """Controllo SLA periodico (con gunicorn usare 1 worker + più thread)."""
    global _sla_started
    if not _sla_started:
        _sla_started = True
        threading.Thread(target=_sla_loop, daemon=True).start()
        threading.Thread(target=_geo_autoload, daemon=True).start()


if __name__ != "__main__":
    start_background()

if __name__ == "__main__":
    if "--demo" in sys.argv:
        import demo_data
        demo_data.run()
    start_background()
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", 8000))
    print(">> Sales Opportunity Hub su http://%s:%d" % (host, port))
    app.run(host=host, port=port, threaded=True)
