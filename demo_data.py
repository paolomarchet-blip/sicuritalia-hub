"""Dati dimostrativi: gerarchia per l'area Triveneto e Lombardia Est + opportunità di esempio.
Tutte le password demo sono 'demo1234' (cambio richiesto al primo accesso disattivato per comodità)."""
import random
from datetime import timedelta

from werkzeug.security import generate_password_hash

import app as A

CITTA = ["Verona", "Padova", "Treviso", "Vicenza", "Venezia", "Brescia", "Trento", "Bolzano", "Udine", "Mantova"]
NOMI = ["Rossi", "Bianchi", "Ferrari", "Esposito", "Romano", "Colombo", "Ricci", "Marino", "Greco", "Bruno",
        "Gallo", "Conti", "De Luca", "Costa", "Giordano", "Mancini", "Rizzo", "Lombardi", "Moretti", "Barbieri"]
SUFF = ["Srl", "SNC", "& Figli", "Immobiliare", "Autotrasporti", "Studio", "Ristorante", "Officina"]


def run():
    c = A.connect()
    if c.execute("SELECT COUNT(*) FROM users WHERE ruolo<>'admin'").fetchone()[0]:
        print(">> Dati demo già presenti")
        c.close()
        return
    rnd = random.Random(7)
    pw = generate_password_hash("demo1234")
    aree = {r["nome"]: r["id"] for r in c.execute("SELECT id,nome FROM areas")}
    tv = aree["Triveneto e Lombardia Est"]
    now = A.ts()

    def add(user, nome, ruolo, area, parent=None, terr=""):
        cur = c.execute("INSERT INTO users(username,password_hash,nome,ruolo,area_id,parent_id,territori,must_change,creato) "
                        "VALUES(?,?,?,?,?,?,?,0,?)", (user, pw, nome, ruolo, area, parent, terr, now))
        return cur.lastrowid

    for nome, aid in aree.items():
        if aid != tv:
            add("dir." + A.norm(nome).split()[0], "Direttore " + nome, "direttore", aid)
    dirv = add("direttore", "Direttore Triveneto", "direttore", tv)
    d1 = add("district1", "District Nord-Est", "district", tv, dirv)
    d2 = add("district2", "District Lombardia Est", "district", tv, dirv)
    t1 = add("tl1", "Team Leader Verona", "team_leader", tv, d1)
    t2 = add("tl2", "Team Leader Padova", "team_leader", tv, d1)
    t3 = add("tl3", "Team Leader Brescia", "team_leader", tv, d2)
    agenti = [
        add("agente1", "Luca Moretti", "agente", tv, t1, "Verona,Mantova"),
        add("agente2", "Sara Fontana", "agente", tv, t1, "Verona,Vicenza"),
        add("agente3", "Marco Galli", "agente", tv, t2, "Padova,Venezia"),
        add("agente4", "Elena Villa", "agente", tv, t2, "Treviso,Venezia"),
        add("agente5", "Paolo Serra", "agente", tv, t3, "Brescia,Mantova"),
        add("agente6", "Giulia Neri", "agente", tv, d2, "Brescia"),
    ]
    rec = add("recuperatore1", "Davide Pace (recuperatore)", "recuperatore", tv, dirv)

    # zone di lavoro geografiche degli agenti (regione > province > comuni) e comuni da cui pescare le opportunità demo
    TERR = ["Verona,Mantova", "Verona,Vicenza", "Padova,Venezia", "Treviso,Venezia", "Brescia,Mantova", "Brescia"]
    ZPOOL = {rec: list(CITTA)}
    for uid, cities in zip(agenti, TERR):
        provs, regs, pool = set(), set(), []
        for ct in cities.split(","):
            r = c.execute("SELECT gc.provincia_id pid, gp.regione_id rid FROM geo_comuni gc JOIN geo_province gp ON gp.id=gc.provincia_id "
                          "WHERE gc.nome=?", (ct,)).fetchone()
            if r:
                provs.add(r["pid"])
                regs.add(r["rid"])
        for pid in provs:
            for r in c.execute("SELECT id,nome FROM geo_comuni WHERE provincia_id=?", (pid,)):
                c.execute("INSERT OR IGNORE INTO user_geo(user_id,tipo,geo_id) VALUES(?,?,?)", (uid, "comune", r["id"]))
                pool.append(r["nome"])
        for rid in regs:
            c.execute("INSERT OR IGNORE INTO user_geo(user_id,tipo,geo_id) VALUES(?,?,?)", (uid, "regione", rid))
        for pid in provs:
            c.execute("INSERT OR IGNORE INTO user_geo(user_id,tipo,geo_id) VALUES(?,?,?)", (uid, "provincia", pid))
        ZPOOL[uid] = pool or list(CITTA)

    def camp(nome, cat, sub, n, pool):
        cid = c.execute("INSERT INTO campaigns(nome,categoria,sottotipo,area_id,creato_da,creato,regola,totale) "
                        "VALUES(?,?,?,?,?,?,?,?)", (nome, cat, sub, tv, dirv, now, "territorio", n)).lastrowid
        for i in range(n):
            ag = pool[i % len(pool)]
            citta = rnd.choice(ZPOOL[ag])
            cli = "%s %s" % (rnd.choice(NOMI), rnd.choice(SUFF))
            età_ore = rnd.choice([2, 6, 12, 20, 30, 40, 55, 70, 90])
            ass = A.now() - timedelta(hours=età_ore)
            stato, appt, gest, nota, rip = "da_gestire", None, None, "", 0
            roll = rnd.random()
            if età_ore < 24 and roll < 0.25 or età_ore >= 24 and roll < 0.7:
                gest = A.ts(ass + timedelta(hours=rnd.randint(1, 20)))
                stato = rnd.choices(["appuntamento", "contratto", "ko", "da_riprogrammare"], [4, 3, 2, 1.5])[0]
                if stato == "appuntamento":
                    d = A.now().date() + timedelta(days=rnd.randint(1, 8))
                    while d.weekday() >= 5:
                        d += timedelta(days=1)
                    appt = "%s %s:00" % (d.isoformat(), rnd.choice(A.SLOT_ORARI))
                elif stato == "ko":
                    nota = rnd.choice(["Non interessato", "Ha già un contratto", "Irreperibile"])
                elif stato == "contratto":
                    nota = "Firmato in sede"
                elif stato == "da_riprogrammare":
                    rip, nota = 1, "Richiamare la prossima settimana"
            cur = c.execute(
                "INSERT INTO opportunities(campaign_id,area_id,cliente,telefono,localita,indirizzo,agente_id,stato,esito_nota,"
                "assegnata,scadenza,appuntamento,gestita,n_riprogrammazioni,creata) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (cid, tv, cli, "3%d%07d" % (rnd.randint(20, 99), rnd.randint(0, 9999999)), citta, "Via Roma %d" % rnd.randint(1, 120),
                 ag, stato, nota, A.ts(ass), A.ts(ass + timedelta(hours=48)), appt, gest, rip, A.ts(ass)))
            if gest:
                A.log_event(c, cur.lastrowid, ag, stato, nota)

    camp("Rinnovi autunno 2026", "customer_base", "", 36, agenti)
    camp("Lead sito web ottobre", "prospect", "lead", 24, agenti)
    camp("Nuove nate Q4", "prospect", "nuova_nata", 14, agenti)
    camp("Disdette novembre", "disdetta", "", 10, [rec])
    c.commit()
    A.resolve_missing(c)
    A.check_sla(c)
    c.close()
    print(">> Dati demo creati. Accessi (password 'demo1234'): direttore, district1, district2, tl1, tl2, tl3, "
          "agente1..agente6, recuperatore1; amministratore: admin / admin")
