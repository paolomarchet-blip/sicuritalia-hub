"""Caricamento dell'elenco completo dei comuni italiani (regione > provincia > comune, con coordinate).

Fonti supportate:
 - GeoNames "IT.zip" / "IT.txt" (codici postali, licenza CC-BY 3.0, www.geonames.org) - scaricabile in automatico;
 - un file CSV/TSV con colonne tipo: regione; provincia; sigla; comune; lat; lon (anche l'elenco ISTAT, senza coordinate).
"""
import csv
import io
import os
import urllib.request
import zipfile

from geo_seed import key

GEONAMES_URL = os.environ.get("SOH_GEO_URL", "https://download.geonames.org/export/zip/IT.zip")


def _decode(b):
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    return b.decode("utf-8", errors="replace")


def _num(x):
    try:
        return float(str(x).strip().replace(",", "."))
    except ValueError:
        return None


def parse_geonames(text):
    acc = {}
    for line in text.splitlines():
        c = line.split("\t")
        if len(c) < 11 or c[0] != "IT":
            continue
        regione, prov, sigla, comune = c[3].strip(), c[5].strip(), c[6].strip().upper(), c[2].strip()
        lat, lon = _num(c[9]), _num(c[10])
        if not (regione and comune and (prov or sigla)):
            continue
        k = (sigla or key(prov), key(comune))
        a = acc.setdefault(k, {"regione": regione, "provincia": prov, "sigla": sigla, "comune": comune, "lat": [], "lon": []})
        if lat is not None and lon is not None and 35 < lat < 48 and 6 < lon < 19:
            a["lat"].append(lat)
            a["lon"].append(lon)
    out = []
    for a in acc.values():
        a["lat"] = sum(a["lat"]) / len(a["lat"]) if a["lat"] else None
        a["lon"] = sum(a["lon"]) / len(a["lon"]) if a["lon"] else None
        out.append(a)
    return out


def parse_csv(text):
    first = text.splitlines()[0] if text.strip() else ""
    delim = max((";", "\t", ","), key=lambda d: first.count(d))
    rows = list(csv.reader(io.StringIO(text), delimiter=delim))
    if len(rows) < 2:
        return []
    hdr = [key(h) for h in rows[0]]

    def find(*needles, exclude=()):
        for n in needles:
            for i, h in enumerate(hdr):
                if n in h and not any(e in h for e in exclude):
                    return i
        return None
    ci = {"regione": find("denominazioneregione", "regione"),
          "provincia": find("unitaterritorialesovracomunale", "provincia", exclude=("sigla", "codice")),
          "sigla": find("sigla"),
          "comune": find("denominazioneinitaliano", "comune", "denominazione", exclude=("regione", "unita", "provincia")),
          "lat": find("latitudine", "lat"), "lon": find("longitudine", "lng", "lon")}
    if ci["comune"] is None or ci["regione"] is None or (ci["provincia"] is None and ci["sigla"] is None):
        raise ValueError("Colonne non riconosciute: servono Regione, Provincia (o Sigla) e Comune")
    out = []
    for r in rows[1:]:
        def g(k):
            i = ci[k]
            return r[i].strip() if i is not None and i < len(r) else ""
        if not g("comune"):
            continue
        out.append({"regione": g("regione"), "provincia": g("provincia"), "sigla": g("sigla").upper(), "comune": g("comune"),
                    "lat": _num(g("lat")) if ci["lat"] is not None else None, "lon": _num(g("lon")) if ci["lon"] is not None else None})
    return out


def parse_any(filename, data):
    name = (filename or "").lower()
    if name.endswith(".zip") or data[:2] == b"PK":
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            names = [n for n in z.namelist() if n.lower().endswith((".txt", ".csv", ".tsv")) and "readme" not in n.lower()]
            if not names:
                raise ValueError("Nel file zip non c'è nessun elenco leggibile")
            data = z.read(sorted(names, key=lambda n: (not n.upper().endswith("IT.TXT"), n))[0])
    text = _decode(data)
    head = next((l for l in text.splitlines() if l.strip()), "")
    if head.startswith("IT\t") and head.count("\t") >= 10:
        return parse_geonames(text)
    return parse_csv(text)


def download():
    req = urllib.request.Request(GEONAMES_URL, headers={"User-Agent": "SalesOpportunityHub/1.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def import_rows(conn, rows):
    """Aggiunge/aggiorna regioni, province e comuni senza mai cancellare (le zone già assegnate restano valide)."""
    reg = {key(r["nome"]): r["id"] for r in conn.execute("SELECT id,nome FROM geo_regioni")}
    prov_sigla = {r["sigla"]: r["id"] for r in conn.execute("SELECT id,sigla FROM geo_province WHERE sigla<>''")}
    prov_key = {(r["regione_id"], key(r["nome"])): r["id"] for r in conn.execute("SELECT id,nome,regione_id FROM geo_province")}
    com = {(r["provincia_id"], key(r["nome"])): (r["id"], r["lat"]) for r in conn.execute("SELECT id,nome,provincia_id,lat FROM geo_comuni")}
    nuovi = aggiornati = 0
    for r in rows:
        rk = key(r["regione"])
        if not rk:
            continue
        if rk not in reg:
            reg[rk] = conn.execute("INSERT INTO geo_regioni(nome) VALUES(?)", (r["regione"],)).lastrowid
        rid = reg[rk]
        pid = prov_sigla.get(r["sigla"]) if r["sigla"] else None
        if pid is None:
            pid = prov_key.get((rid, key(r["provincia"])))
        if pid is None:
            nome_p = r["provincia"] or r["sigla"]
            if not nome_p:
                continue
            pid = conn.execute("INSERT INTO geo_province(nome,sigla,regione_id) VALUES(?,?,?)", (nome_p, r["sigla"] or "", rid)).lastrowid
            prov_key[(rid, key(nome_p))] = pid
            if r["sigla"]:
                prov_sigla[r["sigla"]] = pid
        ck = (pid, key(r["comune"]))
        if ck in com:
            cid, lat0 = com[ck]
            if r["lat"] is not None:
                conn.execute("UPDATE geo_comuni SET lat=?, lon=? WHERE id=?", (r["lat"], r["lon"], cid))
                aggiornati += 1
        else:
            cid = conn.execute("INSERT INTO geo_comuni(nome,provincia_id,lat,lon) VALUES(?,?,?,?)",
                               (r["comune"], pid, r["lat"], r["lon"])).lastrowid
            com[ck] = (cid, r["lat"])
            nuovi += 1
    return {"nuovi": nuovi, "aggiornati": aggiornati}
