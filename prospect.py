"""Ricerca di possibili prospect (attività) su OpenStreetMap tramite Overpass: nome, indirizzo, telefono, sito, P.IVA se censita."""
import json
import math
import re
import socket
import time
import urllib.error
import urllib.parse
import urllib.request

SERVERS = ["https://overpass-api.de/api/interpreter", "https://overpass.kumi.systems/api/interpreter",
           "https://overpass.private.coffee/api/interpreter"]

# parole chiave italiane -> (etichetta, selettori [(chiave, [valori] | None)])
CATEGORIE = [
    (("mobilific", "arredament", "mobili", "cucine"), "Mobilifici e arredamento", [("shop", ["furniture", "interior_decoration", "kitchen", "bed", "carpet"])]),
    (("ristorant", "pizzeri", "trattori", "osteri"), "Ristoranti e pizzerie", [("amenity", ["restaurant", "fast_food"])]),
    (("hotel", "alberghi", "albergo", "b&b", "bed and breakfast", "agriturism", "residence"), "Hotel e strutture ricettive", [("tourism", ["hotel", "guest_house", "motel", "apartment", "chalet"])]),
    (("farmaci",), "Farmacie", [("amenity", ["pharmacy"])]),
    (("gioiellerie", "oreficeri", "gioiell"), "Gioiellerie", [("shop", ["jewelry"])]),
    (("supermerc", "alimentari", "minimarket", "market"), "Supermercati e alimentari", [("shop", ["supermarket", "convenience", "greengrocer", "butcher", "bakery"])]),
    (("officin", "autoriparaz", "gommist", "carrozzer"), "Officine e autoriparazioni", [("shop", ["car_repair", "tyres"]), ("craft", ["car_repair"])]),
    (("concessionari", "autosalon", "automobil"), "Concessionarie auto", [("shop", ["car", "motorcycle", "car_parts"])]),
    (("banche", "banca", "assicuraz"), "Banche e assicurazioni", [("amenity", ["bank"]), ("office", ["insurance", "financial"])]),
    (("dentist", "studi medici", "medici", "ambulator", "cliniche", "clinica", "veterinar"), "Studi medici e dentisti", [("amenity", ["dentist", "doctors", "clinic", "veterinary"])]),
    (("palestr", "fitness", "centri sportivi"), "Palestre e centri sportivi", [("leisure", ["fitness_centre", "sports_centre"])]),
    (("bar", "caffe", "caffè", "pub", "pasticceri"), "Bar e caffetterie", [("amenity", ["bar", "cafe", "pub"]), ("shop", ["pastry"])]),
    (("parrucch", "estetist", "barbier", "benessere", "centri estetici"), "Parrucchieri ed estetisti", [("shop", ["hairdresser", "beauty"])]),
    (("abbigliament", "calzatur", "boutique"), "Abbigliamento e calzature", [("shop", ["clothes", "shoes", "boutique", "fashion"])]),
    (("tabacch",), "Tabaccherie", [("shop", ["tobacco"])]),
    (("scuol", "asili", "asilo"), "Scuole e asili", [("amenity", ["school", "kindergarten"])]),
    (("magazzin", "logistic", "capannon", "depositi"), "Magazzini e logistica", [("industrial", ["warehouse", "logistics"]), ("building", ["warehouse"])]),
    (("artigian", "falegnam", "idraulic", "elettricist", "imprese edili", "edili"), "Artigiani e imprese", [("craft", None)]),
    (("aziende", "azienda", "industri", "imprese", "societ", "fabbrich", "stabilimenti"), "Aziende e industrie", [("office", ["company"]), ("industrial", None), ("man_made", ["works"])]),
    (("uffici", "studi professionali", "studi", "commercialist", "avvocat", "notai"), "Uffici e studi professionali", [("office", None)]),
    (("negozi", "negozio", "commerci", "attività commerciali", "attivita commerciali"), "Negozi", [("shop", None)]),
]


def categoria_da_testo(testo):
    t = (testo or "").lower()
    for keys, label, sel in CATEGORIE:
        for k in keys:
            k = k.strip()
            pat = (r"\b" + re.escape(k) + r"\b") if len(k) <= 3 else (r"(?<![a-zà-ÿ])" + re.escape(k))
            if re.search(pat, t):
                return label, sel
    return None, None


def raggio_da_testo(testo, default=3):
    m = re.search(r"(?:entro|raggio(?: di)?|nel raggio di|a)\s*(\d{1,2})\s*km", (testo or "").lower())
    return max(1, min(15, int(m.group(1)))) if m else default


def _clausole(selettori):
    """Un filtro per ogni coppia chiave=valore (più veloce del filtro con espressione regolare)."""
    out = []
    for key, vals in selettori:
        if vals is None:
            out.append('["%s"]' % key)
        else:
            out += ['["%s"="%s"]' % (key, v) for v in vals]
    return out


def query(lat, lon, raggio_km, selettori, limit=250):
    parts = ["nwr%s(around:%d,%.5f,%.5f);" % (f, int(raggio_km * 1000), lat, lon) for f in _clausole(selettori)]
    return "[out:json][timeout:40];(%s);out center tags %d;" % ("".join(parts), limit)


def query_provincia(sigla, nome, selettori, limit=300):
    """Tutte le attività di una provincia intera (confine amministrativo di OpenStreetMap)."""
    sig = re.sub(r"[^A-Z]", "", (sigla or "").upper())[:2]
    nm = re.sub(r"[^A-Za-zÀ-ÿ '\-]", "", nome or "")[:60]
    aree = []
    if sig:
        aree.append('area["ISO3166-2"="IT-%s"];' % sig)
    if nm:
        aree.append('area["boundary"="administrative"]["admin_level"="6"]["name"="%s"];' % nm)
    parts = ["nwr%s(area.a);" % f for f in _clausole(selettori)]
    return "[out:json][timeout:60];(%s)->.a;(%s);out center tags %d;" % ("".join(aree), "".join(parts), limit)


def piva_valida(v):
    v = re.sub(r"\D", "", re.sub(r"^IT", "", (v or "").strip().upper()))
    if len(v) != 11:
        return ""
    s = 0
    for i, ch in enumerate(v[:10]):
        d = int(ch)
        if i % 2:
            d *= 2
            if d > 9:
                d -= 9
        s += d
    return v if (10 - s % 10) % 10 == int(v[10]) else ""


def _dist(lat1, lon1, lat2, lon2):
    p = math.pi / 180
    a = math.sin((lat2 - lat1) * p / 2) ** 2 + math.cos(lat1 * p) * math.cos(lat2 * p) * math.sin((lon2 - lon1) * p / 2) ** 2
    return 12742 * math.asin(math.sqrt(a))


def parse(data, lat, lon):
    out, seen = [], set()
    for e in (data or {}).get("elements", []):
        t = e.get("tags") or {}
        nome = (t.get("name") or t.get("brand") or "").strip()
        if not nome:
            continue
        la = e.get("lat") or (e.get("center") or {}).get("lat")
        lo = e.get("lon") or (e.get("center") or {}).get("lon")
        via = " ".join(x for x in (t.get("addr:street"), t.get("addr:housenumber")) if x)
        citta = t.get("addr:city") or t.get("addr:town") or t.get("addr:village") or ""
        cap = t.get("addr:postcode") or ""
        ind = ", ".join(x for x in (via, (cap + " " + citta).strip()) if x)
        piva = ""
        for k in ("ref:vatin", "vatin", "contact:vatin", "vat", "ref:VAT", "vat:id"):
            if t.get(k):
                piva = piva_valida(t[k]) or ""
                if piva:
                    break
        tel = (t.get("phone") or t.get("contact:phone") or t.get("contact:mobile") or "").split(";")[0].strip()
        key = (nome.lower(), via.lower())
        if key in seen:
            continue
        seen.add(key)
        out.append({"nome": nome, "indirizzo": ind, "via": via, "citta": citta, "cap": cap, "telefono": tel,
                    "sito": (t.get("website") or t.get("contact:website") or "").strip(), "email": (t.get("email") or t.get("contact:email") or "").strip(),
                    "piva": piva, "osm": "%s/%s" % (e.get("type"), e.get("id")), "lat": la, "lon": lo,
                    "km": round(_dist(lat, lon, la, lo), 1) if (la and lo and lat is not None and lon is not None) else None})
    out.sort(key=lambda r: (r["km"] is None, r["km"] or 0, (r["citta"] or "").lower(), r["nome"].lower()))
    return out


_cache = {}
_CACHE_SEC = 1800


def overpass(q, timeout=40):
    """Esegue la query provando più server; risultati ripetuti restano in memoria 30 minuti.
    Solleva RuntimeError con un messaggio comprensibile."""
    now = time.time()
    for k in [k for k, (t, _d) in _cache.items() if now - t > _CACHE_SEC]:
        _cache.pop(k, None)
    if q in _cache:
        return _cache[q][1]
    inizio = time.time()
    lenti, altri = 0, []
    for url in SERVERS:
        resto = timeout + 20 - (time.time() - inizio)
        if resto < 8:
            break
        try:
            req = urllib.request.Request(url, data=urllib.parse.urlencode({"data": q}).encode(), headers={"User-Agent": "SicuritaliaOpportunityHub/3.4"})
            with urllib.request.urlopen(req, timeout=min(timeout, resto)) as r:
                data = json.loads(r.read().decode("utf-8"))
            if isinstance(data, dict) and data.get("remark") and not data.get("elements") and "timed out" in str(data.get("remark")).lower():
                lenti += 1
                continue
            if len(_cache) > 60:
                _cache.clear()
            _cache[q] = (time.time(), data)
            return data
        except urllib.error.HTTPError as ex:
            if ex.code in (429, 502, 503, 504):
                lenti += 1
            else:
                altri.append("errore %d" % ex.code)
        except (socket.timeout, TimeoutError):
            lenti += 1
        except Exception as ex:
            if "timed out" in str(ex).lower():
                lenti += 1
            else:
                altri.append(str(ex)[:80])
    if lenti:
        raise RuntimeError("Il servizio gratuito di ricerca (OpenStreetMap) è lento o sovraccarico in questo momento. "
                           "Riprova tra qualche minuto; se cerchi in una zona grande, prova un comune alla volta.")
    raise RuntimeError("Non riesco a collegarmi al servizio di ricerca: il server potrebbe non avere accesso a Internet "
                       "o un firewall blocca overpass-api.de. Dettaglio: %s" % (altri[-1] if altri else "nessuna risposta"))
