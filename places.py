"""Ricerca di attività con Google Places API (New) - Text Search. Solo libreria standard.
Costi: i campi telefono e sito web rientrano nel livello «Enterprise» (circa 35 USD ogni 1.000 richieste, 1.000 gratuite al mese)."""
import json
import math
import os
import socket
import time
import urllib.error
import urllib.request
from datetime import date

URL = os.environ.get("SOH_GOOGLE_URL", "https://places.googleapis.com/v1/places:searchText")
MAX_GIORNO = int(os.environ.get("SOH_GOOGLE_MAX_GIORNO", "300"))
MAX_RICHIESTE_RICERCA = 12
CAMPI = ("places.id,places.displayName,places.formattedAddress,places.nationalPhoneNumber,places.internationalPhoneNumber,"
         "places.websiteUri,places.location,places.businessStatus,nextPageToken")


class PlacesError(Exception):
    pass


_uso = {"giorno": None, "n": 0}
_cache = {}


def usate_oggi():
    return _uso["n"] if _uso["giorno"] == date.today() else 0


def _conta():
    oggi = date.today()
    if _uso["giorno"] != oggi:
        _uso.update(giorno=oggi, n=0)
    if _uso["n"] >= MAX_GIORNO:
        raise PlacesError("Raggiunto il limite giornaliero di ricerche Google (%d) impostato per controllare i costi. Riprova domani." % MAX_GIORNO)
    _uso["n"] += 1


def _errore_http(e):
    try:
        msg = (json.loads(e.read().decode("utf-8", "replace")).get("error") or {}).get("message", "")
    except Exception:
        msg = ""
    if e.code in (400, 403) and ("API key" in msg or "API_KEY" in msg or "not valid" in msg):
        return "La chiave Google non è valida."
    if e.code == 403:
        return ("Google ha rifiutato la richiesta: controlla che «Places API (New)» sia abilitata, che la fatturazione sia attiva "
                "e che la chiave non sia limitata a un altro indirizzo. Dettaglio: " + msg[:160])
    if e.code == 429:
        return "Google segnala troppe richieste: riprova tra poco."
    if e.code >= 500:
        return "Il servizio Google è momentaneamente non disponibile."
    return "Errore Google (%d): %s" % (e.code, msg[:160])


def _post(chiave, corpo, campi):
    req = urllib.request.Request(URL, data=json.dumps(corpo).encode("utf-8"), method="POST", headers={
        "Content-Type": "application/json", "X-Goog-Api-Key": chiave, "X-Goog-FieldMask": campi,
        "User-Agent": "SicuritaliaOpportunityHub/3.13"})
    _conta()
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return json.loads(r.read().decode("utf-8", "replace"))
    except urllib.error.HTTPError as e:
        raise PlacesError(_errore_http(e))
    except (socket.timeout, TimeoutError):
        raise PlacesError("Google non ha risposto in tempo.")
    except (urllib.error.URLError, OSError) as e:
        raise PlacesError("Non riesco a collegarmi a Google (serve Internet e l'accesso a places.googleapis.com). Dettaglio: %s" % e)
    except ValueError:
        raise PlacesError("Risposta di Google non valida.")


def prova(chiave):
    """Richiesta minima (solo ID: livello gratuito) per verificare la chiave."""
    _post(chiave, {"textQuery": "ristorante Verona", "pageSize": 1, "languageCode": "it"}, "places.id")


def _km(la1, lo1, la2, lo2):
    p = math.pi / 180
    a = 0.5 - math.cos((la2 - la1) * p) / 2 + math.cos(la1 * p) * math.cos(la2 * p) * (1 - math.cos((lo2 - lo1) * p)) / 2
    return 12742 * math.asin(math.sqrt(a))


def cerca(chiave, testo, lat=None, lon=None, raggio_km=None, pagine=3, budget=None):
    """Restituisce (righe, troncato). Ogni riga ha le stesse chiavi dei risultati OpenStreetMap."""
    k = (testo.lower(), round(lat or 0, 3), round(lon or 0, 3), raggio_km, pagine)
    now = time.time()
    if k in _cache and now - _cache[k][0] < 1800:
        return _cache[k][1], _cache[k][2]
    corpo = {"textQuery": testo, "languageCode": "it", "regionCode": "IT", "pageSize": 20}
    if lat is not None and lon is not None:
        corpo["locationBias"] = {"circle": {"center": {"latitude": lat, "longitude": lon}, "radius": float(min(50, max(1, raggio_km or 5)) * 1000)}}
    righe, tok, troncato = [], None, False
    for i in range(pagine):
        if budget is not None:
            if budget[0] <= 0:
                troncato = True
                break
            budget[0] -= 1
        if tok:
            corpo["pageToken"] = tok
        d = _post(chiave, corpo, CAMPI)
        for p in d.get("places") or []:
            if p.get("businessStatus") == "CLOSED_PERMANENTLY":
                continue
            nome = ((p.get("displayName") or {}).get("text") or "").strip()
            if not nome:
                continue
            ind = (p.get("formattedAddress") or "").replace(", Italia", "").strip()
            loc = p.get("location") or {}
            la, lo = loc.get("latitude"), loc.get("longitude")
            righe.append({"nome": nome, "indirizzo": ind, "via": ind.split(",")[0], "citta": "", "cap": "",
                          "telefono": (p.get("nationalPhoneNumber") or p.get("internationalPhoneNumber") or "").strip(),
                          "sito": p.get("websiteUri") or "", "email": "", "piva": "", "osm": "g/" + str(p.get("id")), "lat": la, "lon": lo,
                          "km": round(_km(lat, lon, la, lo), 1) if None not in (lat, lon, la, lo) else None})
        tok = d.get("nextPageToken")
        if not tok:
            break
        if i == pagine - 1:
            troncato = True
    _cache[k] = (now, righe, troncato)
    if len(_cache) > 60:
        _cache.clear()
    return righe, troncato
