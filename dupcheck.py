"""Verifica duplicati: stesso cliente presente in più liste (customer base, campagne, proposte).

Livelli di corrispondenza:
  certo     stessa partita IVA / codice fiscale, oppure stesso telefono
  probabile stesso nome (ignorando punteggiatura, maiuscole, srl/spa/snc ecc. e ordine delle parole)
  simile    nome molto simile (refusi, parole in più)
"""
import difflib
import re
import unicodedata

FORME = {"srl", "srls", "spa", "snc", "sas", "sapa", "sdf", "scarl", "scrl", "soc", "societa", "coop", "cooperativa", "ditta",
         "ind", "individuale", "di", "e", "del", "della", "dei", "il", "la", "lo", "the", "and", "unipersonale", "s", "r", "l", "p", "a", "c", "n"}
LIVELLI = {"certo": 4, "probabile": 3, "indirizzo": 2, "simile": 1}
ETICHETTA = {"certo": "Già presente", "probabile": "Probabilmente già presente", "indirizzo": "Stesso indirizzo", "simile": "Nome simile"}
MAX_CAND = 300


def _ascii(s):
    return unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()


def tokens(nome):
    t = re.sub(r"[^a-z0-9 ]+", " ", _ascii(nome).replace("&", " e ").replace(".", ""))
    out = [w for w in t.split() if w not in FORME]
    return out or [w for w in t.split()]


def norm_nome(nome):
    return " ".join(sorted(tokens(nome)))


def norm_ident(s):
    """Partita IVA (11 cifre) o codice fiscale (16 caratteri); altrimenti vuoto."""
    v = re.sub(r"[^A-Za-z0-9]", "", str(s or "")).upper()
    if v.startswith("IT") and len(v) == 13 and v[2:].isdigit():
        v = v[2:]
    if v.isdigit() and len(v) == 10:      # Excel toglie lo zero iniziale
        v = v.zfill(11)
    if len(v) == 11 and v.isdigit():
        return v
    if len(v) == 16 and re.match(r"^[A-Z]{6}\d{2}[A-Z]\d{2}[A-Z]\d{3}[A-Z]$", v):
        return v
    return ""


def piva_da_testo(testo):
    m = re.search(r"(?:p\.?\s*iva|partita\s+iva|c\.?\s*f\.?)\D{0,6}(\d{11})\b", str(testo or ""), re.I)
    return m.group(1) if m else ""


ABBR = {"v": "via", "vle": "viale", "v.le": "viale", "p.zza": "piazza", "pza": "piazza", "pzza": "piazza", "p": "piazza", "c.so": "corso", "cso": "corso",
        "str": "strada", "loc": "localita", "lgo": "largo", "n": "", "nr": "", "civ": "", "civico": ""}


def norm_ind(s):
    """Indirizzo normalizzato (via + civico). Vuoto se manca il numero civico."""
    t = _ascii(s).replace("p.zza", "piazza").replace("c.so", "corso").replace("v.le", "viale")
    t = re.sub(r"[^a-z0-9 ]+", " ", t)
    w = [ABBR.get(x, x) for x in t.split()]
    w = [x for x in w if x]
    t = " ".join(w)
    if len(t) < 6 or not re.search(r"\d", t) or not re.search(r"[a-z]{3}", t):
        return ""
    return t


def norm_tels(s):
    out = set()
    for p in re.split(r"[/;,|]|\be\b|\bo\b", str(s or "")):
        d = re.sub(r"\D", "", p)
        if d.startswith("0039"):
            d = d[4:]
        elif d.startswith("39") and len(d) >= 11:
            d = d[2:]
        if len(d) >= 6:
            out.add(d)
    return out


class Indice:
    def __init__(self):
        self.voci = []
        self.by_id, self.by_tel, self.by_nome, self.by_tok, self.by_ind = {}, {}, {}, {}, {}

    def add(self, fonte, nome, ident="", tel="", info="", meta=None, ind="", loc=""):
        v = {"fonte": fonte, "nome": nome, "nn": norm_nome(nome), "ident": norm_ident(ident), "tels": norm_tels(tel),
             "ind": norm_ind(ind), "loc": norm_nome(loc), "info": info, "meta": meta or {},
             "tk": frozenset(w for w in tokens(nome) if len(w) >= 3)}
        i = len(self.voci)
        self.voci.append(v)
        if v["ident"]:
            self.by_id.setdefault(v["ident"], []).append(i)
        for t in v["tels"]:
            self.by_tel.setdefault(t, []).append(i)
        if v["ind"]:
            self.by_ind.setdefault(v["ind"], []).append(i)
        if v["nn"]:
            self.by_nome.setdefault(v["nn"], []).append(i)
            for w in set(tokens(nome)):
                if len(w) >= 3:
                    self.by_tok.setdefault(w, []).append(i)

    def cerca(self, nome, ident="", tel="", salta=None, max_ris=5, ind="", loc=""):
        """Restituisce le corrispondenze migliori: [{livello, motivo, fonte, info, nome, meta}]."""
        trovati = {}

        def put(i, liv, motivo):
            v = self.voci[i]
            if salta and salta(v):
                return
            if i not in trovati or LIVELLI[liv] > LIVELLI[trovati[i][0]]:
                trovati[i] = (liv, motivo)
        idn = norm_ident(ident)
        if idn:
            for i in self.by_id.get(idn, []):
                put(i, "certo", "stessa partita IVA / codice fiscale")
        for t in norm_tels(tel):
            for i in self.by_tel.get(t, []):
                put(i, "certo", "stesso telefono")
        ni = norm_ind(ind)
        if ni:
            nl = norm_nome(loc)
            for i in self.by_ind.get(ni, []):
                v = self.voci[i]
                if nl and v["loc"] and nl != v["loc"]:
                    continue
                put(i, "indirizzo", "stesso indirizzo")
        nn = norm_nome(nome)
        if nn:
            for i in self.by_nome.get(nn, []):
                put(i, "probabile", "stesso nome")
            qt = frozenset(w for w in tokens(nome) if len(w) >= 3)
            cand, cont = set(), set()
            for w in qt:
                lst = self.by_tok.get(w, [])
                if len(w) >= 4 and len(lst) <= MAX_CAND:
                    cand.update(lst)
                cont.update(lst)
            # nome contenuto nell'altro (es. «Rossi Mobili» / «Mobilificio Rossi Mobili di Rossi Mario»)
            for i in cont:
                if i in trovati:
                    continue
                vt = self.voci[i]["tk"]
                if not qt or not vt:
                    continue
                piccolo, grande = (qt, vt) if len(qt) <= len(vt) else (vt, qt)
                if not piccolo <= grande:
                    continue
                if len(piccolo) >= 2 or (len(piccolo) == 1 and len(next(iter(piccolo))) >= 6 and len(grande) <= 3):
                    put(i, "simile", "nome contenuto: «%s»" % self.voci[i]["nome"])
            for i in cand:
                if i in trovati:
                    continue
                v = self.voci[i]
                if abs(len(v["nn"]) - len(nn)) > 8:
                    continue
                if difflib.SequenceMatcher(None, nn, v["nn"]).ratio() >= 0.9:
                    put(i, "simile", "nome molto simile: «%s»" % v["nome"])
        ris = sorted(trovati.items(), key=lambda kv: (-LIVELLI[kv[1][0]], not self.voci[kv[0]]["meta"].get("lavorazione"), kv[0]))[:max_ris]
        return [{"livello": liv, "motivo": mot, "fonte": self.voci[i]["fonte"], "info": self.voci[i]["info"],
                 "nome": self.voci[i]["nome"], "meta": self.voci[i]["meta"]} for i, (liv, mot) in ris]


PREFISSO = {"certo": "", "probabile": "Probabile: ", "indirizzo": "Stesso indirizzo: ", "simile": "Nome simile: "}


def riassunto(matches):
    """Testo breve per badge e note."""
    return " | ".join(PREFISSO[m["livello"]] + m["info"] for m in matches)


def migliore(matches):
    return max((m["livello"] for m in matches), key=lambda l: LIVELLI[l], default="")
