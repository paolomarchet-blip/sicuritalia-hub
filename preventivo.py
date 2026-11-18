# -*- coding: utf-8 -*-
"""Preventivo Sicuritalia: dati dei servizi quotati e generazione del PDF nel formato dell'«Offerta tipo».
Le pagine fisse (copertina, presentazione, referenze, certificazioni, condizioni generali) sono in assets/offerta_base.pdf;
qui si compilano l'intestazione («Spett.le», data) e le tabelle dell'offerta economica, poi si rinumerano le pagine."""
import base64
import io
import json
import os
import re
from datetime import datetime

from pypdf import PageObject, PdfReader, PdfWriter

import pdfmini as pm

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "offerta_base.pdf")
LUOGO = "Treviso"
NAVY = (0, 32, 96)
GRIGIO = (217, 217, 217)
BORDO = (190, 190, 190)

SERVIZI = {
    "ispettivo": "Ispettivo",
    "allarme": "Collegamento di allarme (per ogni periferica) e Pronto Intervento",
    "video": "Videosorveglianza",
    "videoisp": "Videoispezioni",
    "incendio": "Collegamento incendio (per ogni periferica) e Pronto Intervento",
}
TARIFFA_EXTRA = "€ 20,00"


def euro(v):
    """Importo scritto all'italiana (1.250,50), all'inglese (1250.50) o come numero; None se non valido."""
    if isinstance(v, (int, float)):
        x = float(v)
    else:
        t = str(v or "").replace("€", "").replace(" ", "").strip()
        if not t:
            return None
        if "," in t:
            t = t.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"\d{1,3}(\.\d{3})+", t):
            t = t.replace(".", "")
        try:
            x = float(t)
        except ValueError:
            return None
    return round(x, 2) if 0 <= x < 1e8 else None


def fmt_euro(x):
    s = "{:,.2f}".format(float(x)).replace(",", "X").replace(".", ",").replace("X", ".")
    return "€ " + s


def _int(v, minimo=0, massimo=100000):
    try:
        n = int(str(v if v is not None else "").strip() or 0)
    except ValueError:
        return None
    return n if minimo <= n <= massimo else None


def valida_riga(r):
    """Restituisce (riga pulita, None) oppure (None, messaggio di errore)."""
    s = (r or {}).get("servizio")
    if s not in SERVIZI:
        return None, "Scegli il servizio da quotare"
    nome = SERVIZI[s]
    canone = euro(r.get("canone"))
    if not canone:
        return None, "%s: indica il canone mensile (importo in euro maggiore di zero)" % nome
    out = {"servizio": s, "canone": canone}
    if s == "ispettivo":
        for k in ("int_nott", "int_diur", "est_nott", "est_diur"):
            n = _int(r.get(k))
            if n is None:
                return None, "%s: il numero di ispezioni non è valido" % nome
            out[k] = n
        if not any(out[k] for k in ("int_nott", "int_diur", "est_nott", "est_diur")):
            return None, "%s: indica quante ispezioni (interne/esterne, notturne/diurne)" % nome
    elif s in ("allarme", "incendio"):
        if s == "allarme":
            if r.get("dotazione") not in ("comodato", "periferica"):
                return None, "%s: scegli se è in comodato l'allarme o solo la periferica di collegamento" % nome
            out["dotazione"] = r["dotazione"]
        else:
            out["periferica"] = bool(r.get("periferica") in (True, "si", "sì", 1, "1", "true"))
        n = _int(r.get("franchigia"))
        if n is None:
            return None, "%s: indica gli interventi in franchigia (anche 0)" % nome
        out["franchigia"] = n
        ut = r.get("una_tantum")
        if str(ut or "").strip():
            x = euro(ut)
            if x is None:
                return None, "%s: il costo una tantum di installazione non è valido" % nome
            out["una_tantum"] = x
    elif s == "video":
        n = _int(r.get("telecamere"), 1)
        if n is None:
            return None, "%s: indica il numero di telecamere da collegare" % nome
        out["telecamere"] = n
    elif s == "videoisp":
        n = _int(r.get("n"), 1)
        if n is None:
            return None, "%s: indica il numero di videoispezioni" % nome
        out["n"] = n
    return out, None


def punti(r):
    """Righe «modalità operative» di un servizio, come nell'offerta tipo."""
    s = r["servizio"]
    if s == "ispettivo":
        out = []
        for k, (sing, plur) in (("int_nott", ("ispezione interna NOTTURNA", "ispezioni interne NOTTURNE")),
                                ("int_diur", ("ispezione interna DIURNA", "ispezioni interne DIURNE")),
                                ("est_nott", ("ispezione esterna NOTTURNA", "ispezioni esterne NOTTURNE")),
                                ("est_diur", ("ispezione esterna DIURNA", "ispezioni esterne DIURNE"))):
            n = r.get(k) or 0
            if n:
                out.append("%d %s." % (n, sing if n == 1 else plur))
        return out
    if s in ("allarme", "incendio"):
        out = []
        if s == "allarme":
            out.append("Il servizio prevede l'installazione di %s." % (
                "un impianto di allarme in comodato d'uso, completo di periferica di collegamento" if r.get("dotazione") == "comodato"
                else "una sola periferica di collegamento (allarme del Cliente)"))
        else:
            out.append("Il servizio prevede %s." % (
                "l'installazione di una periferica di collegamento" if r.get("periferica")
                else "il collegamento senza installazione di periferica (periferica non prevista)"))
        n = r.get("franchigia") or 0
        out.append("Il canone comprende n. %d interventi su allarme in franchigia." % n if n
                   else "Il canone non comprende interventi su allarme in franchigia.")
        out.append("Ogni singolo intervento successivo a quelli previsti in franchigia verrà tariffato a %s + IVA cad." % TARIFFA_EXTRA
                   if n else "Ogni singolo intervento verrà tariffato a %s + IVA cad." % TARIFFA_EXTRA)
        if r.get("una_tantum"):
            out.append("Per l'installazione di quanto sopra sarà previsto un contributo installativo una tantum pari a %s." % fmt_euro(r["una_tantum"]))
        return out
    if s == "video":
        n = r["telecamere"]
        return ["Il servizio prevede il collegamento di n. %d %s." % (n, "telecamera" if n == 1 else "telecamere")]
    if s == "videoisp":
        n = r["n"]
        return ["Il servizio prevede n. %d %s sulle telecamere collegate." % (n, "videoispezione" if n == 1 else "videoispezioni")]
    return []


# ---------------------------------------------------------------- pagine dell'offerta economica
X0, X1, C1, C2 = 57.0, 538.7, 149.5, 408.0
F_BODY, LH = 8.7, 11.6


def _pagine_economiche(righe, note=""):
    pagine = []
    pg = None
    y = 0

    def nuova(prima):
        p = pm.Pagina()
        if not prima:
            p.rett(54, 68, 492, 60, fill=(255, 255, 255))      # copre il titolo di sezione: qui è solo la continuazione
        return p, (146.0 if prima else 92.0)

    pg, y = nuova(True)
    for r in righe:
        pts = []
        for t in punti(r):
            pts.append(pm.a_capo(t, "F1", F_BODY, C2 - C1 - 36))
        h2 = 22 + sum(len(x) for x in pts) * LH + 4 * max(0, len(pts) - 1)
        c1 = pm.a_capo(SERVIZI[r["servizio"]], "F2", 8.2, C1 - X0 - 18)
        h = max(h2, 20 + len(c1) * 10.2, 48)
        if y + 35 + h > 780 and y > 150:
            pagine.append(pg)
            pg, y = nuova(False)
        # intestazione
        pg.rett(X0, y, X1 - X0, 35, fill=GRIGIO)
        pg.linea(C1, y, C1, y + 35, BORDO)
        pg.linea(C2, y, C2, y + 35, BORDO)
        pg.testo((X0 + C1) / 2, y + 15, "TIPOLOGIA", "F2", 9.6, NAVY, "c")
        pg.testo((X0 + C1) / 2, y + 27, "SERVIZIO", "F2", 9.6, NAVY, "c")
        pg.testo((C1 + C2) / 2, y + 22, "MODALITÀ OPERATIVE", "F2", 9.6, NAVY, "c")
        pg.testo((C2 + X1) / 2, y + 22, "CANONE MESE", "F2", 9.6, NAVY, "c")
        y += 35
        pg.rett(X0, y, X1 - X0, h, stroke=BORDO)
        pg.linea(C1, y, C1, y + h, BORDO)
        pg.linea(C2, y, C2, y + h, BORDO)
        # nome servizio
        ty = y + (h - len(c1) * 10.2) / 2 + 8
        for ln in c1:
            pg.testo(X0 + 8, ty, ln, "F2", 8.2, NAVY)
            ty += 10.2
        # modalità
        blocco = sum(len(x) for x in pts) * LH + 4 * max(0, len(pts) - 1)
        ty = y + (h - blocco) / 2.0 + 8.3
        for bl in pts:
            pg.rett(C1 + 12, ty - 5.6, 3.6, 3.6, fill=(0, 0, 0))
            for ln in bl:
                pg.testo(C1 + 24, ty, ln, "F1", F_BODY)
                ty += LH
            ty += 4
        # canone
        pg.testo((C2 + X1) / 2, y + h / 2 + 3, "%s + IVA" % fmt_euro(r["canone"]), "F1", 9.4, (0, 0, 0), "c")
        y += h + 16
    note = (note or "").strip()
    if note:
        righe_n = []
        for par in note.replace("\r", "").split("\n"):
            righe_n += pm.a_capo(par, "F1", F_BODY, X1 - X0 - 16) if par.strip() else [""]
        h = 24 + len(righe_n) * LH
        if y + h > 780 and y > 150:
            pagine.append(pg)
            pg, y = nuova(False)
        pg.rett(X0, y, X1 - X0, 18, fill=GRIGIO)
        pg.testo(X0 + 8, y + 12.5, "NOTE", "F2", 9.6, NAVY)
        pg.rett(X0, y + 18, X1 - X0, h - 18, stroke=BORDO)
        ty = y + 18 + 14
        for ln in righe_n:
            if ln:
                pg.testo(X0 + 8, ty, ln, "F1", F_BODY)
            ty += LH
        y += h + 10
    pagine.append(pg)
    return pagine


# ------------------------------------------------------------------ listino materiali (allegato «Elenco materiale»)
LISTINO_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "listino.json")
_LISTINO = None
IMG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "listino_img")


def listino():
    """[{c: codice, n: descrizione, k: [{q, c, n}] (solo i kit)}], dal file assets/listino.json."""
    global _LISTINO
    if not _LISTINO:                      # se manca il file non si memorizza il vuoto: ricarica appena viene copiato
        try:
            _LISTINO = json.load(open(LISTINO_FILE, encoding="utf-8"))
        except (OSError, ValueError):
            _LISTINO = []
    return _LISTINO


def _jpeg_size(b):
    """(larghezza, altezza) di un JPEG, leggendo il marcatore SOF (senza librerie esterne)."""
    i = 2
    while i + 9 < len(b):
        if b[i] != 0xFF:
            i += 1
            continue
        mk = b[i + 1]
        if mk in (0xC0, 0xC1, 0xC2):
            return (b[i + 7] << 8) | b[i + 8], (b[i + 5] << 8) | b[i + 6]
        if mk in (0xD8, 0x01) or 0xD0 <= mk <= 0xD7:
            i += 2
        else:
            i += 2 + ((b[i + 2] << 8) | b[i + 3])
    raise ValueError("JPEG non valido")


def valida_materiale(m):
    """m: lista di {chiave, qta}. Tiene solo le quantità > 0 e copia nome/composizione (così il preventivo non cambia se il listino cambia).
    Ritorna (lista, errore)."""
    if m in (None, ""):
        return [], None
    if not isinstance(m, list):
        return [], "Elenco materiale non valido"
    per_chiave = {(x["c"] or x["n"]): x for x in listino()}
    out, visti = [], set()
    for r in m:
        if not isinstance(r, dict):
            return [], "Elenco materiale non valido"
        ch = str(r.get("chiave") or "")
        a = per_chiave.get(ch)
        if a is None:
            return [], "Articolo non presente nel listino"
        try:
            q = int(float(str(r.get("qta") or 0).replace(",", ".")))
        except ValueError:
            return [], "Quantità non valida per «%s»" % a["n"]
        if q < 0 or q > 9999:
            return [], "Quantità non valida per «%s»" % a["n"]
        if q == 0 or ch in visti:
            continue
        visti.add(ch)
        e = {"chiave": ch, "c": a["c"], "n": a["n"], "qta": q}
        if a.get("k"):
            e["k"] = a["k"]
        if a.get("i") is not None:
            try:
                e["img"] = base64.b64encode(open(os.path.join(IMG_DIR, "%d.jpg" % a["i"]), "rb").read()).decode()
            except OSError:
                pass
        out.append(e)
    if len(out) > 150:
        return [], "Troppi articoli nell'elenco materiale (massimo 150)"
    return out, None


def _pagine_materiale(materiale, cliente, data_iso):
    """Pagine dell'allegato «Elenco materiale»: codice, descrizione e quantità (per i kit anche la composizione)."""
    pagine = []
    CW, FW, QW = 74.0, 46.0, 52.0             # larghezze colonne codice, foto e quantità
    XF = X0 + CW
    XC, XQ = XF + FW, X1 - QW

    def nuova(prima):
        p = pm.Pagina()
        p.rett(54, 68, 492, 60, fill=(255, 255, 255))
        p.rett(X0, 74, X1 - X0, 34, fill=NAVY)
        p.testo((X0 + X1) / 2, 88, "ALLEGATO - ELENCO MATERIALE DELL'IMPIANTO", "F2", 10.2, (255, 255, 255), "c")
        p.testo((X0 + X1) / 2, 100, "(l'elenco comprende solo i materiali con quantità maggiore di zero)", "F1", 7.4, (255, 255, 255), "c")
        y = 122.0
        if prima:
            p.testo(X0, y + 8, "Cliente: %s" % (cliente or ""), "F2", 9)
            p.testo(X1, y + 8, "Data: %s" % _data_it(data_iso), "F1", 9, (0, 0, 0), "r")
            y += 22
        return p, y

    def intest(p, y):
        p.rett(X0, y, X1 - X0, 20, fill=GRIGIO)
        p.linea(XF, y, XF, y + 20, BORDO)
        p.linea(XC, y, XC, y + 20, BORDO)
        p.linea(XQ, y, XQ, y + 20, BORDO)
        p.testo(X0 + CW / 2, y + 13.5, "CODICE", "F2", 9, NAVY, "c")
        p.testo((XF + XC) / 2, y + 13.5, "FOTO", "F2", 9, NAVY, "c")
        p.testo((XC + XQ) / 2, y + 13.5, "DESCRIZIONE", "F2", 9, NAVY, "c")
        p.testo((XQ + X1) / 2, y + 13.5, "Q.TÀ", "F2", 9, NAVY, "c")
        return y + 20

    pg, y = nuova(True)
    y = intest(pg, y)
    for a in materiale:
        nome = pm.a_capo(a["n"], "F2", 8.4, XQ - XC - 14)
        sub = []
        for k in a.get("k") or []:
            t = "%d x %s%s" % (k["q"] * a["qta"], (k["c"] + "  ") if k.get("c") else "", k["n"])
            sub += pm.a_capo(t, "F1", 7.4, XQ - XC - 30)
        h = max(9 + len(nome) * 10.4 + (len(sub) * 9.2 + 3 if sub else 0), 44 if a.get("img") else 0)
        if y + h > 790:
            pagine.append(pg)
            pg, y = nuova(False)
            y = intest(pg, y)
        pg.rett(X0, y, X1 - X0, h, stroke=BORDO)
        pg.linea(XF, y, XF, y + h, BORDO)
        pg.linea(XC, y, XC, y + h, BORDO)
        pg.linea(XQ, y, XQ, y + h, BORDO)
        pg.testo(X0 + CW / 2, y + 13, a["c"] or "-", "F1", 8.2, (0, 0, 0), "c")
        if a.get("img"):
            try:
                jp = base64.b64decode(a["img"])
                pw, ph = _jpeg_size(jp)
                sc = min((FW - 8) / pw, (h - 8) / ph, 2.0)
                iw, ih = pw * sc, ph * sc
                pg.img(XF + (FW - iw) / 2, y + (h - ih) / 2, iw, ih, jp, pw, ph)
            except Exception:
                pass
        ty = y + 13
        for ln in nome:
            pg.testo(XC + 7, ty, ln, "F2", 8.4, NAVY)
            ty += 10.4
        ty -= 1
        for ln in sub:
            pg.testo(XC + 18, ty + 3, ln, "F1", 7.4, (70, 70, 70))
            ty += 9.2
        pg.testo((XQ + X1) / 2, y + 13, str(a["qta"]), "F2", 10, (0, 0, 0), "c")
        y += h
    pagine.append(pg)
    return pagine


def _num(pg, n):
    pg.rett(530, 811, 12, 13, fill=(255, 255, 255))
    pg.testo(537.5, 820.3, str(n), "F1", 7, NAVY, "r")


def _data_it(iso):
    try:
        return datetime.strptime(str(iso)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return str(iso or "")


def genera_pdf(p):
    """p: {cliente, indirizzo, data_richiesta (AAAA-MM-GG), righe:[...]} -> bytes PDF."""
    base = open(BASE, "rb").read()

    unico = PdfReader(io.BytesIO(base))

    def pagina(i, ov=None):
        # tutte le pagine condividono lo stesso lettore: le immagini restano un solo oggetto anche se la pagina 7 si ripete
        if i == 7:       # la pagina-modello si ripete: ogni copia è una pagina nuova che ne importa il contenuto (risorse condivise)
            src = unico.pages[7]
            pg = PageObject.create_blank_page(w, src.mediabox.width, src.mediabox.height)
            pg.merge_page(src)
            w.add_page(pg)
            pg = w.pages[-1]
        else:
            pg = w.add_page(unico.pages[i])
        if ov is not None:
            pg.merge_page(PdfReader(io.BytesIO(ov.pdf())).pages[0])
        return pg
    eco = _pagine_economiche(p["righe"], p.get("note"))
    k = len(eco)
    w = PdfWriter()
    pagina(0)
    # lettera: Spett.le + data
    ov = pm.Pagina()
    ov.rett(272, 66, 52, 20, fill=(255, 255, 255))          # «INDICE» (qui fuori posto: va sulla pagina dell'indice)
    ov.rett(52, 252, 90, 20, fill=(255, 255, 255))          # «Trevio li»
    y = 226
    nome = pm.a_capo((p.get("cliente") or "").strip(), "F2", 10, 235)
    for ln in nome[:2]:
        ov.testo(305, y, ln, "F2", 10)
        y += 12
    ind = pm.a_capo((p.get("indirizzo") or "").strip(), "F1", 9.5, 235)
    for ln in ind[: 3 - min(2, len(nome))]:
        ov.testo(305, y, ln, "F1", 9.5)
        y += 11
    ov.testo(56.8, 266, "%s, li %s" % (LUOGO, _data_it(p.get("data_richiesta"))), "F1", 10)
    pagina(1, ov)
    # indice
    ov = pm.Pagina()
    ov.testo(297.7, 82, "INDICE", "F2", 10, NAVY, "c")
    ov.rett(531, 217, 17, 16, fill=(255, 255, 255))
    ov.testo(544.4, 229.3, "7", "F1", 10, NAVY, "r")
    ov.rett(536, 231, 12, 16, fill=(255, 255, 255))
    ov.testo(544.4, 242.6, str(7 + k), "F1", 10, NAVY, "r")
    pagina(2, ov)
    for i in (3, 4, 5, 6):
        pagina(i)
    for j, e in enumerate(eco):
        _num(e, 7 + j)
        pagina(7, e)
    for j, i in enumerate((8, 9)):
        ov = pm.Pagina()
        _num(ov, 7 + k + j)
        pagina(i, ov)
    mat = p.get("materiale") or []
    if mat:
        for j, e in enumerate(_pagine_materiale(mat, p.get("cliente"), p.get("data_richiesta"))):
            _num(e, 9 + k + j)
            pagina(7, e)
    w.add_metadata({"/Title": "Proposta tecnico-economica - %s" % (p.get("cliente") or ""), "/Author": "Sicuritalia"})
    out = io.BytesIO()
    w.write(out)
    return out.getvalue()
