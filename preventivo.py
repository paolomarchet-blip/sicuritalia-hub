# -*- coding: utf-8 -*-
"""Preventivo Sicuritalia: dati dei servizi quotati e generazione del PDF nel formato dell'«Offerta tipo».
Le pagine fisse (copertina, presentazione, referenze, certificazioni, condizioni generali) sono in assets/offerta_base.pdf;
qui si compilano l'intestazione («Spett.le», data) e le tabelle dell'offerta economica, poi si rinumerano le pagine."""
import io
import os
import re
from datetime import datetime

from pypdf import PdfReader, PdfWriter

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


def _pagine_economiche(righe):
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
        # le pagine usate una volta sola condividono il lettore (le immagini restano un solo oggetto); la pagina 7 si ripete
        pg = (PdfReader(io.BytesIO(base)) if i == 7 else unico).pages[i]
        if ov is not None:
            pg.merge_page(PdfReader(io.BytesIO(ov.pdf())).pages[0])
        return pg
    eco = _pagine_economiche(p["righe"])
    k = len(eco)
    w = PdfWriter()
    w.add_page(pagina(0))
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
    w.add_page(pagina(1, ov))
    # indice
    ov = pm.Pagina()
    ov.testo(297.7, 82, "INDICE", "F2", 10, NAVY, "c")
    ov.rett(531, 217, 17, 16, fill=(255, 255, 255))
    ov.testo(544.4, 229.3, "7", "F1", 10, NAVY, "r")
    ov.rett(536, 231, 12, 16, fill=(255, 255, 255))
    ov.testo(544.4, 242.6, str(7 + k), "F1", 10, NAVY, "r")
    w.add_page(pagina(2, ov))
    for i in (3, 4, 5, 6):
        w.add_page(pagina(i))
    for j, e in enumerate(eco):
        _num(e, 7 + j)
        w.add_page(pagina(7, e))
    for j, i in enumerate((8, 9)):
        ov = pm.Pagina()
        _num(ov, 7 + k + j)
        w.add_page(pagina(i, ov))
    w.add_metadata({"/Title": "Proposta tecnico-economica - %s" % (p.get("cliente") or ""), "/Author": "Sicuritalia"})
    out = io.BytesIO()
    w.write(out)
    return out.getvalue()
