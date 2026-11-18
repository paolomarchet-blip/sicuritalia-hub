"""Lettura delle mail di Caring (LEAD) e Ufficio Disdette: estrae i dati per creare un'opportunità.

Formati accettati: PDF (mail stampata da Outlook), .eml, .txt/.html oppure testo incollato.
"""
import email
import html as _html
import io
import re
from email import policy

MAX_CHARS = 60000


def _clean(t):
    t = t.replace("\xa0", " ").replace("​", "").replace("\r", "")
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n[ \t]+", "\n", t)
    return t.strip()


def _html_to_text(h):
    h = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", h)
    h = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>|</li>|</h\d>", "\n", h)
    h = re.sub(r"(?s)<[^>]+>", " ", h)
    return _html.unescape(h)


def _decode(b):
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    return b.decode("utf-8", errors="replace")


def extract_text(filename, data):
    name = (filename or "").lower()
    if name.endswith(".msg"):
        raise ValueError("I file .msg non sono leggibili: salva la mail come PDF (Stampa > Microsoft Print to PDF) "
                         "oppure copia e incolla il testo della mail.")
    if name.endswith(".pdf") or data[:5] == b"%PDF-":
        try:
            from pypdf import PdfReader
            r = PdfReader(io.BytesIO(data))
            if r.is_encrypted:
                raise ValueError("Il PDF è protetto da password")
            txt = "\n".join((p.extract_text() or "") for p in r.pages[:12])
        except ValueError:
            raise
        except Exception:
            raise ValueError("PDF non leggibile")
        if len(txt.strip()) < 40:
            raise ValueError("Il PDF non contiene testo selezionabile (è una scansione?). Incolla il testo della mail.")
        return _clean(txt)[:MAX_CHARS]
    if name.endswith(".eml") or data[:200].lower().find(b"\nsubject:") >= 0 or data[:20].lower().startswith(b"received:"):
        msg = email.message_from_bytes(data, policy=policy.default)
        body = msg.get_body(preferencelist=("plain", "html"))
        txt = ""
        if body is not None:
            c = body.get_content()
            txt = _html_to_text(c) if body.get_content_type() == "text/html" else c
        head = "Da: %s\nInviato: %s\nOggetto: %s\n" % (msg.get("From", ""), msg.get("Date", ""), msg.get("Subject", ""))
        return _clean(head + txt)[:MAX_CHARS]
    txt = _decode(data)
    if name.endswith((".htm", ".html")) or re.search(r"(?i)<html|<body|<div", txt[:2000]):
        txt = _html_to_text(txt)
    return _clean(txt)[:MAX_CHARS]


# ----------------------------------------------------------------- utilità di parsing
def _date(s):
    m = re.search(r"(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})", s or "")
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:
        y += 2000
    if not (1 <= d <= 31 and 1 <= mo <= 12):
        return None
    return "%04d-%02d-%02d" % (y, mo, d)


def _it(iso):
    return "%s/%s/%s" % (iso[8:10], iso[5:7], iso[:4]) if iso else ""


def _phones(s):
    out = []
    for m in re.finditer(r"(?<![\d])(\+?39[ .]?)?(3\d{2}[ .\-]?\d{6,7}|0\d{1,3}[ .\-]?\d{5,8})(?![\d])", s or ""):
        p = re.sub(r"[ .\-]", "", (m.group(0) or ""))
        if p not in out:
            out.append(p)
    return out


def _field(text, label, stop=None):
    """Valore dopo «Etichetta:» sulla stessa riga (e righe successive fino a `stop`)."""
    m = re.search(r"(?im)^\s*" + label + r"\s*:?[ \t]*(.*)$", text)
    if not m:
        return ""
    val = m.group(1).strip()
    if stop:
        rest = text[m.end():]
        extra = []
        for line in rest.split("\n")[1:]:
            if not line.strip() or re.match(stop, line.strip(), re.I):
                break
            extra.append(line.strip())
        val = " ".join([val] + extra).strip()
    return val


def _block(text, header, stop_re):
    """Testo tra una intestazione di sezione e la successiva."""
    m = re.search(r"(?im)^\s*" + header + r"\s*:?\s*$", text)
    if not m:
        return ""
    rest = text[m.end():]
    s = re.search(stop_re, rest)
    return (rest[:s.start()] if s else rest).strip()


def _title(s):
    PICCOLE = {"DEL", "DI", "DA", "DAL", "DEI", "DELLA", "DELLE", "DEGLI", "DELL", "NEL", "SUL", "IN", "E", "AL", "ALLA"}
    out = []
    for i, w in enumerate((s or "").split()):
        if w.isupper() and w in PICCOLE and i > 0:
            out.append(w.lower())
        else:
            out.append(w.capitalize() if w.isupper() and len(w) > 3 else w)
    return " ".join(out)


def detect(text):
    head = text[:1500]
    subj = re.search(r"(?im)^\s*Oggetto\s*:\s*(.+)$", text)
    s = subj.group(1) if subj else head.split("\n")[0]
    if re.search(r"(?i)\bdisdett", s) or re.search(r"(?i)data di cessazione desiderata", text):
        return "disdetta", s
    # testo incollato senza la riga «Oggetto»: si riconosce dal corpo della mail
    if re.search(r"(?i)recupero (della )?disdetta|disdetta (immediata|con cliente)|Data (spedizione|ricezione) disdetta|Canone mensile originale|Agente Gestore\s*:", head + text[-800:]):
        return "disdetta", s
    if re.search(r"(?i)\bLEAD\b", s) or re.search(r"(?i)INDIRIZZO APPUNTAMENTO|NOMINATIVO\s*-\s*TELEFONO", text):
        return "lead", s
    return "", s


def parse_lead(text, subj):
    r = {"tipo": "lead"}
    m = re.search(r"LEAD\s*[nN°.]*\s*(\d+)\s*[-–—]\s*(.+)$", subj)
    comune = prov = cliente = ""
    if m:
        r["rif_esterno"] = "LEAD " + m.group(1)
        parts = [p.strip() for p in re.split(r"\s+[-–—]\s+", m.group(2).strip()) if p.strip()]
        if parts:
            last = parts[-1] if len(parts) > 1 else ""
            mm = re.match(r"(.+?)\s*\((\w{2})\)\s*$", last)
            if mm:
                comune, prov = mm.group(1).strip(), mm.group(2).upper()
                cliente = " - ".join(parts[:-1])
            else:
                cliente = " - ".join(parts)
    ind = _block(text, r"INDIRIZZO APPUNTAMENTO", r"(?im)^\s*(NOMINATIVO|NOTE APPUNTAMENTO|COMPILAZIONE)")
    ind = " ".join(l.strip() for l in ind.split("\n") if l.strip())
    if not prov:
        mm = re.search(r"\((\w{2})\)\s*$", ind)
        prov = mm.group(1).upper() if mm else ""
    if not comune and ind:
        mm = re.search(r"[-–—]\s*([^-–—()]+?)\s*(?:\(\w{2}\))?\s*$", ind)
        comune = mm.group(1).strip() if mm else ""
    nom = _block(text, r"NOMINATIVO\s*[-–—]\s*TELEFONO", r"(?im)^\s*(NOTE APPUNTAMENTO|COMPILAZIONE|INDIRIZZO)")
    nom = " ".join(l.strip() for l in nom.split("\n") if l.strip())
    tel = _phones(nom) or _phones(text)
    referente = re.sub(r"(?i)\s*[-–—]?\s*tel\.?\s*:?.*$", "", nom).strip(" -–—:")
    referente = re.sub(r"(?i)^(sig\.?ra|sig\.?|dott\.?|dr\.?|ing\.?|geom\.?)\s+", "", referente).strip()
    note = _block(text, r"NOTE APPUNTAMENTO", r"(?im)^\s*(COMPILAZIONE ESITO|ESITO LEAD)")
    note = " ".join(l.strip() for l in note.split("\n") if l.strip())
    mod = re.search(r"(?i)MODALIT\S*\s*:\s*([^\n]+)", text)
    dm = re.search(r"(?im)^\s*DATA\s*:\s*(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})(?:\s*alle\s*ore\s*(\d{1,2})[.:](\d{2}))?", text)
    if dm:
        iso = _date(dm.group(1))
        if iso:
            r["appuntamento"] = "%s %02d:%s" % (iso, int(dm.group(2) or 9), dm.group(3) or "00")
    r.update({"cliente": cliente or _title(referente), "referente": _title(referente), "telefono": tel[0] if tel else "",
              "localita": _title(comune), "provincia": prov, "indirizzo": ind, "modalita": (mod.group(1).strip() if mod else "")})
    righe = []
    if r["referente"]:
        righe.append("Referente: %s%s" % (r["referente"], " - tel. " + r["telefono"] if r["telefono"] else ""))
    if r["modalita"]:
        righe.append("Modalità appuntamento: %s" % r["modalita"].title())
    if r.get("appuntamento"):
        righe.append("Appuntamento richiesto: %s ore %s" % (_it(r["appuntamento"][:10]), r["appuntamento"][11:16]))
    if ind:
        righe.append("Indirizzo appuntamento: %s" % ind)
    if note:
        righe.append("Note: %s" % note)
    r["note"] = "\n".join(righe)
    return r


def parse_disdetta(text, subj):
    r = {"tipo": "disdetta"}
    cl = _field(text, r"Cliente")
    m = re.match(r"(.+?)\s*[-–—]\s*(\d+)\s*$", cl)
    cliente, cod = (m.group(1).strip(), m.group(2)) if m else (cl, "")
    if not cliente:
        ms = re.search(r"(?i)disdetta\s+(.+?)\s*[-–—]\s*(\d+)", subj)
        if ms:
            cliente, cod = ms.group(1).strip(), ms.group(2)
    via, cap, prov, loc = _field(text, r"Via"), _field(text, r"CAP"), _field(text, r"Provincia"), _field(text, r"Localit\S*")
    contratto = _field(text, r"Contratto")
    contratto = contratto.split()[0] if contratto else ""
    cess = _date(_field(text, r"Data di cessazione desiderata"))
    ric = _date(_field(text, r"Data ricezione disdetta"))
    motiv = _field(text, r"Motivazione", stop=r"(Data ricezione|Canone|Importo|Agente)")
    canone = _field(text, r"Canone complessivo")
    canone_o = _field(text, r"Canone mensile originale")
    scaduto = _field(text, r"Importo Scaduto del Cliente")
    acq, gest = _field(text, r"Agente acquisitore"), _field(text, r"Agente gestore")
    verb = _block(text, r"Verbale telefonico", r"(?im)^\s*(Recapiti\s*:|Attendiamo riscontro|Vi chiediamo)")
    verb = " ".join(l.strip() for l in verb.split("\n") if l.strip())
    rec = _field(text, r"Recapiti")
    tel = _phones(rec) or _phones(verb) or _phones(text)
    ref = "DISDETTA " + (contratto or cod) if (contratto or cod) else ""
    m_c = re.search(r"(\d[\d.]*(?:,\d{1,2})?)", canone or "")
    try:
        can_num = float(m_c.group(1).replace(".", "").replace(",", ".")) if m_c else None
    except ValueError:
        can_num = None
    r["canone"] = can_num
    r.update({"cliente": _title(cliente), "telefono": tel[0] if tel else "", "localita": _title(loc), "provincia": prov.upper()[:2],
              "indirizzo": (via + (" - CAP " + cap if cap else "")).strip(), "rif_esterno": ref, "agente_gestore": gest,
              "agente_acquisitore": acq, "cessazione": cess})
    righe = ["Disdetta contratto %s (cliente n. %s)" % (contratto or "?", cod or "?")]
    if cess or ric:
        righe.append("Cessazione richiesta: %s - Disdetta ricevuta il: %s" % (_it(cess) or "?", _it(ric) or "?"))
    if motiv:
        righe.append("Motivazione: %s" % motiv)
    if canone or canone_o or scaduto:
        righe.append("Canone: %s - Canone mensile originale: %s - Importo scaduto: %s" % (canone or "?", canone_o or "?", scaduto or "?"))
    if acq or gest:
        righe.append("Agente acquisitore: %s - Agente gestore: %s" % (acq or "?", gest or "?"))
    if rec:
        righe.append("Recapiti: %s" % rec)
    if verb:
        righe.append("Verbale telefonico: %s" % verb)
    r["note"] = "\n".join(righe)
    return r


def parse(text):
    tipo, subj = detect(text)
    if tipo == "lead":
        out = parse_lead(text, subj)
    elif tipo == "disdetta":
        out = parse_disdetta(text, subj)
    else:
        out = {"tipo": "", "cliente": "", "telefono": (_phones(text) or [""])[0], "localita": "", "provincia": "", "indirizzo": "",
               "rif_esterno": "", "note": text[:1500]}
    out["oggetto"] = subj.strip() if re.search(r"(?im)^\s*Oggetto\s*:", text) else ""
    return out
