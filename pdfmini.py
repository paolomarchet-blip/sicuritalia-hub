# -*- coding: utf-8 -*-
"""Piccolo scrittore di PDF (solo Python, nessuna libreria esterna): testo, rettangoli e linee su una pagina A4.
Serve a compilare il preventivo sopra le pagine fisse del modello Sicuritalia (le unisce pypdf).
Font: Helvetica e Helvetica-Bold (standard PDF, metriche identiche ad Arial)."""
import zlib

_W = {'F1': [761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584, 761, 556, 278, 222, 556, 333, 1000, 556, 556, 333, 1000, 667, 333, 1000, 278, 611, 278, 278, 222, 222, 333, 333, 350, 556, 1000, 333, 1000, 500, 333, 944, 278, 500, 667, 278, 333, 556, 556, 556, 556, 260, 556, 333, 737, 370, 556, 584, 333, 737, 333, 400, 584, 333, 333, 333, 556, 537, 278, 333, 333, 365, 556, 834, 834, 834, 611, 667, 667, 667, 667, 667, 667, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611, 556, 556, 556, 556, 556, 556, 889, 500, 556, 556, 556, 556, 278, 278, 278, 278, 556, 556, 556, 556, 556, 556, 556, 584, 611, 556, 556, 556, 556, 500, 556, 500], 'F2': [761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 761, 278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556, 333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584, 761, 556, 278, 278, 556, 500, 1000, 556, 556, 333, 1000, 667, 333, 1000, 278, 611, 278, 278, 278, 278, 500, 500, 350, 556, 1000, 333, 1000, 556, 333, 944, 278, 500, 667, 278, 333, 556, 556, 556, 556, 280, 556, 333, 737, 370, 556, 584, 333, 737, 333, 400, 584, 333, 333, 333, 611, 556, 278, 333, 333, 365, 556, 834, 834, 834, 611, 722, 722, 722, 722, 722, 722, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611, 556, 556, 556, 556, 556, 556, 889, 556, 556, 556, 556, 556, 278, 278, 278, 278, 611, 611, 611, 611, 611, 611, 611, 584, 611, 611, 611, 611, 611, 556, 611, 556]}
A4 = (595.2756, 841.8898)


def _esc(s):
    b = s.encode("cp1252", "replace")
    return b.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)")


def larghezza(s, font="F1", size=10):
    t = _W[font]
    return sum(t[b] for b in s.encode("cp1252", "replace")) * size / 1000.0


def a_capo(s, font, size, maxw):
    """Spezza il testo in righe che stanno in maxw punti."""
    righe = []
    for par in str(s).split("\n"):
        cur = ""
        for w in par.split():
            prova = (cur + " " + w).strip()
            if larghezza(prova, font, size) <= maxw or not cur:
                cur = prova
            else:
                righe.append(cur)
                cur = w
        righe.append(cur)
    return righe


class Pagina:
    def __init__(self):
        self.ops = []
        self.imgs = []      # (jpeg, larghezza px, altezza px)

    def _c(self, rgb):
        return "%.3f %.3f %.3f" % tuple(x / 255.0 for x in rgb)

    def rett(self, x, y, w, h, fill=None, stroke=None, lw=0.6):
        """Rettangolo con angolo in alto a sinistra (x, y) misurato dall'alto della pagina."""
        yy = A4[1] - y - h
        if fill:
            self.ops.append("%s rg %.2f %.2f %.2f %.2f re f" % (self._c(fill), x, yy, w, h))
        if stroke:
            self.ops.append("%s RG %.2f w %.2f %.2f %.2f %.2f re S" % (self._c(stroke), lw, x, yy, w, h))

    def linea(self, x1, y1, x2, y2, rgb=(0, 0, 0), lw=0.6):
        self.ops.append("%s RG %.2f w %.2f %.2f m %.2f %.2f l S" % (self._c(rgb), lw, x1, A4[1] - y1, x2, A4[1] - y2))

    def testo(self, x, y, s, font="F1", size=10, rgb=(0, 0, 0), align="l"):
        """y = linea di base misurata dall'alto. align: l, c (x = centro), r (x = bordo destro)."""
        w = larghezza(s, font, size)
        if align == "c":
            x -= w / 2.0
        elif align == "r":
            x -= w
        self.ops.append(("BT %s rg /%s %.2f Tf %.2f %.2f Td" % (self._c(rgb), font, size, x, A4[1] - y), _esc(s)))

    def img(self, x, y, w, h, jpeg, pw, ph):
        """Immagine JPEG nel rettangolo con angolo in alto a sinistra (x, y), dall'alto della pagina."""
        self.imgs.append((jpeg, pw, ph))
        self.ops.append("q %.2f 0 0 %.2f %.2f %.2f cm /Im%d Do Q" % (w, h, x, A4[1] - y - h, len(self.imgs) - 1))

    def pdf(self):
        parts = []
        for op in self.ops:
            if isinstance(op, tuple):
                parts.append(op[0].encode("latin-1") + b" (" + op[1] + b") Tj ET")
            else:
                parts.append(op.encode("latin-1"))
        cont = zlib.compress(b"\n".join(parts))
        objs = [b"<< /Type /Catalog /Pages 2 0 R >>",
                b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                ("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %.4f %.4f] /Resources << /Font << /F1 5 0 R /F2 6 0 R >> >> /Contents 4 0 R >>" % A4).encode(),
                b"<< /Length %d /Filter /FlateDecode >>\nstream\n" % len(cont) + cont + b"\nendstream",
                b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
                b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>"]
        if self.imgs:
            xo = " ".join("/Im%d %d 0 R" % (i, 7 + i) for i in range(len(self.imgs)))
            objs[2] = ("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %.4f %.4f] /Resources << /Font << /F1 5 0 R /F2 6 0 R >> /XObject << %s >> >> /Contents 4 0 R >>" % (A4[0], A4[1], xo)).encode()
            for jp, pw, ph in self.imgs:
                objs.append(("<< /Type /XObject /Subtype /Image /Width %d /Height %d /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length %d >>\nstream\n" % (pw, ph, len(jp))).encode() + jp + b"\nendstream")
        out = bytearray(b"%PDF-1.4\n")
        offs = []
        for i, o in enumerate(objs, 1):
            offs.append(len(out))
            out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
        x = len(out)
        out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
        for o in offs:
            out += b"%010d 00000 n \n" % o
        out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, x)
        return bytes(out)
