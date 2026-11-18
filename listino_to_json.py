# Converte il listino Excel in assets/listino.json e le immagini in assets/listino_img/N.jpg
# uso:  python3 listino_to_json.py LISTINO.xlsx assets/listino.json assets/listino_img
import sys, re, json, io, os, openpyxl
from PIL import Image
src, dst = sys.argv[1], sys.argv[2]
dimg = sys.argv[3] if len(sys.argv) > 3 else os.path.join(os.path.dirname(dst), "listino_img")
ws = openpyxl.load_workbook(src, data_only=True).active
img_riga = {}
for im in ws._images:                      # per ogni riga si tiene la prima immagine (in alto a sinistra) della cella
    a = im.anchor._from
    k = (a.rowOff or 0, a.colOff or 0)
    r = a.row + 1
    if r not in img_riga or k < img_riga[r][0]:
        img_riga[r] = (k, im)
os.makedirs(dimg, exist_ok=True)
for f in os.listdir(dimg):
    os.remove(os.path.join(dimg, f))
out, visti = [], set()
for r in range(3, ws.max_row + 1):
    nome = (ws.cell(r, 1).value or "")
    cod = ws.cell(r, 2).value
    desc = ws.cell(r, 4).value or ""
    if not str(nome).strip():
        continue
    nome = " ".join(str(nome).split())
    cod = "" if cod is None else str(cod).strip()
    chiave = cod or nome
    if chiave in visti:
        continue
    visti.add(chiave)
    comp = []
    for ln in str(desc).split("\n"):
        m = re.match(r"^\s*N°\s*(\d+)\s+(?:(\d{8})\s+)?(.+?)\s*$", ln)
        if m:
            comp.append({"q": int(m.group(1)), "c": m.group(2) or "", "n": " ".join(m.group(3).split())})
    e = {"c": cod, "n": nome}
    if comp:
        e["k"] = comp
    if r in img_riga:
        try:
            im = Image.open(io.BytesIO(img_riga[r][1]._data())).convert("RGB")
            if max(im.size) > 160:
                im.thumbnail((160, 160))
            im.save(os.path.join(dimg, "%d.jpg" % len(out)), "JPEG", quality=90)
            e["i"] = len(out)
        except Exception:
            pass
    out.append(e)
json.dump(out, open(dst, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
print(len(out), "articoli,", sum(1 for x in out if x.get("k")), "kit,", sum(1 for x in out if "i" in x), "con immagine")
