#!/usr/bin/env python3
"""Extract text, tables, figures and EMPHASIS (underline / highlight) from every file in a folder.

Usage:  python extract_sources.py <folder> <out_dir> [--include "sub1,sub2"]

--include  comma-separated names of first-level subfolders whose files (recursively) are extracted too.
           Without it, subfolders are only LISTED in inventory.md (so the user can be asked about them).

Writes to <out_dir>:
  inventory.md            one row per file: type, size, role guess, emphasis counts; a Subfolders table
  sources/<name>.md       reading-order text; emphasis inline as  ⟦U: ...⟧  (underlined)
                          and  ⟦H:<colour>: ...⟧  (highlighted / shaded); tables as | rows |;
                          figures as [FIGURE: media/<file>]
  emphasis.md             every emphasised passage, grouped by file and by mark type
  media/                  images from .docx/.pptx/.pdf (EMF/WMF converted to PNG on Windows)

Only the standard library is required. Optional: pypdf or PyMuPDF (PDF), openpyxl (xlsx), Pillow.
"""
import io, os, re, sys, json, zipfile, shutil, subprocess, collections
from xml.dom import minidom

TEXT_EXT = {".txt", ".md", ".csv", ".tsv", ".py", ".fll", ".json", ".tex", ".bib", ".m", ".r", ".yaml", ".yml"}
IMG_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".svg", ".emf", ".wmf"}
SKIP_PREFIX = ("~$", ".~", "~WRL")
SKIP_DIRS = {"__pycache__", "node_modules", ".git", "venv", ".venv"}
MAX_TEXT_CHARS = 60000

def safe(name):
    return re.sub(r"[^\w.\-]+", "_", name)[:120]

# ------------------------------------------------------------------ DOCX
def docx_extract(path, media_dir, tag):
    z = zipfile.ZipFile(path)
    doc = minidom.parseString(z.read("word/document.xml"))
    rels = {}
    if "word/_rels/document.xml.rels" in z.namelist():
        for r in minidom.parseString(z.read("word/_rels/document.xml.rels")).getElementsByTagName("Relationship"):
            rels[r.getAttribute("Id")] = r.getAttribute("Target")
    # character/paragraph styles that carry underline or highlight
    style_u, style_h = set(), {}
    if "word/styles.xml" in z.namelist():
        for st in minidom.parseString(z.read("word/styles.xml")).getElementsByTagName("w:style"):
            sid = st.getAttribute("w:styleId")
            for u in st.getElementsByTagName("w:u"):
                if u.getAttribute("w:val") not in ("", "none"): style_u.add(sid)
            for h in st.getElementsByTagName("w:highlight"):
                style_h[sid] = h.getAttribute("w:val")
    saved_media = {}

    def save_media(rid):
        tgt = rels.get(rid)
        if not tgt: return None
        zp = "word/" + tgt if not tgt.startswith("/") else tgt[1:]
        if zp not in z.namelist(): return None
        if zp not in saved_media:
            fn = f"{tag}__{os.path.basename(zp)}"
            with open(os.path.join(media_dir, fn), "wb") as f: f.write(z.read(zp))
            saved_media[zp] = fn
        return saved_media[zp]

    def run_marks(r, pstyle):
        u, h = False, None
        rpr = r.getElementsByTagName("w:rPr")
        if rpr:
            rp = rpr[0]
            for el in rp.getElementsByTagName("w:u"):
                if el.getAttribute("w:val") not in ("", "none"): u = True
            for el in rp.getElementsByTagName("w:highlight"):
                if el.getAttribute("w:val") not in ("", "none"): h = el.getAttribute("w:val")
            for el in rp.getElementsByTagName("w:shd"):
                fill = el.getAttribute("w:fill")
                if fill and fill.lower() not in ("auto", "ffffff", "none") and not h: h = "shade-" + fill
            for el in rp.getElementsByTagName("w:rStyle"):
                sid = el.getAttribute("w:val")
                if sid in style_u: u = True
                if sid in style_h and not h: h = style_h[sid]
        if pstyle in style_u: u = True
        return u, h

    def para_text(p):
        ps = p.getElementsByTagName("w:pStyle")
        pstyle = ps[0].getAttribute("w:val") if ps else ""
        segs = []  # (text, u, h)
        for node in p.getElementsByTagName("w:r"):
            t = "".join(x.firstChild.data for x in node.getElementsByTagName("w:t") if x.firstChild)
            t += "\t" * len(node.getElementsByTagName("w:tab"))
            for blip in node.getElementsByTagName("a:blip"):
                fn = save_media(blip.getAttribute("r:embed"))
                if fn: segs.append((f" [FIGURE: media/{fn}] ", False, None))
            for im in node.getElementsByTagName("v:imagedata"):
                fn = save_media(im.getAttribute("r:id"))
                if fn: segs.append((f" [FIGURE: media/{fn}] ", False, None))
            if t:
                u, h = run_marks(node, pstyle)
                segs.append((t, u, h))
        # merge consecutive segments with identical marks
        merged = []
        for t, u, h in segs:
            if merged and merged[-1][1] == u and merged[-1][2] == h and not t.startswith(" [FIGURE"):
                merged[-1] = (merged[-1][0] + t, u, h)
            else:
                merged.append((t, u, h))
        out, emph = [], []
        for t, u, h in merged:
            if (u or h) and t.strip():
                if u: out.append(f"⟦U: {t}⟧"); emph.append(("underline", t.strip()))
                if h and not u: out.append(f"⟦H:{h}: {t}⟧"); emph.append((f"highlight:{h}", t.strip()))
                if h and u: emph.append((f"highlight:{h}", t.strip()))
            else:
                out.append(t)
        return pstyle, "".join(out), emph

    lines, emphasis = [], []
    body = doc.getElementsByTagName("w:body")[0]
    for el in body.childNodes:
        if el.nodeName == "w:p":
            pstyle, t, e = para_text(el)
            emphasis += e
            if t.strip():
                prefix = "## " if re.match(r"(?i)(heading|nag|title|tytu)", pstyle or "") else ""
                lines.append(prefix + t.strip())
        elif el.nodeName == "w:tbl":
            lines.append("")
            for tr in el.getElementsByTagName("w:tr"):
                cells = []
                for tc in tr.getElementsByTagName("w:tc"):
                    parts = []
                    for p in tc.getElementsByTagName("w:p"):
                        _, t, e = para_text(p); emphasis += e; parts.append(t.strip())
                    cells.append(" ".join(x for x in parts if x))
                lines.append("| " + " | ".join(cells) + " |")
            lines.append("")
    # footnotes / comments are useful context
    for extra, label in (("word/footnotes.xml", "Footnotes"), ("word/comments.xml", "Comments")):
        if extra in z.namelist():
            d = minidom.parseString(z.read(extra))
            txt = [ "".join(t.firstChild.data for t in p.getElementsByTagName("w:t") if t.firstChild) for p in d.getElementsByTagName("w:p")]
            txt = [t for t in txt if t.strip()]
            if txt: lines += ["", f"## [{label}]"] + txt
    # media not referenced inline
    for n in z.namelist():
        if n.startswith("word/media/") and n not in saved_media:
            fn = f"{tag}__{os.path.basename(n)}"
            with open(os.path.join(media_dir, fn), "wb") as f: f.write(z.read(n))
            saved_media[n] = fn
    return "\n".join(lines), emphasis, list(saved_media.values())

# ------------------------------------------------------------------ PPTX
def pptx_extract(path, media_dir, tag):
    z = zipfile.ZipFile(path)
    slides = sorted([n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)], key=lambda n: int(re.findall(r"\d+", n)[-1]))
    lines, emphasis, media = [], [], []
    for sn in slides:
        d = minidom.parseString(z.read(sn))
        lines.append(f"## Slide {re.findall(r'\d+', sn)[-1]}")
        for p in d.getElementsByTagName("a:p"):
            out = []
            for r in p.getElementsByTagName("a:r"):
                t = "".join(x.firstChild.data for x in r.getElementsByTagName("a:t") if x.firstChild)
                rpr = r.getElementsByTagName("a:rPr")
                u = rpr and rpr[0].getAttribute("u") not in ("", "none")
                hl = rpr and rpr[0].getElementsByTagName("a:highlight")
                if u and t.strip(): out.append(f"⟦U: {t}⟧"); emphasis.append(("underline", t.strip()))
                elif hl and t.strip(): out.append(f"⟦H:ppt: {t}⟧"); emphasis.append(("highlight:ppt", t.strip()))
                else: out.append(t)
            if "".join(out).strip(): lines.append("".join(out))
        notes = sn.replace("slides/slide", "notesSlides/notesSlide")
        if notes in z.namelist():
            nd = minidom.parseString(z.read(notes))
            nt = " ".join(x.firstChild.data for x in nd.getElementsByTagName("a:t") if x.firstChild)
            if nt.strip(): lines.append(f"[Notes] {nt.strip()}")
    for n in z.namelist():
        if n.startswith("ppt/media/"):
            fn = f"{tag}__{os.path.basename(n)}"
            with open(os.path.join(media_dir, fn), "wb") as f: f.write(z.read(n))
            media.append(fn)
    return "\n".join(lines), emphasis, media

# ------------------------------------------------------------------ PDF
def pdf_extract(path, media_dir, tag):
    emphasis, media, lines = [], [], []
    try:
        import fitz  # PyMuPDF: best — text under highlight/underline annotations + images
        doc = fitz.open(path)
        for i, page in enumerate(doc, 1):
            lines.append(f"## Page {i}")
            lines.append(page.get_text("text"))
            for a in page.annots() or []:
                kind = a.type[1]
                if kind in ("Highlight", "Underline", "Squiggly", "StrikeOut"):
                    t = page.get_textbox(a.rect).strip() or (a.info.get("content") or "").strip()
                    if t and kind != "StrikeOut":
                        emphasis.append(("underline" if kind == "Underline" else "highlight:pdf", t))
                        lines.append(("⟦U: " if kind == "Underline" else "⟦H:pdf: ") + t + "⟧")
                elif a.info.get("content"):
                    lines.append(f"[Annotation p.{i}] {a.info['content']}")
            for j, img in enumerate(page.get_images(full=True)):
                try:
                    pix = fitz.Pixmap(doc, img[0])
                    if pix.width < 150 or pix.height < 150: continue
                    if pix.n > 4: pix = fitz.Pixmap(fitz.csRGB, pix)
                    fn = f"{tag}__p{i}_img{j}.png"; pix.save(os.path.join(media_dir, fn)); media.append(fn)
                    lines.append(f"[FIGURE: media/{fn}]")
                except Exception:
                    pass
        return "\n".join(lines), emphasis, media
    except ImportError:
        pass
    try:
        import pypdf
        r = pypdf.PdfReader(path)
        for i, p in enumerate(r.pages, 1):
            lines.append(f"## Page {i}")
            lines.append(p.extract_text() or "")
            for a in p.get("/Annots") or []:
                a = a.get_object()
                st = a.get("/Subtype")
                if st in ("/Highlight", "/Underline", "/Squiggly"):
                    t = str(a.get("/Contents") or "").strip() or f"(annotation on page {i}; text not recoverable without PyMuPDF)"
                    emphasis.append(("underline" if st == "/Underline" else "highlight:pdf", t))
        return "\n".join(lines), emphasis, media
    except ImportError:
        return "(PDF text extraction unavailable: pip install pymupdf or pypdf)", [], []

# ------------------------------------------------------------------ XLSX
def xlsx_extract(path):
    try:
        import openpyxl
    except ImportError:
        return "(xlsx: pip install openpyxl to read values)", []
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    lines = []
    for ws in wb.worksheets:
        lines.append(f"## Sheet: {ws.title}")
        for k, row in enumerate(ws.iter_rows(values_only=True)):
            if k >= 200: lines.append("| ... (truncated at 200 rows) |"); break
            if any(v is not None for v in row):
                lines.append("| " + " | ".join("" if v is None else str(v) for v in row) + " |")
    return "\n".join(lines), []

# ------------------------------------------------------------------ helpers

def run_ps(script):
    """Run a multi-line PowerShell script reliably (-EncodedCommand avoids newline/quoting issues)."""
    import base64
    enc = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    return subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", enc], capture_output=True, text=True, encoding="utf-8", errors="replace")
def convert_metafiles(media_dir):
    """EMF/WMF -> PNG via .NET System.Drawing (Windows only). Silent no-op elsewhere."""
    metas = [f for f in os.listdir(media_dir) if f.lower().endswith((".emf", ".wmf"))]
    if not metas or os.name != "nt": return
    ps = r'''
Add-Type -AssemblyName System.Drawing
Get-ChildItem -LiteralPath "%s" | Where-Object { $_.Extension -in ".emf",".wmf" } | ForEach-Object {
  try {
    $m = New-Object System.Drawing.Imaging.Metafile($_.FullName)
    $s = [Math]::Min(3.0, 2400 / [Math]::Max($m.Width, $m.Height))
    $b = New-Object System.Drawing.Bitmap([int]($m.Width*$s), [int]($m.Height*$s))
    $g = [System.Drawing.Graphics]::FromImage($b); $g.Clear([System.Drawing.Color]::White)
    $g.SmoothingMode = "HighQuality"; $g.InterpolationMode = "HighQualityBicubic"
    $g.DrawImage($m, 0, 0, $b.Width, $b.Height)
    $b.Save(($_.FullName -replace '\.(emf|wmf)$', '_$1.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $b.Dispose(); $m.Dispose()
  } catch {}
}''' % os.path.abspath(media_dir)
    run_ps(ps)

REVISION_CUES = re.compile(r"(?i)(proponowany|propozycja|\blub:|\bor:|proposed text|alternative wording|tu wstaw|todo|\[repository|\[doi|do poprawy)")
def revision_hint(t):
    polish = len(re.findall(r"[\u0105\u0107\u0119\u0142\u0144\u00f3\u015b\u017a\u017c]", t.lower())) >= 2
    if REVISION_CUES.search(t) or polish:
        return "   (!) likely revision/editorial mark (draft wording or non-English) - confirm before using as key content"
    return ""

def image_size(p):
    try:
        from PIL import Image
        with Image.open(p) as im: return f"{im.width}x{im.height}"
    except Exception:
        return "?"

def role_guess(name, text):
    n = name.lower()
    if name.startswith(SKIP_PREFIX) or n.endswith((".tmp", ".lock")) or n in ("skills-lock.json",): return "temp/lock — ignore"
    if os.path.splitext(n)[1] in IMG_EXT: return "figure"
    if any(k in n for k in ("guide for authors", "guideline", "template", "style", "how to", "instruction")): return "style/meta (guidelines) — not content"
    if "prompt" in n: return "style/meta (prompt) — not content"
    if n.endswith(".pptx") and ("presentation" in n or "deck" in n or "slides" in n): return "earlier deck — style/meta unless confirmed"
    if n.endswith((".docx", ".pdf", ".tex")) and len(text) > 15000: return "main content source (long document)"
    if n.endswith((".xlsx", ".csv")): return "data source"
    if os.path.splitext(n)[1] in IMG_EXT: return "figure"
    return "possible content source"

# ------------------------------------------------------------------ main
def main():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)
    folder, out = sys.argv[1], sys.argv[2]
    args = sys.argv[3:]
    include = [s.strip().strip("/\\") for s in args[args.index("--include") + 1].split(",") if s.strip()] if "--include" in args else []
    src_dir, media_dir = os.path.join(out, "sources"), os.path.join(out, "media")
    os.makedirs(src_dir, exist_ok=True); os.makedirs(media_dir, exist_ok=True)
    inv, emph_all = [], collections.OrderedDict()
    # subfolders: listed in the inventory; their files are extracted only when named in --include
    subfolders = []
    for d in sorted(os.listdir(folder)):
        dp = os.path.join(folder, d)
        if not os.path.isdir(dp) or d.startswith(".") or d in SKIP_DIRS: continue
        files = []
        for root, dirs, fns in os.walk(dp):
            dirs[:] = [x for x in dirs if not x.startswith(".") and x not in SKIP_DIRS]
            for fn in sorted(fns):
                if not fn.startswith(".") and not fn.startswith(SKIP_PREFIX):
                    files.append(os.path.relpath(os.path.join(root, fn), folder).replace("\\", "/"))
        subfolders.append((d, files, d in include))
    entries = [n for n in sorted(os.listdir(folder)) if not os.path.isdir(os.path.join(folder, n))]
    for d, files, inc in subfolders:
        if inc: entries += files
    for name in entries:
        p = os.path.join(folder, name)
        if os.path.basename(name).startswith("."):
            continue
        ext = os.path.splitext(name)[1].lower()
        tag = safe(os.path.splitext(os.path.basename(name))[0])[:24]  # short: Windows GDI+ fails on paths > 260 chars
        text, emph, media = "", [], []
        if os.path.basename(name).startswith(SKIP_PREFIX) or ext in (".tmp",):
            inv.append((name, ext, os.path.getsize(p), "temp/lock — ignore", {}, 0)); continue
        try:
            if ext == ".docx": text, emph, media = docx_extract(p, media_dir, tag)
            elif ext == ".pptx": text, emph, media = pptx_extract(p, media_dir, tag)
            elif ext == ".pdf": text, emph, media = pdf_extract(p, media_dir, tag)
            elif ext in (".xlsx", ".xlsm"): text, emph = xlsx_extract(p)
            elif ext in TEXT_EXT:
                with open(p, encoding="utf-8", errors="replace") as f: text = f.read(MAX_TEXT_CHARS)
            elif ext in IMG_EXT:
                mname = safe(name)
                shutil.copy(p, os.path.join(media_dir, mname)); media = [mname]
                text = f"[FIGURE: media/{mname}] ({image_size(p)} px)"
            else:
                text = "(binary or unsupported file type — not extracted)"
        except Exception as e:
            text = f"(extraction failed: {e})"
        with open(os.path.join(src_dir, safe(name) + ".md"), "w", encoding="utf-8") as f:
            f.write(f"# {name}\n\n{text}\n")
        counts = collections.Counter(k for k, _ in emph)
        if emph: emph_all[name] = emph
        inv.append((name, ext, os.path.getsize(p), role_guess(name, text), dict(counts), len(media)))
    convert_metafiles(media_dir)
    with open(os.path.join(out, "inventory.md"), "w", encoding="utf-8") as f:
        f.write("# Inventory\n\n| File | Type | Size [kB] | Role (guess) | Emphasis marks | Media |\n|---|---|---|---|---|---|\n")
        for name, ext, size, role, counts, nm in inv:
            c = ", ".join(f"{k}: {v}" for k, v in counts.items()) or "—"
            f.write(f"| {name} | {ext} | {size/1024:.0f} | {role} | {c} | {nm} |\n")
        unconv = [fn for fn in os.listdir(media_dir) if fn.lower().endswith((".emf", ".wmf")) and not os.path.exists(os.path.join(media_dir, re.sub(r"\.(emf|wmf)$", r"_\1.png", fn, flags=re.I)))]
        if unconv:
            f.write("\n**WARNING - metafiles not converted to PNG** (cannot be placed on slides as-is): " + ", ".join(unconv) + ". Look for a PNG version of the same figure in the folder, or tell the user.\n")
        f.write("\nMedia files (with pixel size):\n\n")
        for fn in sorted(os.listdir(media_dir)):
            f.write(f"- media/{fn} ({image_size(os.path.join(media_dir, fn))})\n")
        if subfolders:
            f.write("\n## Subfolders\n\nFiles in subfolders are extracted only when the subfolder is named in `--include`. "
                    "Ask the user what role each non-empty subfolder plays before using or ignoring it.\n\n"
                    "| Subfolder | Files | Status | Examples |\n|---|---|---|---|\n")
            for d, files, inc in subfolders:
                ex = ", ".join(files[:5]) + (" …" if len(files) > 5 else "")
                st = "extracted" if inc else ("empty" if not files else "NOT extracted - ask the user")
                f.write(f"| {d} | {len(files)} | {st} | {ex or '—'} |\n")
    with open(os.path.join(out, "emphasis.md"), "w", encoding="utf-8") as f:
        f.write("# Emphasised passages\n\nUnderline is the default signal of key content. Highlights are often revision marks — check their colour and wording before treating them as key content.\n\n")
        if not emph_all: f.write("No underlined or highlighted text found in any file.\n")
        for name, items in emph_all.items():
            f.write(f"## {name}\n\n")
            by = collections.OrderedDict()
            for k, t in items: by.setdefault(k, []).append(t)
            for k, ts in by.items():
                f.write(f"### {k} ({len(ts)} passages, {sum(len(t) for t in ts)} chars)\n\n")
                for t in ts: f.write(f"- {t}{revision_hint(t)}\n")
                f.write("\n")
    print(json.dumps({"files": len(inv), "with_emphasis": list(emph_all.keys()), "out": out,
                      "subfolders": [{"name": d, "files": len(fs), "extracted": inc} for d, fs, inc in subfolders]}, ensure_ascii=False))

if __name__ == "__main__":
    main()
