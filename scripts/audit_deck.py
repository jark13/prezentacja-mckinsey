#!/usr/bin/env python3
"""Audit a generated deck: structure, speaker notes, time budget and NUMBER GROUNDING.

Usage:  python audit_deck.py <deck.pptx> <sources_dir> [--minutes N] [--only "name1,name2"] [--json out.json]

--only  comma-separated substrings of the CONTENT source files (as confirmed by the user); style/meta
        files are then ignored, so their numbers cannot make an invented number look "grounded".

<sources_dir> is the extract/sources folder written by extract_sources.py (all *.md are read).

Checks per slide
  - action title present (shape named ActionTitle) and not too long (> 2 lines ~ > 120 chars)
  - number of topics (distinct N in shapes named Topic-N-* / BigNumber-N-*) <= 3
  - words on the slide (body text, excluding notes and the method footnote) — warns above 70
  - speaker notes present; word count
  - every number on the slide and in the notes appears somewhere in the sources
  - charts: every axis has a title (quantity + unit); any chart or figure needs a SourceNote
  - statistical detail (p-values, CIs, n, SD) belongs in the MethodNote footnote, not the main area
Deck level
  - notes word total vs. minutes x 125 (+/-15 %), slide count vs. minutes
Exit code 0 always; read the report. Items marked FAIL must be fixed, WARN must be judged.
"""
import io, os, re, sys, json, zipfile
from xml.dom import minidom

NUM_RE = re.compile(r"(?<![\w.])[-−+]?\d+(?:[.,]\d+)?(?:[eE][-−+]?\d+)?")
TRIVIAL = {str(i) for i in range(0, 11)} | {"100"}

def norm_num(s):
    s = s.replace("−", "-").replace(",", ".").lstrip("+")
    try:
        v = float(s)
    except ValueError:
        return None
    return v

def slide_shapes(xml):
    d = minidom.parseString(xml)
    shapes = []
    for sp in d.getElementsByTagName("p:sp") + d.getElementsByTagName("p:pic") + d.getElementsByTagName("p:graphicFrame"):
        c = sp.getElementsByTagName("p:cNvPr")
        name = c[0].getAttribute("name") if c else ""
        paras = []
        for p in sp.getElementsByTagName("a:p"):
            t = "".join(x.firstChild.data for x in p.getElementsByTagName("a:t") if x.firstChild)
            if t.strip(): paras.append(t)
        shapes.append((name, "\n".join(paras)))
    return shapes

def slide_charts(z, slide_name):
    """XML of charts referenced by the slide."""
    rel = slide_name.replace("slides/", "slides/_rels/") + ".rels"
    out = []
    if rel not in z.namelist(): return out
    for tgt in re.findall(r'Target="([^"]+chart[^"]*\.xml)"', z.read(rel).decode("utf-8", "ignore")):
        # pptxgenjs writes package-absolute targets ("/ppt/charts/chart1.xml"); others are relative
        path = tgt.lstrip("/") if tgt.startswith("/") else os.path.normpath(os.path.join(os.path.dirname(slide_name), tgt)).replace("\\", "/")
        if path in z.namelist():
            out.append(z.read(path).decode("utf-8", "ignore"))
    return out

def chart_numbers(charts):
    """Numbers inside the slide's charts (they are evidence and must be grounded too)."""
    return [n for xml in charts for n in re.findall(r"<c:v>([-0-9.eE]+)</c:v>", xml)]

def charts_missing_axis_titles(charts):
    """Axes (cat/val/date) without a <c:title> — every axis needs quantity + unit."""
    missing = 0
    for xml in charts:
        if "<c:pieChart" in xml or "<c:doughnutChart" in xml: continue
        for ax in re.findall(r"<c:(?:catAx|valAx|dateAx)>(.*?)</c:(?:catAx|valAx|dateAx)>", xml, re.S):
            if "<c:delete val=\"1\"/>" in ax: continue
            if "<c:title>" not in ax: missing += 1
    return missing

STATS_RE = re.compile(r"(?i)(\bp\s*[<=>≤]\s*0?[.,]\d|\b\d{2}\s*%\s*CI\b|\bconfidence interval|\bCI\s*[\[(]|\bstd\.?\s*dev|\bσ\s*=|\bn\s*=\s*\d)")

def main():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    args = sys.argv[1:]
    if len(args) < 2: print(__doc__); sys.exit(1)
    deck, src = args[0], args[1]
    minutes = float(args[args.index("--minutes") + 1]) if "--minutes" in args else None
    jout = args[args.index("--json") + 1] if "--json" in args else None
    only = [s.strip().lower() for s in args[args.index("--only") + 1].split(",")] if "--only" in args else []

    corpus = ""
    for fn in os.listdir(src):
        if fn.endswith(".md") and (not only or any(o in fn.lower() for o in only)):
            with open(os.path.join(src, fn), encoding="utf-8", errors="ignore") as f: corpus += f.read() + "\n"
    src_vals = set()
    for m in NUM_RE.findall(corpus):
        v = norm_num(m)
        if v is not None: src_vals.add(round(v, 6))

    def grounded(tok):
        v = norm_num(tok)
        if v is None: return True
        if tok.lstrip("-−+") in TRIVIAL: return True
        for cand in (v, abs(v), v / 100, v * 100, v / 1000, v * 1000):   # %, mm<->m, kg<->g
            if round(cand, 6) in src_vals: return True
        # rounding tolerance: displayed value is a rounding of a source value
        dec = len(tok.split(".")[-1]) if "." in tok.replace(",", ".") else 0
        tol = 0.5 * 10 ** (-dec)
        return any(abs(abs(v) - abs(sv)) <= tol + 1e-9 for sv in src_vals)

    z = zipfile.ZipFile(deck)
    slides = sorted([n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)], key=lambda n: int(re.findall(r"\d+", n)[-1]))
    report, total_notes, fails, warns = [], 0, 0, 0
    for sn in slides:
        idx = int(re.findall(r"\d+", sn)[-1])
        shapes = slide_shapes(z.read(sn))
        names = [n for n, _ in shapes]
        title = next((t for n, t in shapes if n == "ActionTitle"), "")
        topics = sorted({int(m.group(1)) for n in names for m in [re.match(r"(?:Topic|BigNumber)-(\d+)-", n)] if m})
        body_words = sum(len(t.split()) for n, t in shapes if n not in ("ActionTitle", "PageNum", "SourceNote", "Tracker", "MethodNote"))
        method = " ".join(t for n, t in shapes if n == "MethodNote")
        main_text = " ".join(t for n, t in shapes if n not in ("PageNum", "SourceNote", "MethodNote"))
        charts = slide_charts(z, sn)
        has_evidence = bool(charts) or any(n.startswith("Evidence") for n in names)
        source_note = next((t for n, t in shapes if n == "SourceNote"), "")
        notes_xml = sn.replace("slides/slide", "notesSlides/notesSlide")
        notes = ""
        if notes_xml in z.namelist():
            nd = minidom.parseString(z.read(notes_xml))
            for sp in nd.getElementsByTagName("p:sp"):
                ph = sp.getElementsByTagName("p:ph")
                if ph and ph[0].getAttribute("type") in ("sldNum", "sldImg"): continue
                notes += " ".join(x.firstChild.data for x in sp.getElementsByTagName("a:t") if x.firstChild) + " "
        nwords = len(notes.split()); total_notes += nwords
        issues = []
        if not title: issues.append("FAIL no ActionTitle shape")
        elif len(title) > 120: issues.append(f"WARN title long ({len(title)} chars) — aim for <=2 lines")
        if len(topics) > 3: issues.append(f"FAIL {len(topics)} topics (max 3): {topics}")
        if body_words > 70: issues.append(f"WARN {body_words} words on slide — likely overloaded")
        if nwords == 0: issues.append("FAIL no speaker notes")
        if has_evidence and not source_note.strip():
            issues.append("FAIL chart/figure without a source note (SourceNote)")
        nmiss = charts_missing_axis_titles(charts)
        if nmiss: issues.append(f"FAIL {nmiss} chart axis/axes without a title (quantity + unit) — use chartOpts({{xTitle, yTitle}})")
        if STATS_RE.search(main_text):
            issues.append("WARN statistical detail (p-value / CI / n / SD) in the main slide area — move it to methodNote()")
        if len(method.split()) > 45:
            issues.append(f"WARN method note long ({len(method.split())} words) — keep <=2 lines, move the rest to the notes")
        slide_text = " ".join(t for n, t in shapes if n not in ("PageNum",))
        ungrounded = sorted({tok for tok in NUM_RE.findall(slide_text + " " + notes) if not grounded(tok)}
                            | {tok for tok in chart_numbers(charts) if not grounded(tok)})
        # ignore slide-number / timing tokens like "3 / 14" or "[~60 s]"
        ungrounded = [u for u in ungrounded if not re.search(r"\[~?%s\s*s\]" % re.escape(u), notes)]
        if ungrounded: issues.append("FAIL numbers not found in sources: " + ", ".join(ungrounded))
        fails += sum(i.startswith("FAIL") for i in issues); warns += sum(i.startswith("WARN") for i in issues)
        report.append({"slide": idx, "title": title, "topics": topics, "body_words": body_words, "notes_words": nwords, "issues": issues})

    print(f"Deck: {deck}\nSlides: {len(slides)}   Notes words: {total_notes}")
    print("Sources checked: " + (", ".join(only) if only else "ALL files in sources dir (pass --only to restrict to content sources)"))
    if minutes:
        target = minutes * 125
        lo, hi = target * 0.85, target * 1.15
        flag = "OK" if lo <= total_notes <= hi else "WARN"
        print(f"Time budget: {minutes:g} min -> notes target ~{target:.0f} words ({lo:.0f}-{hi:.0f}): {flag}")
        if flag == "WARN": warns += 1
        if not (minutes <= len(slides) <= minutes + 2):
            print(f"WARN slide count {len(slides)} vs {minutes:g} min (expected {minutes:g}-{minutes+2:g} incl. title and closing)"); warns += 1
    print()
    for r in report:
        print(f"[{r['slide']:>2}] {r['title'][:90]}")
        print(f"     topics={len(r['topics'])} words={r['body_words']} notes={r['notes_words']}")
        for i in r["issues"]: print("     " + i)
    print(f"\nSUMMARY: {fails} FAIL, {warns} WARN")
    if jout:
        with open(jout, "w", encoding="utf-8") as f: json.dump({"slides": report, "fails": fails, "warns": warns, "notes_words": total_notes}, f, indent=2, ensure_ascii=False)

if __name__ == "__main__":
    main()
