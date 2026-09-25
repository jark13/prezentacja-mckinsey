#!/usr/bin/env python3
"""Render every slide of a .pptx to PNG (for visual QA) and build a contact sheet.

Usage:  python render_slides.py <deck.pptx> <out_dir> [--slides 2,5,7]

Tries, in order: Microsoft PowerPoint via COM (Windows) -> LibreOffice (soffice) + pdftoppm.
Writes <out_dir>/slide-NN.png and <out_dir>/contact.png (2 columns, needs Pillow).
"""
import os, re, sys, glob, shutil, subprocess

def run_ps(script):
    """Run a multi-line PowerShell script reliably (-EncodedCommand avoids newline/quoting issues)."""
    import base64
    enc = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    return subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", enc], capture_output=True, text=True, encoding="utf-8", errors="replace")

def via_powerpoint(deck, out, only):
    if os.name != "nt": return False
    # PowerPoint cannot open/export paths longer than ~218 characters: work in a short temp dir, then move.
    import tempfile
    # Always work on a copy in a short temp dir: avoids long-path limits and never touches a deck the user has open.
    short = tempfile.mkdtemp(prefix="pr_")
    if short:
        tmp_deck = os.path.join(short, "deck.pptx"); shutil.copy(deck, tmp_deck)
        ok = _pp_export(tmp_deck, short, only)
        for f in glob.glob(os.path.join(short, "slide-*.png")):
            shutil.move(f, os.path.join(out, os.path.basename(f)))
        shutil.rmtree(short, ignore_errors=True)
        return ok and bool(glob.glob(os.path.join(out, "slide-*.png")))
    return _pp_export(deck, out, only)

def _pp_export(deck, out, only):
    sel = ",".join(str(i) for i in only) if only else ""
    ps = r'''
$pp = New-Object -ComObject PowerPoint.Application
# If the user already has PowerPoint open, COM attaches to THEIR instance: never Quit it then.
$pre = $pp.Presentations.Count
$p = $null
try {
  $p = $pp.Presentations.Open("%s", $true, $false, $false)
  if ($null -eq $p) { throw "PowerPoint could not open the file" }
  $sel = "%s"
  foreach ($s in $p.Slides) {
    if ($sel -eq "" -or ($sel.Split(",") -contains "$($s.SlideIndex)")) {
      $s.Export(("%s\slide-{0:D2}.png" -f $s.SlideIndex), "PNG", 1600, 900)
    }
  }
} finally {
  if ($null -ne $p) { $p.Close() }
  if ($pre -eq 0 -and $pp.Presentations.Count -eq 0) { $pp.Quit() }
}
''' % (os.path.abspath(deck), sel, os.path.abspath(out))
    r = run_ps(ps)
    ok = r.returncode == 0 and bool(glob.glob(os.path.join(out, "slide-*.png")))
    if not ok:
        err = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.stderr or "")).strip()[-600:]
        print("PowerPoint rendering failed:", err or "(no message)")
        if "80080005" in err or "CO_E_SERVER_EXEC_FAILURE" in err:
            print("Hint: a hidden/stuck POWERPNT.EXE usually causes this. Ask the user to close PowerPoint "
                  "(or for permission to end the hidden instance), then run again.")
    return ok

def via_libreoffice(deck, out, only):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice or not shutil.which("pdftoppm"): return False
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", out, deck], capture_output=True)
    pdf = os.path.join(out, os.path.splitext(os.path.basename(deck))[0] + ".pdf")
    if not os.path.exists(pdf): return False
    subprocess.run(["pdftoppm", "-png", "-r", "110", pdf, os.path.join(out, "slide")], capture_output=True)
    for f in glob.glob(os.path.join(out, "slide-*.png")):
        n = int(os.path.basename(f)[6:-4])
        dst = os.path.join(out, f"slide-{n:02d}.png")
        if f != dst: os.replace(f, dst)
        if only and n not in only: os.remove(dst)
    return True

def contact(out):
    try:
        from PIL import Image
    except ImportError:
        return None
    files = sorted(glob.glob(os.path.join(out, "slide-*.png")))
    if not files: return None
    tw, th = 800, 450
    rows = (len(files) + 1) // 2
    sheet = Image.new("RGB", (2 * tw, rows * th), "gray")
    for k, f in enumerate(files):
        im = Image.open(f).convert("RGB").resize((tw, th))
        sheet.paste(im, ((k % 2) * tw, (k // 2) * th))
    p = os.path.join(out, "contact.png"); sheet.save(p); return p

def main():
    a = sys.argv[1:]
    if len(a) < 2: print(__doc__); sys.exit(1)
    deck, out = a[0], a[1]
    only = [int(x) for x in a[a.index("--slides") + 1].split(",")] if "--slides" in a else []
    os.makedirs(out, exist_ok=True)
    for f in glob.glob(os.path.join(out, "slide-*.png")):
        n = int(os.path.basename(f)[6:-4])
        if not only or n in only: os.remove(f)
    ok = via_powerpoint(deck, out, only) or via_libreoffice(deck, out, only)
    if not ok:
        print("No renderer available (PowerPoint COM or LibreOffice+pdftoppm). Visual QA could not run - "
              "tell the user and ask them to open the deck and check it, or to fix the renderer."); sys.exit(2)
    print("\n".join(sorted(glob.glob(os.path.join(out, "slide-*.png")))))
    c = contact(out)
    if c: print("contact sheet:", c)

if __name__ == "__main__":
    main()
