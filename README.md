# prezentacja-mckinsey — Claude Code skill

Builds a McKinsey-style scientific PowerPoint (`.pptx`) from the files in the folder you are working in: conclusion-led action titles, at most three topics per slide, methodology and statistics in a small footnote, fully labelled charts, English speaker notes on every slide (notes pane only), framing tuned to the aims & scope of the target journal or audience, and a ReAct self-review (number grounding against the sources + rendered visual QA) before the file is saved.

## Installation

Clone the repository into your Claude Code skills directory:

```bash
# for all projects (user level)
git clone https://github.com/jark13/prezentacja-mckinsey.git ~/.claude/skills/prezentacja-mckinsey

# or for one project only
git clone https://github.com/jark13/prezentacja-mckinsey.git <project>/.claude/skills/prezentacja-mckinsey
```

Restart Claude Code (or run `/reload-plugins`) and the skill appears in `/skills`.

## Requirements

- Python 3 (standard library; optional: `Pillow`, `pypdf`/`PyMuPDF`, `openpyxl` for richer extraction)
- Node.js — the skill installs `pptxgenjs` and `sharp` itself in a temporary work directory, never in your folder
- Slide rendering for visual QA: Microsoft PowerPoint (Windows, COM) or LibreOffice + `pdftoppm`

## Usage

1. `cd` into the folder that holds your material (manuscript, figures, data, author guidelines…).
2. Start Claude Code and ask, e.g. *"Zbuduj prezentację z plików w tym folderze"* or *"Build a 12-minute talk from these files"*.
3. The skill asks once for: duration, target journal or audience, which files are content sources, the meaning of any subfolders, and how to read highlights.
4. It reads the journal's aims & scope, shows you a storyboard (slide titles + timing) for approval, builds the deck, audits and renders it, and saves `<Topic>_presentation_<N>min.pptx` in your folder (never overwriting an existing file).

All intermediate files (extracted text, storyboard, journal brief, build script, review log) stay in a temporary work directory outside your folder.

## Contents

| Path | Purpose |
|---|---|
| `SKILL.md` | the workflow Claude follows |
| `references/mckinsey-style.md` | slide-writing rules (action titles, one message per slide, rigour) |
| `references/build-guide.md` | pptxgenjs helper kit, visual system, charts, speaker notes |
| `references/react-review.md` | the review loop that must pass before the deck is final |
| `scripts/extract_sources.py` | text, tables, figures and emphasis extraction (docx, pptx, pdf, xlsx, text, images) |
| `scripts/deck_kit.js` | slide/topic/chart/method-footnote helpers for pptxgenjs |
| `scripts/audit_deck.py` | structure, notes budget, chart axis titles, source notes and number grounding |
| `scripts/render_slides.py` | renders every slide to PNG for visual QA |
| `evals/evals.json` | test prompts for the skill |
