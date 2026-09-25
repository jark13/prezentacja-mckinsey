# prezentacja-mckinsey — Claude Code skill

Builds a McKinsey-style scientific PowerPoint (`.pptx`) from the files in the folder you are working in: conclusion-led action titles, at most three topics per slide, methodology and statistics in a small footnote, fully labelled charts, English speaker notes on every slide (notes pane only), framing tuned to the aims & scope of the target journal or audience, and a ReAct self-review (number grounding against the sources + rendered visual QA) before the file is saved.

> **Szybki start (PL):** zainstaluj Claude Code, sklonuj to repo do `~/.claude/skills/prezentacja-mckinsey`, uruchom ponownie Claude Code, przejdź do folderu ze swoimi materiałami (manuskrypt, rysunki, dane), wpisz `claude`, a potem np. *„Zbuduj prezentację z plików w tym folderze”*. Skill zada kilka pytań, pokaże storyboard do akceptacji i zapisze gotowy plik `.pptx` w Twoim folderze.

## Prerequisites

- [Claude Code](https://docs.claude.com/en/docs/claude-code) with an active Claude account
- Git
- Python 3 (standard library; optional: `Pillow`, `pypdf`/`PyMuPDF`, `openpyxl` for richer extraction)
- Node.js — the skill installs `pptxgenjs` and `sharp` itself in a temporary work directory, never in your folder
- For visual QA of the slides: Microsoft PowerPoint (Windows) or LibreOffice + `pdftoppm`. Without a renderer the deck is still built, but the report states that visual QA was not done.

## Installation

Clone the repository into your Claude Code skills directory.

**For all projects (user level)** — macOS / Linux / Git Bash:

```bash
git clone https://github.com/jark13/prezentacja-mckinsey.git ~/.claude/skills/prezentacja-mckinsey
```

Windows PowerShell:

```powershell
git clone https://github.com/jark13/prezentacja-mckinsey.git "$HOME\.claude\skills\prezentacja-mckinsey"
```

**For one project only:**

```bash
git clone https://github.com/jark13/prezentacja-mckinsey.git <project>/.claude/skills/prezentacja-mckinsey
```

Then restart Claude Code (or run `/reload-plugins`). Type `/skills` — `prezentacja-mckinsey` should be on the list.

While the repository is private, the owner has to add you as a collaborator first (GitHub → Settings → Collaborators).

## Usage

1. Put your material in one folder: manuscript, figures, data tables, the target journal's author guidelines, etc. Subfolders are fine — the skill will ask what they contain.
2. Open a terminal in that folder and start Claude Code: `claude`.
3. Ask for a talk, in Polish or English, e.g. *"Zbuduj prezentację z plików w tym folderze"* or *"Build a 12-minute talk from these files"*. You can also invoke it explicitly with `/prezentacja-mckinsey`.
4. Answer one round of questions: talk duration, target journal or audience, which files are content sources, what the subfolders contain, and how to read highlighted text.
5. The skill reads the journal's or audience's aims & scope on its website and shows you a storyboard — slide titles and timing. Approve it or ask for changes; this is the cheapest moment to fix the story.
6. It builds the deck, checks every number against your files, renders and inspects every slide, fixes what it finds, and saves `<Topic>_presentation_<N>min.pptx` in your folder. An existing file is never overwritten (`_v2`, `_v3`… are added).
7. You get a short report: storyline, how the journal shaped the angle, what the review fixed, and what you should check yourself.

The speaker notes (what to say on each slide) are in the notes pane: visible in PowerPoint's Presenter View and in "Notes Pages" printouts, never on the projected slides.

## Privacy

- Content comes only from the files in your folder; nothing is invented or taken from the web. The only web access is reading the target journal's or audience's public aims & scope page.
- Your files are not uploaded to this repository or anywhere else by the skill. All intermediate files (extracted text, storyboard, journal brief, build script, review log) stay in a temporary work directory outside your folder; only the final `.pptx` is written into it.
- As with any Claude Code session, the content Claude reads is processed by Claude under your Claude account's terms.

## Updating and uninstalling

```bash
# update to the latest version
git -C ~/.claude/skills/prezentacja-mckinsey pull

# uninstall
rm -rf ~/.claude/skills/prezentacja-mckinsey
```

On Windows PowerShell use `"$HOME\.claude\skills\prezentacja-mckinsey"` as the path (`Remove-Item -Recurse -Force <path>` to uninstall). Restart Claude Code afterwards.

## Troubleshooting

| Symptom | What to do |
|---|---|
| Skill not listed in `/skills` | Check the folder is exactly `~/.claude/skills/prezentacja-mckinsey` and contains `SKILL.md`; restart Claude Code. |
| Slide rendering fails on Windows | A hidden PowerPoint instance is usually stuck — close PowerPoint (Claude will ask before ending the process). |
| Decimal commas in charts, dots in text | Chart number format follows your Office regional settings; switch Office to English (or accept it). |
| A figure is missing from the deck | EMF/WMF images that could not be converted are reported; add a PNG version of the figure to the folder. |

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
