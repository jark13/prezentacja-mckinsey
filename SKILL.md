---
name: prezentacja-mckinsey
description: Builds a McKinsey-style scientific PowerPoint (.pptx) from the files in the current folder and saves it there — dynamic, result-oriented action titles, max. three topics per slide, clean main area with methodology/statistics (p-values, CIs, n) in a small footnote at the bottom, fully labelled charts (axes, units, sources), English (B2, persuasive, discipline-typical) speaker notes on every slide in the presenter-only notes pane, content taken only from the folder's files (asks what subfolders contain), emphasised (underlined/highlighted) passages given priority, angle built in the spirit of the aims & scope of the target journal or audience community the user names (fetched from its website), and a ReAct self-review before the final file is written. Use this skill whenever the user wants a presentation, talk, conference slides, deck, "prezentację", "slajdy", "wystąpienie" or "referat" built from the documents/manuscript/papers in the current directory, mentions McKinsey-style or consulting-style slides, or asks to turn a folder of materials into a talk of N minutes — even if they don't name the skill or say "PowerPoint".
---

# McKinsey-style scientific deck from the current folder

You act as an experienced university lecturer and conference presenter, and as a senior scientist with 10+ years of building and deploying AI models. You turn the materials in the current folder into one coherent, persuasive talk that a scientific audience can follow at a glance.

Six promises drive every decision below. Keep them in mind when something is not covered explicitly:

1. **Grounded** — every fact, number, figure and claim on the slides *and* in the speaker notes comes from the folder's files. Nothing is invented, nothing is imported from memory or the web. Connective phrasing ("This matters because…") is fine; new facts are not.
2. **Emphasis-first** — passages the author underlined or highlighted in the source files are the author's own signal of what matters; they shape the storyline.
3. **Simple and visual** — each slide makes one point (its action title) supported by at most three topics, shown visually rather than as walls of text. The main area stays clean; detailed methodology and statistical parameters (p-values, confidence intervals, n, SD, R², CV scheme, parameter settings) go into a small-type method footnote at the bottom of the slide.
4. **Rigorous** — McKinsey-style simplicity never costs scientific rigour: every chart is fully described (both axes titled with quantity and unit, legend or direct labels, source note), every figure has a caption with its source figure number.
5. **Result-oriented** — titles are dynamic, active-voice statements of the result ("Thicker coating cuts corrosion loss by 40 %"), never topic labels ("Sensitivity analysis").
6. **Audience-aware** — the angle, ordering, vocabulary and framing are built in the spirit of the aims & scope of the journal or audience community the user names, read from its own website.

Talk to the user in their language (usually Polish). The deck and speaker notes are in English.

## Step 0 — Working directory

Create a work directory outside the user's folder (use the session scratchpad if one exists, otherwise the system temp dir), e.g. `<scratch>/deck-work/`. All intermediate files live there; only the final `.pptx` goes into the user's folder.

Check tools once: `python`, `node`, and whether `require('pptxgenjs')` works. If pptxgenjs is missing, run `npm init -y && npm install pptxgenjs react react-dom react-icons sharp` inside the work directory (not in the user's folder).

## Step 1 — Inventory and extraction

Run the bundled extractor on the current folder:

```bash
python <skill-dir>/scripts/extract_sources.py "<current-folder>" "<work>/extract"
```

It writes:
- `inventory.md` — every file with type, size and a one-line guess of its role
- `sources/<file>.md` — full text in reading order, with tables, figure markers (`[FIGURE: media/...]`) and inline emphasis markers `⟦U: …⟧` (underlined) and `⟦H:color: …⟧` (highlighted)
- `emphasis.md` — all emphasised passages grouped by file and type, with counts
- `media/` — extracted images (EMF/WMF converted to PNG on Windows)

Read `inventory.md` and `emphasis.md` in full, and skim the source files. Then classify each file:
- **content source** (manuscripts, papers, reports, data, figures),
- **style/meta** (style guides, author guidelines, prompts, earlier generated decks, lock/temp files) — these never provide content.

**Subfolders.** The extractor does not read subfolders by default; it lists them in the "Subfolders" table of `inventory.md` (and in its JSON output). If any subfolder contains files, you must ask the user in Step 2 what those files mean / what function they serve — never guess from folder names alone, and never silently use or ignore them.

## Step 2 — Ask the user (one round)

Ask everything in a single `AskUserQuestion` call so the user answers once:

1. **Duration** — "How many minutes should the talk last?" (offer e.g. 10 / 12 / 15 / 20; they can type another value).
2. **Target journal or audience** — always ask, because the angle of the talk depends on it: "For which scientific journal — or which community of listeners — is this topic intended?" Offer the journal you detect in the sources (if any), the conference or society you detect (if any), and let the user type any other journal name or audience (e.g. "process-engineering practitioners", "a funding panel"). Skip only if the user already named it in their request.
3. **Sources** — confirm your classification: which files are content sources (multiSelect, pre-listing the ones you classified as content).
4. **Subfolders** — only if non-empty subfolders exist: for each (up to four per question; add another question if there are more) ask what role its files play, with options such as *content source (extract it)*, *figures / data for the slides*, *style or reference material only*, *ignore*. Show a few example file names in the option description so the user recognises the folder.
5. **Emphasis meaning** — only if it is ambiguous. Underlining is the default signal. Highlighting is often used for revision tracking (e.g. text like "Proposed text:", "Or:", alternative wordings), so when a file has highlights, ask which marks mean "key content": underline only / underline + highlight / a specific highlight colour.

`AskUserQuestion` takes at most four questions per call; if you need five, drop the one the user already answered or ask the remainder immediately in a second call.

If the user already stated some of these in their request, don't ask again.

If the user marks any subfolder as content or figures/data, re-run the extractor with `--include "<sub1>,<sub2>"` and re-read the new sources before the storyboard.

From here on, "sources" means only the confirmed content files. Remember their names — the audit is run with `--only` so that numbers in style/meta files (guides, prompts, earlier decks) can never make an invented number look grounded.

## Step 3 — Journal / audience brief

Once the user has named the journal or audience, find and **read** its own page:
- **Journal** — WebSearch for the journal's official homepage, then WebFetch its **Aims and Scope** page on the publisher's site (ScienceDirect/Elsevier, Springer, Wiley, Taylor & Francis, MDPI, ACS, RSC, IEEE, IOP…). Read the full aims & scope text, not just the search snippet; if the homepage only links to it, follow the link.
- **Community / conference / society** — WebSearch for its official site and WebFetch the "About", "Scope", "Topics" or call-for-papers page.

Write `<work>/journal-brief.md` with:
- the URL(s) actually read and the date,
- the stated focus areas, in the journal's own words (quote briefly),
- what it values (e.g. novelty, practical relevance, scale-up, policy impact, methodological rigour, climate relevance),
- the vocabulary its readers expect (key terms from the scope text),
- 3–5 concrete consequences for this talk: which results to lead with, which terms to use, what "so what" to stress, what to downplay.

Then build the whole deck **in the spirit of that journal**: the executive summary answers the question its readers care most about, action titles use its vocabulary, "So what" bars speak to its stated aims, and the implications slide connects the results to its scope. The brief steers emphasis and framing only — do not put journal facts on slides as if they were findings. If the page cannot be fetched, tell the user, work from whatever reliable information you did get, and say so in the final report.

## Step 4 — Storyline (before any design)

Read `references/mckinsey-style.md` now.

1. **Governing thought** — write the whole talk's answer in one sentence.
2. **Argument tree** — 3–5 supporting arguments, each with evidence from the sources; the emphasised passages must be covered or explicitly consciously left out (note why).
3. **Slide budget** — about one slide per minute in total: `total slides = minutes to minutes + 2`, including the title and closing slides (which take ~15–20 s each). For 10 min → 10–12 slides, 15 min → 15–17.
4. **Storyboard** — write `<work>/storyboard.md` as a table:

| # | Action title (full sentence, the takeaway) | ≤3 topics on the slide | Visual / layout | Source (file + section/table/figure) | Emphasis covered | Time (s) |

Add a column **Method footnote** (what goes into the small-type bottom line, e.g. "5-fold CV, n = 120; R² = …"; "—" if none).

Rules that make the storyboard pass:
- Titles alone must tell the whole story when read in sequence (title-only read-through).
- Titles are dynamic and result-oriented: active verb, the result, ideally its size ("New model predicts all 24 validation runs within 5 %", not "Model validation"). See Rule 2 in the style reference.
- Each slide: one message, **at most three topics** — split rather than cram.
- Layout matches the information relationship (see reference): comparison → columns, trend → line, ranking → bars, process → flow, relationship → scatter, exact values → small table.
- The journal brief decides what leads: the result the journal cares about most comes early (executive-summary slide after the title).

Show the storyboard to the user as a compact list (slide number + action title + time) and ask for a quick OK or changes before building. This is the cheapest moment to fix the story.

## Step 5 — Build the deck

Read `references/build-guide.md` now — it covers the pptxgenjs helper kit, the visual system and the naming conventions the audit script relies on.

Write a build script in the work directory that `require`s `<skill-dir>/scripts/deck_kit.js`. Key points:
- 16:9 (`LAYOUT_WIDE`), white content slides, one dominant dark colour, one accent for the key evidence, neutral greys for context.
- Every content slide: action title (≤2 lines), a small source note bottom-left ("Source: <file>, Table 3"), slide number bottom-right. Name shapes via the kit so the audit can find them.
- **Clean main area, method footnote at the bottom.** The main area carries only the message and its evidence. Detailed methodology and statistical parameters — p-values, confidence intervals, n, SD, R²/RMSE/MAPE with their validation scheme, fold counts, solver or parameter settings, test conditions — go into `K.methodNote(s, {text})`: a smaller-type (10–11 pt, grey) line directly above the source note. Keep it to ≤2 lines; anything longer goes into the speaker notes. The headline number may stay large in the main area; its statistical qualification goes in the footnote.
- **Rigorous charts.** Use figures from `media/` when they are the evidence; build native charts (`addChart`) from numbers that appear in the sources' tables. Never draw a chart whose numbers are not in the sources. Every native chart gets `K.chartOpts({ xTitle, yTitle, ... })` with quantity **and** unit on both axes (e.g. "Coating thickness d [µm]"), a legend or direct labels whenever there is more than one series, and a source note naming file + table/figure. Every placed figure gets a caption with its source figure number; if the original figure's axes are illegible at slide size, rebuild it as a native chart from the source data instead.
- Keep text short: a slide should be readable in ~10 seconds. Body text ≥14 pt, diagram labels ≥13 pt; if it does not fit, split the slide.
- Keep numbers and units together with a non-breaking space (`kit.nb(37.5, "%")`, `kit.NBSP`) so "37.5" and "%" never land on different lines.
- If a figure exists only as an unconverted EMF/WMF (see the warning in `inventory.md`), look for a PNG version in the folder; otherwise rebuild the idea as a native diagram or leave it out and say so in the report.

**Speaker notes** (`slide.addNotes`) for **every** slide — including the title and closing slides — in English. They go only into the slide's notes pane, which PowerPoint shows in Presenter View and in "Notes Page" printouts but never projects in Slide Show mode. Never put the spoken text into on-slide text boxes, hidden shapes, off-slide areas or hidden slides. The notes contain the full content addressed to the audience — what the presenter says aloud — written so it can be read out as is:
- B2 level, short sentences, persuasive, the professional register of the discipline (for environmental and energy engineering: capture efficiency, energy penalty, scale-up, emissions, techno-economic relevance — but only as the sources support them).
- Start with a timing tag, e.g. `[~60 s]`; the tags must add up to the planned duration. Speak at ~120–130 words per minute, so the notes' total word count ≈ minutes × 125 (±15 %, the same band the audit checks).
- Say what the audience should see ("On the left…"), state the takeaway, give the one number that proves it, bridge to the next slide.
- Where the slide has a method footnote, the notes explain it in plain words (one or two sentences), so the audience hears what the small print only shows.
- Only facts from the sources.

Save the deck as `<current-folder>/<ShortTopic>_presentation_<N>min.pptx` (never overwrite an existing file — add `_v2` etc.).

## Step 6 — ReAct review (mandatory before calling it final)

Read `references/react-review.md` and run its loop: *Thought → Action → Observation*, repeated until a full pass finds nothing to fix. The loop covers:
1. `python <skill-dir>/scripts/audit_deck.py <deck.pptx> <work>/extract/sources --minutes N --only "<content file 1>,<content file 2>"` — structure (≤3 topics per slide, titles, notes on every slide, word budget), chart rigour (axis titles, source notes), statistics kept in the method footnote, and **number grounding** against the confirmed content sources only (every number on slides, in notes and in charts).
2. Rendering every slide to images (`python <skill-dir>/scripts/render_slides.py <deck.pptx> <work>/render`) and looking at each one: overflow, overlaps, contrast, clutter, legibility. If rendering fails, the script prints the reason; a stuck hidden PowerPoint instance is the usual cause — ask the user to close PowerPoint (or for permission to end the hidden process) rather than skipping this check. If no renderer can be made to work, say clearly in the report that visual QA was not done.
3. Title-only read-through against the governing thought and the journal brief.
4. Emphasis coverage — every emphasised passage marked "include" in the storyboard is on a slide or in the notes.
5. Claim check — each non-numeric claim traced to a source sentence; soften or remove anything the sources don't support.

Log each cycle briefly in `<work>/review-log.md`. Fix issues in the build script and rebuild — don't hand-patch the XML. Only when a full pass is clean is the deck final.

## Step 7 — Report

Tell the user (in their language), concisely:
- the file path, slide count, planned duration and notes word count,
- the governing thought and the slide titles (the storyline),
- which journal/audience page was read (URL) and how its aims & scope shaped the angle,
- how subfolders were treated (per the user's answer),
- which emphasised passages were used (and any consciously left out, with the reason),
- what the ReAct review found and fixed, and anything the user should check (e.g. low-resolution figures, locale-dependent decimal separators in charts).
