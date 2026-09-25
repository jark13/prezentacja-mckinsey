# ReAct review — the gate before the deck is final

ReAct = reason about what to check (**Thought**), do one concrete check with a tool (**Action**), read the result honestly (**Observation**), then fix or move on. Repeat until one full pass over all checks produces no fixes. Log each cycle in `<work>/review-log.md` in this form:

```
### Cycle 1
Thought: Numbers on slides must all come from the sources.
Action: audit_deck.py deck.pptx extract/sources --minutes 12
Observation: slide 7 "0.087" OK; slide 9 FAIL "2.5" not in sources (I rounded 2.6 wrongly).
Fix: corrected to 2.6 in build.js; rebuilt.
```

Why this matters: the deck will be presented to experts under the author's name. One invented number, a claim the data don't support, or a slide nobody can read from the back row costs more credibility than the whole talk can win back.

## The checks (run all of them in every full pass)

1. **Structure & grounding (script)**
   `python <skill-dir>/scripts/audit_deck.py <deck> <work>/extract/sources --minutes N --only "<content files>"`
   - every FAIL must be fixed: missing title/notes, >3 topics, numbers not found in the sources, chart axes without titles, chart/figure without a source note;
   - judge every WARN: long titles, >70 words on a slide, statistics in the main area (move to `methodNote`), method note >2 lines, notes budget off by >15 %, slide count off.
   Note: single-digit integers and multipliers ("3×", "5×") are not checked automatically — verify them by hand; a derived ratio (e.g. 0.60/0.12 = 5) is only allowed if the source states it.
   A number flagged as ungrounded is either a typo, a wrong rounding, a derived value (then show the source numbers or drop it), or invented (remove it).

2. **Claims (reasoning)** — for every sentence on the slides and every factual sentence in the notes, find the supporting sentence/table in `extract/sources`. If the source hedges ("hypothesis", "expected", "preliminary"), the slide must hedge too. Remove any fact that comes from memory, the web or the journal page.

3. **Storyline (titles only)** — list the action titles in order. Do they read as one argument from problem to conclusion, consistent with the governing thought? Does the executive summary match the final takeaways? Is the order the one the journal brief suggests? Is every title dynamic and result-oriented (active verb, the result, a number where the source has one) — rewrite any label-style title ("Results", "Methodology", "Sensitivity analysis")? Do titles and "So what" bars use the journal's scope vocabulary?

3a. **Rigour** — for every chart: both axes titled with quantity + unit, legend/direct labels for >1 series, axis range honest, source note present. For every figure: caption with source figure number, axis labels legible at slide size. For every slide: detailed methodology and statistics sit in the small method footnote, not in the main area, and the notes explain them.

4. **Emphasis coverage** — compare `emphasis.md` (the marks the user said mean "key content") with the storyboard. Each emphasised item is on a slide or in the notes, or is listed as consciously left out with a reason.

5. **Visual (render and look)** — `python <skill-dir>/scripts/render_slides.py <deck> <work>/render` then open every `slide-NN.png` (not only the contact sheet). Look for: text overflow or cut-off, overlaps, elements closer than ~0.3", misalignment, low contrast, text < 12 pt, crowded slides, figures too small or pixelated, empty regions that look unfinished (especially panels whose lower half is empty — shrink them or give the space to evidence), leftover placeholder text. Ask: could someone state this slide's conclusion in 5 seconds?

6. **Language** — English B2, persuasive, discipline-typical terminology; consistent units and symbols (subscripts written the same way everywhere); no Polish leftovers; notes read naturally aloud; timing tags sum to the target duration (±10 %). Every slide (title and closing included) has notes in the notes pane; no spoken text sits on the slide itself.

7. **File** — deck opens (rendering succeeded), saved in the user's folder under the agreed name, nothing else in the folder was modified.

The method footnote (10–11 pt) and source line (9–10 pt) are the only deliberate exceptions to the "text < 12 pt" rule in check 5.

## Stop rule
Final only after a pass where checks 1–7 (incl. 3a) produce no changes. If a problem cannot be fixed from the sources (e.g. a figure exists only in low resolution), keep it and tell the user in the report.
