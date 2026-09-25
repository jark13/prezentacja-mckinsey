# Build guide — pptxgenjs + deck_kit.js

## Setup
Work directory (outside the user's folder): `npm init -y && npm install pptxgenjs react react-dom react-icons sharp` if `require('pptxgenjs')` fails. Write `build.js` there; it reads images from `<work>/extract/media/` and writes the final `.pptx` into the user's folder.

```js
const pptxgen = require("pptxgenjs");
const kit = require("<skill-dir>/scripts/deck_kit.js");
(async () => {
  const pres = new pptxgen();
  const K = kit.init(pres, { footer: "<Conference or short title>", total: <slide count> });
  // title slide
  K.titleSlide({ kicker: "...", title: "...", subtitle: "...", authors: "...", affiliations: "...",
                 imagePath: "<media png>", imageBox: { x: 8.6, y: 0.9, w: 4.1, h: 4.6 } });
  // content slide
  const s = K.contentSlide({ title: "<Action title = takeaway>", tracker: "<Section>", source: "Source: <file>, <Table/Fig/Section>" });
  const px = await kit.imageSize("<media png>");
  K.imageFit(s, { path: "<media png>", px, x: 0.6, y: 1.8, w: 6.2, h: 4.9, caption: "..." });
  K.bigNumber(s, 1, { x: 7.3, y: 1.9, w: 5.4, value: "−40 %", label: "short label from the source" });
  K.topic(s, 2, { x: 7.3, y: 3.6, w: 5.4, h: 2.2, head: "...", body: ["...", "..."] });
  K.methodNote(s, { text: "<validation scheme, n, p / CI, parameter settings — from the source>" });
  // native chart: axis titles with units are mandatory
  // s.addChart(pres.charts.BAR, data, K.chartOpts({ x: 0.6, y: 1.8, w: 6.2, h: 4.4, xTitle: "Input variable", yTitle: "Share of output variance [%]", barDir: "bar" }));
  s.addNotes("[~60 s]\n...");  // spoken text for the audience — notes pane only, never on the slide
  await pres.writeFile({ fileName: "<folder>/<Topic>_presentation_<N>min.pptx" });
})();
```

## Kit functions
| Function | Use |
|---|---|
| `contentSlide({title, tracker, source})` | white page: tracker label, action title (24 pt, navy), source note, page number |
| `titleSlide({...})`, `closingSlide({title, lines, note})` | dark navy cover / closing |
| `topic(s, n, {x,y,w,h, head, body, fill})` | one topic (n = 1..3); `body` string or array (bullets); `fill:false` for no panel |
| `bigNumber(s, n, {x,y,w, value, label, color, panel, h})` | striking number as a topic; `panel:true` adds a light card (good for executive-summary rows) |
| `soWhat(s, {text})` | implication bar at the bottom (not counted as a topic); with a method note pass `y: 5.6` |
| `methodNote(s, {text, label, lines, size})` | small-type (10.5 pt, grey) methodology/statistics footnote above the source line, ≤2 lines; returns `{y}` = its top edge — keep other content above it. Not a topic, not counted in slide words |
| `imageFit(s, {path, px, x,y,w,h, caption})` | figure scaled inside a box, aspect preserved |
| `chartOpts({xTitle, yTitle, ...})` | quiet chart defaults (accent + grey, no legend, light gridlines) **plus mandatory axis titles** — throws if `xTitle`/`yTitle` are missing |
| `flow(s, steps, {x,y,w,h})` | left-to-right process (method pipeline) |
| `K.C`, `K.F`, `K.W`, `K.H`, `K.M` | palette, fonts, slide size, margin |
| `kit.nb(value, unit)`, `kit.NBSP` | number + non-breaking space + unit |
| `kit.imageSize(path)` | pixel size for `imageFit` (reads via buffer, safe for long Windows paths) |

Every separate topic must go through `topic()` or `bigNumber()` — the audit counts `Topic-N-*` / `BigNumber-N-*` names to enforce the three-topic limit. Evidence (images, charts) and the `soWhat` bar are not topics.

## Visual system (McKinsey-like, restrained)
- White content pages; navy for titles and dark cover/closing; ONE accent (blue) for the key evidence; grey for context; `alert` red only for losses/risks.
- Arial throughout (renders identically in PowerPoint and in QA). Title 24 pt bold, topic header 17 pt, body 15 pt (≥14), diagram labels ≥13 pt, chart axis titles 12 pt, method footnote 10–11 pt, captions/sources 9–10 pt.
- Vertical zones: title 0.55–1.6" · content 1.8–~6.4" · method footnote ~6.5–7.0" · source + page number 7.08".
- Numbers and units: join with a non-breaking space (`kit.nb(0.25, "m")`) so they never split across lines; the same for "R² = 0.92"-style expressions in titles.
- Grid: 0.6" margins; content starts at y ≈ 1.8"; source line at y ≈ 7.08". Keep ≥0.3" between blocks. Align left edges.
- **Panels hug their content.** A topic panel should end ~0.25" below its last line: height ≈ 0.77" (padding + header) + 0.25" per body line at 15 pt + 0.08" per extra bullet + 0.34" (padding). Half-empty panels look unfinished. Use freed space for evidence (a figure, a chart, a big number) or a `soWhat` bar — or enlarge the text (up to 18 pt) if the slide is light. Pass `hAuto: true` to `topic()` to size the panel from the text.
- No decorative lines under titles, no edge stripes, no shadows, no clip-art. Icons only if they carry meaning.
- Layouts vary by information type (see mckinsey-style.md, Rule 4); don't repeat one layout for every slide.

## Charts (native, editable)
- Only numbers that appear in the sources (tables, text).
- Always `K.chartOpts({ xTitle, yTitle, ... })`: both axes titled with quantity + unit (for `barDir: "bar"` the category axis is the vertical one — titles still map cat→`xTitle`, val→`yTitle`). The audit fails charts whose axes lack titles.
- More than one series → `showLegend: true, legendPos: "b"` or direct data labels; bars start at zero.
- Every slide with a chart or figure needs a `source` in `contentSlide()` (file + table/figure).
- Bar for categories/ranking, line for trends, scatter for relationships. Highlight the key bar with the accent, others grey (two series with zeros + `dataLabelFormatCode: '0.0;;;'` hides the zero labels on stacked charts).
- Stacked charts: `dataLabelPosition` must be `ctr`, `inEnd` or `inBase` (`outEnd` corrupts the file).
- Colours: no `#`, 6-digit hex only. Don't reuse one options object across calls (pptxgenjs mutates it).
- Decimal separators follow the presenter's Office locale — mention it in the report.

## Figures
Use `media/*.png` (EMF/WMF were converted to `*_emf.png` / `*_wmf.png`). Look at each candidate image before using it. Prefer the highest-resolution copy (an image file in the folder may be a better version of a figure embedded in the manuscript). Crop nothing that changes meaning; caption with the source figure number.

## Speaker notes
Plain text via `slide.addNotes()` on every slide (title and closing included). This fills the notes pane: visible to the presenter in Presenter View, never shown to the audience in Slide Show. Never simulate notes with on-slide text boxes, off-canvas shapes or hidden slides. Structure per slide: `[~NN s]` → what to look at → the takeaway (mirrors the action title) → the proving number → bridge to the next slide. B2 English: short sentences, common vocabulary plus the discipline's technical terms, active voice, persuasive but not exaggerated ("The data show…", "This means…", "So the key message is…"). Budget ≈ 125 words per minute of the slide's time.

## File hygiene
- Save as `<Topic>_presentation_<N>min.pptx` in the user's folder; if it exists, add `_v2`, `_v3`.
- Keep `build.js`, `storyboard.md`, `journal-brief.md` and `review-log.md` in the work directory so the deck can be rebuilt quickly after feedback.
