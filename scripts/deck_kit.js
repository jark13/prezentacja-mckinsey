// deck_kit.js — McKinsey-style helper kit for pptxgenjs (scientific talks).
// Usage (from a build script in the work directory, where pptxgenjs is installed):
//   const pptxgen = require("pptxgenjs");
//   const kit = require("<skill-dir>/scripts/deck_kit.js");
//   const pres = new pptxgen();
//   const K = kit.init(pres, { footer: "Conference · City 2026", total: 14 });
//   const s = K.contentSlide({ title: "Action title that states the takeaway", tracker: "Results", source: "Source: Manuscript.docx, Table 3" });
//   K.topic(s, 1, { x: 0.6, y: 1.9, w: 3.9, h: 2.2, head: "Short header", body: "One or two short sentences." });
//   s.addNotes("[~60 s] ...");
//   await pres.writeFile({ fileName: out });
//
// Shape names written here (ActionTitle, Topic-N-*, BigNumber-N-*, SoWhat, MethodNote, SourceNote, PageNum, Tracker)
// are read by audit_deck.py — create every separate topic through topic()/bigNumber() so the
// "max three topics per slide" check can count them.

const path = require("path");

const PALETTE = {
  navy: "051C2C",      // dominant: titles, dark slides, primary shapes
  accent: "2251FF",    // ONE accent: the key evidence on each slide
  alert: "B42318",     // sparingly: losses, risks, negative deltas
  grey: "5F6B76",      // secondary text, captions, context series
  light: "EEF2F6",     // card / panel tint
  line: "D5DCE3",      // hairlines, gridlines
  ink: "111827",       // body text
  white: "FFFFFF",
};
const FONT = { head: "Arial", body: "Arial" };
const W = 13.333, H = 7.5, M = 0.6; // LAYOUT_WIDE, 0.6" margins

function init(pres, opts = {}) {
  pres.layout = "LAYOUT_WIDE";
  const C = Object.assign({}, PALETTE, opts.palette || {});
  const F = Object.assign({}, FONT, opts.fonts || {});
  let page = 0;
  const total = opts.total || null;
  const footer = opts.footer || "";

  const txt = (s, text, o) => s.addText(text, Object.assign({ fontFace: F.body, color: C.ink, margin: 0, isTextBox: true, valign: "top" }, o));

  function pageNum(s, dark) {
    page += 1;
    const label = (footer ? footer + "   |   " : "") + (total ? `${page} / ${total}` : `${page}`);
    txt(s, label, { objectName: "PageNum", x: W - M - 5, y: H - 0.42, w: 5, h: 0.28, fontSize: 9, color: dark ? "9AA8B5" : C.grey, align: "right" });
  }

  function contentSlide({ title, tracker, source, background }) {
    const s = pres.addSlide();
    s.background = { color: background || C.white };
    if (tracker) txt(s, tracker.toUpperCase(), { objectName: "Tracker", x: M, y: 0.28, w: 8, h: 0.26, fontSize: 10, bold: true, charSpacing: 2, color: C.accent });
    txt(s, title, { objectName: "ActionTitle", x: M, y: 0.55, w: W - 2 * M, h: 1.05, fontFace: F.head, fontSize: 24, bold: true, color: C.navy, valign: "middle" });
    if (source) txt(s, source, { objectName: "SourceNote", x: M, y: H - 0.42, w: 7.2, h: 0.28, fontSize: 9, italic: true, color: C.grey });
    pageNum(s, false);
    return s;
  }

  function darkSlide() {
    const s = pres.addSlide();
    s.background = { color: C.navy };
    return s;
  }

  function titleSlide({ kicker, title, subtitle, authors, affiliations, imagePath, imageBox }) {
    const s = darkSlide();
    if (kicker) txt(s, kicker, { objectName: "Kicker", x: M, y: 0.6, w: 7.6, h: 0.35, fontSize: 14, bold: true, color: "7FA7FF" });
    txt(s, title, { objectName: "ActionTitle", x: M, y: 1.15, w: imagePath ? 7.6 : W - 2 * M, h: 2.3, fontFace: F.head, fontSize: 32, bold: true, color: C.white });
    if (subtitle) txt(s, subtitle, { objectName: "Subtitle", x: M, y: 3.5, w: imagePath ? 7.6 : W - 2 * M, h: 0.55, fontSize: 18, color: "C9D6E3" });
    if (authors) txt(s, authors, { objectName: "Authors", x: M, y: 4.35, w: imagePath ? 7.6 : W - 2 * M, h: 0.9, fontSize: 13, color: C.white });
    if (affiliations) txt(s, affiliations, { objectName: "Affiliations", x: M, y: 5.3, w: imagePath ? 7.6 : W - 2 * M, h: 0.9, fontSize: 11, color: "9AA8B5" });
    if (imagePath) s.addImage(Object.assign({ path: imagePath }, imageBox || { x: 8.6, y: 0.9, w: 4.1, h: 4.6 }));
    page += 1;
    return s;
  }

  function closingSlide({ title, lines = [], note }) {
    const s = darkSlide();
    // the title may wrap to two lines: tall box, bottom-aligned, contact lines start below it
    const t0 = title || "Thank you";
    txt(s, t0, { objectName: "ActionTitle", x: M, y: 0.9, w: W - 2 * M, h: 2.0, fontFace: F.head, fontSize: t0.length > 40 ? 30 : 44, bold: true, color: C.white, valign: "bottom" });
    if (lines.length) txt(s, lines.map((l, i) => ({ text: l, options: { breakLine: i < lines.length - 1 } })), { objectName: "Contact", x: M, y: 3.3, w: 8, h: 1.8, fontSize: 16, color: "C9D6E3" });
    if (note) txt(s, note, { objectName: "Acknowledgement", x: M, y: 5.6, w: W - 2 * M, h: 1.1, fontSize: 10, color: "9AA8B5" });
    page += 1;
    return s;
  }

  // Estimated panel height for a topic body (Arial ~ w_in*144/size chars per line). Use the max over a
  // row of panels so side-by-side panels share one height: h = Math.max(...bodies.map(b => K.autoH(w, b)))
  function autoH(w, body, size = 15) {
    const items = Array.isArray(body) ? body : [body];
    const bullets = items.length > 1 ? 0.2 : 0;                        // bullet indent narrows the line
    const perLine = Math.max(10, Math.floor((w - 0.44 - bullets) * 144 / size));
    const lines = items.reduce((a, t) => a + Math.ceil(t.length / perLine), 0);
    // pad 0.22 + header 0.55 + lines at 1.2 line spacing + 6 pt paragraph spacing + pad 0.22 + slack
    return 0.22 + 0.55 + lines * size * 1.2 / 72 + (items.length - 1) * 6 / 72 + 0.22 + 0.12;
  }

  // One topic = header + short body. n = 1..3 (the audit counts distinct n per slide).
  function topic(s, n, { x, y, w, h, head, body, fill, headColor, size = 15, hAuto = false }) {
    if (hAuto && body) h = Math.min(h || 99, autoH(w, body, size));
    if (fill !== false) s.addShape(pres.shapes.RECTANGLE, { objectName: `Topic-${n}-panel`, x, y, w, h, fill: { color: fill || C.light }, line: { color: fill || C.light } });
    const pad = fill === false ? 0 : 0.22;
    txt(s, head, { objectName: `Topic-${n}-head`, x: x + pad, y: y + pad, w: w - 2 * pad, h: 0.5, fontSize: size + 2, bold: true, color: headColor || C.navy });
    if (body) {
      const items = Array.isArray(body) ? body : [body];
      const runs = items.map((t, i) => ({ text: t, options: Object.assign({ breakLine: i < items.length - 1, paraSpaceAfter: 6 }, items.length > 1 ? { bullet: { indent: 14 } } : {}) }));
      txt(s, runs, { objectName: `Topic-${n}-body`, x: x + pad, y: y + pad + 0.55, w: w - 2 * pad, h: h - 2 * pad - 0.55, fontSize: size, color: C.ink });
    }
    return { x, y, w, h };  // actual box (useful with hAuto to place the next element)
  }

  // Big-number callout as a topic (counts toward the three-topic limit).
  function bigNumber(s, n, { x, y, w, value, label, color, valueSize = 44, panel = false, h }) {
    if (panel) {
      const ph = h || valueSize / 60 + 1.35;
      s.addShape(pres.shapes.RECTANGLE, { objectName: `BigNumber-${n}-panel`, x: x - 0.2, y: y - 0.2, w: w + 0.4, h: ph, fill: { color: C.light }, line: { color: C.light } });
    }
    txt(s, value, { objectName: `BigNumber-${n}-value`, x, y, w, h: valueSize / 60 + 0.3, fontFace: F.head, fontSize: valueSize, bold: true, color: color || C.accent });
    txt(s, label, { objectName: `BigNumber-${n}-label`, x, y: y + valueSize / 60 + 0.32, w, h: 0.75, fontSize: 13, color: C.grey });
  }

  // "So what" box: the implication of the evidence (not counted as a topic).
  function soWhat(s, { x = M, y = 6.05, w = W - 2 * M, h = 0.75, text, label = "So what: " }) {
    s.addShape(pres.shapes.RECTANGLE, { objectName: "SoWhat-panel", x, y, w, h, fill: { color: C.navy }, line: { color: C.navy } });
    txt(s, [{ text: label, options: { bold: true, color: "7FA7FF" } }, { text, options: { color: C.white } }],
      { objectName: "SoWhat", x: x + 0.25, y, w: w - 0.5, h, fontSize: 15, valign: "middle" });
  }

  // Method / statistics footnote: detailed methodology and statistical parameters (n, p-values,
  // confidence intervals, R², RMSE, CV scheme, parameter settings) in small type at the bottom of the
  // slide, directly above the source line — keeps the main area clean. Not counted as a topic.
  // Occupies y ≈ 6.52–7.02"; content above it (incl. soWhat — pass y: 5.6) must end by ~6.45".
  function methodNote(s, { text, label = "Method: ", x = M, w = W - 2 * M, size = 10.5, lines = 2 }) {
    const h = lines * size * 0.0195 + 0.08;
    const y = H - 0.42 - 0.06 - h;
    s.addShape(pres.shapes.LINE, { objectName: "MethodNote-rule", x, y: y - 0.06, w, h: 0, line: { color: C.line, width: 0.75 } });
    txt(s, [{ text: label, options: { bold: true, color: C.grey } }, { text, options: { color: C.grey } }],
      { objectName: "MethodNote", x, y, w, h, fontSize: size, valign: "bottom" });
    return { y: y - 0.06 };  // top edge: keep other content above this
  }

  // Image scaled to fit (aspect preserved) inside a box; needs pixel size (use imageSize()).
  function imageFit(s, { path: p, px, x, y, w, h, caption, name }) {
    const r = px.width / px.height;
    let iw = w, ih = w / r;
    if (ih > h) { ih = h; iw = h * r; }
    const ix = x + (w - iw) / 2, iy = y + (h - ih) / 2;
    s.addImage({ objectName: name || "Evidence-image", path: p, x: ix, y: iy, w: iw, h: ih });
    if (caption) txt(s, caption, { objectName: "Caption", x, y: iy + ih + 0.05, w, h: 0.3, fontSize: 10, italic: true, color: C.grey, align: "center" });
    return { x: ix, y: iy, w: iw, h: ih };
  }

  // Quiet, consulting-style chart defaults; merge with your own options.
  // Scientific rigour: xTitle and yTitle are REQUIRED and must carry the quantity AND unit,
  // e.g. xTitle: "Flow velocity u [m/s]", yTitle: "Pressure drop Δp [kPa]".
  // (For horizontal bar charts the category axis is vertical — still pass both.) The audit fails
  // charts whose axes have no titles. Use a legend (showLegend) whenever >1 series is not labelled directly.
  function chartOpts(o = {}) {
    const { xTitle, yTitle, ...rest } = o;
    if (!xTitle || !yTitle) throw new Error("chartOpts: xTitle and yTitle (quantity + unit) are required");
    return Object.assign({
      chartColors: [C.accent, "A7B4C2"],
      showTitle: !!o.title, titleFontSize: 13, titleColor: C.ink, titleFontFace: F.body,
      showValue: true, dataLabelFontSize: 11, dataLabelColor: C.ink,
      catAxisLabelColor: C.grey, valAxisLabelColor: C.grey, catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
      catAxisLabelFontFace: F.body, valAxisLabelFontFace: F.body,
      showCatAxisTitle: true, catAxisTitle: xTitle, catAxisTitleFontSize: 12, catAxisTitleColor: C.ink, catAxisTitleFontFace: F.body,
      showValAxisTitle: true, valAxisTitle: yTitle, valAxisTitleFontSize: 12, valAxisTitleColor: C.ink, valAxisTitleFontFace: F.body,
      valGridLine: { color: "E6EAEE", size: 0.5 }, catGridLine: { style: "none" },
      showLegend: false, legendFontSize: 11, legendFontFace: F.body,
    }, rest);
  }

  // Left-to-right process flow (method pipelines). steps: [{head, body}]
  function flow(s, steps, { x = M, y = 2.1, w = W - 2 * M, h = 2.2, n0 = 1 } = {}) {
    const gap = 0.45, bw = (w - gap * (steps.length - 1)) / steps.length;
    steps.forEach((st, i) => {
      const bx = x + i * (bw + gap);
      s.addShape(pres.shapes.RECTANGLE, { objectName: `Flow-${i + 1}`, x: bx, y, w: bw, h, fill: { color: i === steps.length - 1 ? C.navy : C.light }, line: { color: i === steps.length - 1 ? C.navy : C.light } });
      txt(s, st.head, { objectName: `Flow-${i + 1}-head`, x: bx + 0.18, y: y + 0.18, w: bw - 0.36, h: 0.5, fontSize: 15, bold: true, color: i === steps.length - 1 ? C.white : C.navy });
      if (st.body) txt(s, st.body, { objectName: `Flow-${i + 1}-body`, x: bx + 0.18, y: y + 0.72, w: bw - 0.36, h: h - 0.9, fontSize: 13, color: i === steps.length - 1 ? "E5ECF3" : C.ink });
      if (i < steps.length - 1) s.addShape(pres.shapes.RIGHT_ARROW, { x: bx + bw + 0.08, y: y + h / 2 - 0.16, w: gap - 0.16, h: 0.32, fill: { color: C.grey }, line: { color: C.grey } });
    });
  }

  return { C, F, W, H, M, autoH, contentSlide, titleSlide, closingSlide, topic, bigNumber, soWhat, methodNote, imageFit, chartOpts, flow, txt };
}

// Pixel size of an image (uses sharp, installed with the work-directory dependencies).
async function imageSize(p) {
  const sharp = require(require.resolve("sharp", { paths: [process.cwd()] }));
  // read into a buffer first: sharp's path loader fails on Windows paths > 260 chars
  const m = await sharp(require("fs").readFileSync(p)).metadata();
  return { width: m.width, height: m.height };
}

// Non-breaking space keeps numbers and units on one line: nb(37.5, "%") -> "37.5\u00A0%".
const NBSP = "\u00A0";
const nb = (value, unit) => `${value}${NBSP}${unit}`;

module.exports = { init, imageSize, PALETTE, FONT, NBSP, nb };
