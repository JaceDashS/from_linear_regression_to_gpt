import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";

const C = {
  ink: "#0B0F19",
  muted: "#667085",
  line: "#CBD5E1",
  panel: "#F8FAFC",
  blue: "#2563EB",
  green: "#10B981",
  orange: "#F97316",
  red: "#EF4444",
  white: "#FFFFFF",
};

function addText(slide, text, pos, style = {}) {
  const s = slide.shapes.add({
    geometry: "textbox",
    position: pos,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  s.text = text;
  s.text.style = {
    fontSize: style.fontSize ?? 24,
    bold: style.bold ?? false,
    color: style.color ?? C.ink,
    alignment: style.alignment ?? "left",
  };
  return s;
}

function box(slide, text, pos, opts = {}) {
  const geometry = opts.geometry ?? "roundRect";
  const config = {
    geometry,
    position: pos,
    fill: opts.fill ?? C.white,
    line: { style: "solid", fill: opts.line ?? C.line, width: opts.lineWidth ?? 1.5 },
  };
  if (["rect", "textbox", "roundRect"].includes(geometry)) {
    config.borderRadius = opts.radius ?? "rounded-xl";
  }
  const s = slide.shapes.add(config);
  if (text) {
    s.text = text;
    s.text.style = {
      fontSize: opts.fontSize ?? 22,
      bold: opts.bold ?? true,
      color: opts.color ?? C.ink,
      alignment: opts.alignment ?? "center",
    };
  }
  return s;
}

function circle(slide, x, y, r, color, alphaFill = null) {
  return box(slide, "", { left: x - r, top: y - r, width: r * 2, height: r * 2 }, {
    geometry: "ellipse",
    fill: alphaFill ?? color,
    line: color,
    lineWidth: 1.4,
  });
}

function line(slide, x1, y1, x2, y2, opts = {}) {
  return slide.shapes.add({
    geometry: "line",
    position: { left: Math.min(x1, x2), top: Math.min(y1, y2), width: Math.abs(x2 - x1) || 1, height: Math.abs(y2 - y1) || 1 },
    fill: "none",
    line: { style: opts.style ?? "solid", fill: opts.color ?? C.line, width: opts.width ?? 2 },
  });
}

function title(slide, heading, kicker) {
  addText(slide, kicker, { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, heading, { left: 64, top: 62, width: 1030, height: 62 }, { fontSize: 42, bold: true });
}

function frame(slide) {
  box(slide, "", { left: 64, top: 142, width: 1152, height: 500 }, { fill: C.panel, line: "#E2E8F0", radius: "rounded-lg" });
}

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function notes(slide) {
  slide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard machine learning concepts.");
  slide.speakerNotes.setVisible(true);
}

function chartMapper() {
  const chart = { left: 200, top: 252, width: 760, height: 318 };
  const sx = x => chart.left + ((x - 50) / 50) * chart.width;
  const sy = y => chart.top + chart.height - ((y - 45) / 55) * chart.height;
  return { chart, sx, sy };
}

const data = [
  [52, 61], [55, 48], [57, 67], [61, 58], [63, 73], [66, 54],
  [69, 76], [72, 70], [76, 82], [78, 65], [81, 88], [84, 77],
  [87, 94], [90, 83], [93, 90], [96, 72], [98, 97],
];

function axes(slide, chart, sx, sy) {
  line(slide, chart.left, chart.top + chart.height, chart.left + chart.width, chart.top + chart.height, { color: "#94A3B8", width: 2 });
  line(slide, chart.left, chart.top + chart.height, chart.left, chart.top, { color: "#94A3B8", width: 2 });
  addText(slide, "attendance (%)", { left: 500, top: 590, width: 210, height: 28 }, { fontSize: 20, color: C.muted, alignment: "center" });
  addText(slide, "grade", { left: 128, top: 385, width: 70, height: 28 }, { fontSize: 20, color: C.muted, alignment: "center" });
  for (const [x, y] of data) circle(slide, sx(x), sy(y), 8.2, C.blue);
}

function dottedCurve(slide, sx, sy, fn, color, xStart = 52, xEnd = 99, dots = 66, r = 4.1) {
  for (let i = 0; i <= dots; i++) {
    const x = xStart + ((xEnd - xStart) * i) / dots;
    circle(slide, sx(x), sy(fn(x)), r, color);
  }
}

function addNonlinearSlide(presentation, after) {
  const slide = presentation.slides.insert({ after }).slide;
  slide.background.fill = C.white;
  title(slide, "Nonlinear regression allows curves", "Regression");
  frame(slide);
  const { chart, sx, sy } = chartMapper();
  axes(slide, chart, sx, sy);
  dottedCurve(slide, sx, sy, x => 0.67 * x + 25, C.orange, 52, 99, 54, 3.6);
  dottedCurve(slide, sx, sy, x => 48 + 47 / (1 + Math.exp(-(x - 72) / 7)), C.green, 52, 99, 68, 4);
  box(slide, "linear", { left: 995, top: 315, width: 150, height: 54 }, { fill: "#FEF3C7", line: C.orange, color: C.orange, fontSize: 23 });
  box(slide, "nonlinear", { left: 995, top: 390, width: 150, height: 54 }, { fill: "#DCFCE7", line: C.green, color: C.green, fontSize: 23 });
  addText(slide, "same task, more flexible shape", { left: 355, top: 180, width: 570, height: 36 }, { fontSize: 28, bold: true, color: C.green, alignment: "center" });
  addText(slide, "y_hat = f(x)", { left: 980, top: 485, width: 180, height: 40 }, { fontSize: 30, bold: true, color: C.green, alignment: "center" });
  footer(slide, 6);
  notes(slide);
  return slide;
}

function addOverfitSlide(presentation, after) {
  const slide = presentation.slides.insert({ after }).slide;
  slide.background.fill = C.white;
  title(slide, "Too much flexibility can overfit", "Regression");
  frame(slide);
  const { chart, sx, sy } = chartMapper();
  axes(slide, chart, sx, sy);
  const sorted = [...data].sort((a, b) => a[0] - b[0]);
  for (let i = 0; i < sorted.length - 1; i++) {
    const [x1, y1] = sorted[i];
    const [x2, y2] = sorted[i + 1];
    dottedCurve(slide, sx, sy, x => {
      const t = (x - x1) / (x2 - x1);
      const wiggle = Math.sin(t * Math.PI) * (i % 2 === 0 ? 7 : -7);
      return y1 * (1 - t) + y2 * t + wiggle;
    }, C.red, x1, x2, 8, 3.2);
  }
  dottedCurve(slide, sx, sy, x => 48 + 47 / (1 + Math.exp(-(x - 72) / 7)), C.green, 52, 99, 60, 3.6);
  box(slide, "fits noise", { left: 990, top: 315, width: 160, height: 54 }, { fill: "#FEF2F2", line: C.red, color: C.red, fontSize: 23 });
  box(slide, "better trend", { left: 990, top: 390, width: 160, height: 54 }, { fill: "#DCFCE7", line: C.green, color: C.green, fontSize: 23 });
  addText(slide, "low training error, weak generalization", { left: 300, top: 180, width: 680, height: 36 }, { fontSize: 28, bold: true, color: C.red, alignment: "center" });
  footer(slide, 7);
  notes(slide);
  return slide;
}

async function renumberFooters(presentation) {
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 200000 });
  for (const lineText of snapshot.ndjson.split(/\r?\n/)) {
    if (!lineText.trim()) continue;
    const item = JSON.parse(lineText);
    const bbox = item.bbox || [];
    if (bbox.length === 4 && Math.abs(bbox[0] - 1166) < 4 && Math.abs(bbox[1] - 660) < 8 && item.slide) {
      const shape = presentation.resolve(item.id);
      shape.text = String(item.slide).padStart(2, "0");
      shape.text.style = { fontSize: 13, color: C.muted, alignment: "right" };
    }
  }
}

async function updateRoadmap(presentation) {
  const replacements = new Map([
    ["Slides 2-15", "Slides 2-17"],
    ["Slides 16-20", "Slides 18-22"],
    ["Slides 21-28", "Slides 23-30"],
    ["Slides 29-32", "Slides 31-34"],
    ["Slides 33-38", "Slides 35-40"],
  ]);
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 40000 });
  for (const lineText of snapshot.ndjson.split(/\r?\n/)) {
    if (!lineText.trim()) continue;
    const item = JSON.parse(lineText);
    if (item.slide !== 1 || !replacements.has(item.textPreview)) continue;
    const shape = presentation.resolve(item.id);
    shape.text = replacements.get(item.textPreview);
    shape.text.style = { fontSize: 15, bold: true, color: C.muted, alignment: "center" };
  }
}

async function main() {
  const presentation = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const slide5 = presentation.slides.getItem(4);
  const nonlinear = addNonlinearSlide(presentation, slide5);
  addOverfitSlide(presentation, nonlinear);
  await renumberFooters(presentation);
  await updateRoadmap(presentation);
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(PPTX);
  console.log(PPTX);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
