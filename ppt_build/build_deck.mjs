import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const OUT = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const TMP = "C:\\workspace\\Attention is all you need\\ppt_build\\rendered";
const FORMULA_IMG = "C:\\workspace\\Attention is all you need\\ppt_build\\assets\\mse_formula.png";
const GD_FORMULA_IMG = "C:\\workspace\\Attention is all you need\\ppt_build\\assets\\gd_update_formula.png";
const BT_FORMULA_IMG = "C:\\workspace\\Attention is all you need\\ppt_build\\assets\\backtracking_formula.png";
const PARTIAL_W_IMG = "C:\\workspace\\Attention is all you need\\ppt_build\\assets\\partial_w_formula.png";
const PARTIAL_B_IMG = "C:\\workspace\\Attention is all you need\\ppt_build\\assets\\partial_b_formula.png";
const PARTIAL_BASIC_IMG = "C:\\workspace\\Attention is all you need\\ppt_build\\assets\\partial_derivative_basic.png";

const W = 1280;
const H = 720;
const C = {
  ink: "#0B0F19",
  muted: "#667085",
  line: "#CBD5E1",
  panel: "#F1F5F9",
  panel2: "#E0F2FE",
  blue: "#2563EB",
  sky: "#38BDF8",
  green: "#10B981",
  orange: "#F97316",
  red: "#EF4444",
  purple: "#7C3AED",
  yellow: "#FACC15",
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

async function readImageBlob(imagePath) {
  const bytes = await fs.readFile(imagePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function title(slide, text, kicker = "") {
  if (kicker) addText(slide, kicker, { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, text, { left: 64, top: 62, width: 980, height: 62 }, { fontSize: 42, bold: true });
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

function circle(slide, text, x, y, r, opts = {}) {
  return box(slide, text, { left: x - r, top: y - r, width: r * 2, height: r * 2 }, {
    geometry: "ellipse",
    fill: opts.fill ?? C.white,
    line: opts.line ?? C.line,
    lineWidth: opts.lineWidth ?? 2,
    fontSize: opts.fontSize ?? 20,
    bold: opts.bold ?? true,
    color: opts.color ?? C.ink,
  });
}

function line(slide, x1, y1, x2, y2, opts = {}) {
  const left = Math.min(x1, x2);
  const top = Math.min(y1, y2);
  const width = Math.abs(x2 - x1) || 1;
  const height = Math.abs(y2 - y1) || 1;
  return slide.shapes.add({
    geometry: "line",
    position: { left, top, width, height },
    fill: "none",
    line: { style: opts.style ?? "solid", fill: opts.color ?? C.line, width: opts.width ?? 2 },
  });
}

function freeLine(slide, x1, y1, x2, y2, opts = {}) {
  return slide.shapes.add({
    geometry: "line",
    position: { left: x1, top: y1, width: x2 - x1, height: y2 - y1 },
    fill: "none",
    line: { style: opts.style ?? "solid", fill: opts.color ?? C.line, width: opts.width ?? 2 },
  });
}

function segment(slide, x1, y1, x2, y2, opts = {}) {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const length = Math.sqrt(dx * dx + dy * dy);
  const thickness = opts.width ?? 5;
  return slide.shapes.add({
    geometry: "rect",
    position: {
      left: (x1 + x2) / 2 - length / 2,
      top: (y1 + y2) / 2 - thickness / 2,
      width: length,
      height: thickness,
    },
    rotation: Math.atan2(dy, dx) * 180 / Math.PI,
    fill: opts.color ?? C.orange,
    line: { style: "solid", fill: opts.color ?? C.orange, width: 0 },
  });
}

function arrow(slide, x, y, w, h, text = "", opts = {}) {
  return box(slide, text, { left: x, top: y, width: w, height: h }, {
    geometry: "rightArrow",
    fill: opts.fill ?? C.panel2,
    line: opts.line ?? C.blue,
    lineWidth: opts.lineWidth ?? 1.5,
    fontSize: opts.fontSize ?? 18,
    color: opts.color ?? C.ink,
  });
}

function leftArrow(slide, x, y, w, h, text = "", opts = {}) {
  return box(slide, text, { left: x, top: y, width: w, height: h }, {
    geometry: "leftArrow",
    fill: opts.fill ?? C.panel2,
    line: opts.line ?? C.blue,
    lineWidth: opts.lineWidth ?? 1.5,
    fontSize: opts.fontSize ?? 18,
    color: opts.color ?? C.ink,
  });
}

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function notes(slide, sourceLines) {
  slide.speakerNotes.textFrame.setText([
    "[Sources]",
    ...sourceLines,
  ].join("\n"));
  slide.speakerNotes.setVisible(true);
}

function frame(slide) {
  box(slide, "", { left: 64, top: 142, width: 1152, height: 500 }, { fill: "#F8FAFC", line: "#E2E8F0", radius: "rounded-lg" });
}

function addSlide(p, heading, visual, sourceLines, kicker = "") {
  const slide = p.slides.add();
  slide.background.fill = C.white;
  title(slide, heading, kicker);
  frame(slide);
  visual(slide);
  footer(slide, p.slides.items.length);
  notes(slide, sourceLines);
  return slide;
}

function scatterLine(slide) {
  line(slide, 190, 570, 1060, 570, { color: "#94A3B8", width: 2 });
  line(slide, 190, 570, 190, 205, { color: "#94A3B8", width: 2 });
  const pts = [[260,510],[330,475],[420,452],[500,410],[610,378],[700,330],[800,305],[900,260],[1010,230]];
  for (const [x,y] of pts) circle(slide, "", x, y, 9, { fill: C.blue, line: C.blue });
  line(slide, 235, 525, 1030, 225, { color: C.orange, width: 5 });
  addText(slide, "y = wx + b", { left: 760, top: 490, width: 230, height: 48 }, { fontSize: 34, bold: true, color: C.orange });
}

const regressionData = [
  [1, 1.2],
  [2, 2.0],
  [3, 2.7],
  [4, 3.7],
  [5, 4.4],
  [6, 5.2],
];

function regressionMseSlide(slide, opts, formulaBlob) {
  slide.images.add({
    blob: formulaBlob,
    contentType: "image/png",
    alt: "MSE formula and linear model formula",
    fit: "contain",
    position: { left: 250, top: 163, width: 780, height: 74 },
  });
  const chart = { left: 210, top: 305, width: 780, height: 260 };
  const xMin = 0.5, xMax = 6.5, yMin = 0, yMax = 6;
  const sx = x => chart.left + ((x - xMin) / (xMax - xMin)) * chart.width;
  const sy = y => chart.top + ((yMax - y) / (yMax - yMin)) * chart.height;
  const pred = x => opts.w * x + opts.b;
  const mse = regressionData.reduce((sum, [x, y]) => sum + (y - pred(x)) ** 2, 0) / regressionData.length;

  line(slide, chart.left, chart.top + chart.height, chart.left + chart.width, chart.top + chart.height, { color: "#94A3B8", width: 2 });
  line(slide, chart.left, chart.top + chart.height, chart.left, chart.top, { color: "#94A3B8", width: 2 });
  box(slide, `MSE ≈ ${mse.toFixed(2)}`, { left: 1015, top: 285, width: 145, height: 66 }, { fill: opts.mseFill, line: opts.mseLine, fontSize: 24, color: opts.mseLine });
  addText(slide, opts.caption, { left: 995, top: 370, width: 190, height: 78 }, { fontSize: 22, bold: true, color: opts.mseLine, alignment: "center" });

  for (const [x, y] of regressionData) {
    const px = sx(x);
    const py = sy(y);
    const pY = sy(pred(x));
    line(slide, px, py, px, pY, { color: C.red, width: opts.residualWidth ?? 3 });
    circle(slide, "", px, py, 9, { fill: C.blue, line: C.blue });
  }

  for (let i = 0; i <= 54; i++) {
    const x = 0.8 + (5.4 * i) / 54;
    const y = pred(x);
    circle(slide, "", sx(x), sy(y), 4.2, { fill: C.orange, line: C.orange, lineWidth: 0 });
  }
  addText(slide, "data", { left: 260, top: 267, width: 80, height: 28 }, { fontSize: 19, bold: true, color: C.blue });
  addText(slide, "model line", { left: 705, top: opts.labelTop, width: 160, height: 30 }, { fontSize: 21, bold: true, color: C.orange, alignment: "center" });
}

function attendanceGradeExample(slide) {
  const chart = { left: 200, top: 260, width: 780, height: 300 };
  const xMin = 50, xMax = 100, yMin = 45, yMax = 100;
  const sx = x => chart.left + ((x - xMin) / (xMax - xMin)) * chart.width;
  const sy = y => chart.top + chart.height - ((y - yMin) / (yMax - yMin)) * chart.height;
  const data = [
    [52, 61],
    [55, 48],
    [57, 67],
    [61, 58],
    [63, 73],
    [66, 54],
    [69, 76],
    [72, 70],
    [76, 82],
    [78, 65],
    [81, 88],
    [84, 77],
    [87, 94],
    [90, 83],
    [93, 90],
    [96, 72],
    [98, 97],
  ];
  const pred = x => 0.67 * x + 25;

  line(slide, chart.left, chart.top + chart.height, chart.left + chart.width, chart.top + chart.height, { color: "#94A3B8", width: 2 });
  line(slide, chart.left, chart.top + chart.height, chart.left, chart.top, { color: "#94A3B8", width: 2 });
  addText(slide, "attendance (%)", { left: 500, top: 590, width: 210, height: 28 }, { fontSize: 20, color: C.muted, alignment: "center" });
  addText(slide, "grade", { left: 128, top: 385, width: 70, height: 28 }, { fontSize: 20, color: C.muted, alignment: "center" });

  for (let i = 0; i <= 54; i++) {
    const x = 53 + (44 * i) / 54;
    circle(slide, "", sx(x), sy(pred(x)), 4.2, { fill: C.orange, line: C.orange, lineWidth: 0 });
  }
  for (const [x, y] of data) {
    circle(slide, "", sx(x), sy(y), 8.5, { fill: C.blue, line: C.blue });
  }

  box(slide, "x = attendance", { left: 995, top: 285, width: 170, height: 58 }, { fill: C.panel2, line: C.blue, fontSize: 22, color: C.blue });
  box(slide, "y = grade", { left: 995, top: 360, width: 170, height: 58 }, { fill: "#DCFCE7", line: C.green, fontSize: 22, color: C.green });
  box(slide, "line = prediction", { left: 995, top: 435, width: 170, height: 58 }, { fill: "#FEF3C7", line: C.orange, fontSize: 21, color: C.orange });
  addText(slide, "more attendance,\nhigher predicted grade", { left: 342, top: 178, width: 520, height: 60 }, { fontSize: 28, bold: true, color: C.blue, alignment: "center" });
}

function addFormulaImage(slide, blob, pos, alt) {
  slide.images.add({
    blob,
    contentType: "image/png",
    alt,
    fit: "contain",
    position: pos,
  });
}

function lossCurve(slide, opts = {}) {
  const chart = { left: 185, top: 265, width: 760, height: 300 };
  const f = x => 0.055 * (x - 5.1) * (x - 5.1) + 1.0;
  const sx = x => chart.left + (x / 10) * chart.width;
  const sy = y => chart.top + chart.height - ((y - 0.7) / 2.25) * chart.height;
  line(slide, chart.left, chart.top + chart.height, chart.left + chart.width, chart.top + chart.height, { color: "#94A3B8", width: 2 });
  line(slide, chart.left, chart.top + chart.height, chart.left, chart.top, { color: "#94A3B8", width: 2 });
  for (let i = 0; i <= 80; i++) {
    const x = (10 * i) / 80;
    circle(slide, "", sx(x), sy(f(x)), 3.2, { fill: C.blue, line: C.blue, lineWidth: 0 });
  }
  addText(slide, "parameter", { left: 495, top: 590, width: 160, height: 26 }, { fontSize: 18, color: C.muted, alignment: "center" });
  addText(slide, "MSE", { left: 125, top: 380, width: 60, height: 26 }, { fontSize: 18, color: C.muted, alignment: "center" });
  const points = opts.points ?? [];
  points.forEach((pt) => {
    const x = sx(pt.x);
    const y = sy(f(pt.x));
    circle(slide, pt.label ?? "", x, y, pt.r ?? 13, { fill: pt.fill ?? C.orange, line: pt.line ?? C.orange, fontSize: 14, color: C.white });
  });
  return { chart, f, sx, sy };
}

function stepArrow(slide, curve, x1, x2, opts = {}) {
  const color = opts.color ?? C.orange;
  for (let i = 0; i <= 16; i++) {
    const t = i / 16;
    const x = x1 + (x2 - x1) * t;
    circle(slide, "", curve.sx(x), curve.sy(curve.f(x)), opts.dot ?? 3.8, { fill: color, line: color, lineWidth: 0 });
  }
  circle(slide, "", curve.sx(x2), curve.sy(curve.f(x2)), opts.r ?? 11, { fill: color, line: color });
}

function tangentLine(slide, curve, x0, opts = {}) {
  const x = curve.sx(x0);
  const y = curve.sy(curve.f(x0));
  line(slide, x - 130, y - 18, x + 130, y - 18, { color: opts.color ?? C.purple, width: opts.width ?? 2 });
}

function partialDerivativeSlide(slide, partialBasicBlob) {
  addFormulaImage(slide, partialBasicBlob, { left: 115, top: 170, width: 1050, height: 190 }, "Basic partial derivative example");
  box(slide, "1", { left: 160, top: 410, width: 58, height: 58 }, { geometry: "ellipse", fill: C.panel2, line: C.blue, fontSize: 26, color: C.blue });
  addText(slide, "Choose the variable", { left: 235, top: 420, width: 250, height: 34 }, { fontSize: 25, bold: true, color: C.ink });
  box(slide, "2", { left: 160, top: 495, width: 58, height: 58 }, { geometry: "ellipse", fill: "#FEF3C7", line: C.orange, fontSize: 26, color: C.orange });
  addText(slide, "Treat the other variable as a constant", { left: 235, top: 505, width: 485, height: 34 }, { fontSize: 25, bold: true, color: C.ink });
  box(slide, "3", { left: 760, top: 452, width: 58, height: 58 }, { geometry: "ellipse", fill: "#DCFCE7", line: C.green, fontSize: 26, color: C.green });
  addText(slide, "Differentiate normally", { left: 835, top: 462, width: 300, height: 34 }, { fontSize: 25, bold: true, color: C.ink });
  addText(slide, "Partial derivative = slope with respect to one variable.", { left: 245, top: 600, width: 790, height: 34 }, { fontSize: 27, bold: true, color: C.blue, alignment: "center" });
}

function bars(slide, x, y, vals, color) {
  vals.forEach((v, i) => {
    box(slide, "", { left: x + i * 48, top: y + (100 - v), width: 30, height: v }, { fill: color, line: color, radius: "rounded-sm" });
  });
}

const common = ["Educational synthesis based on standard machine learning concepts."];
const transformerSrc = ["Vaswani et al., 2017, 'Attention Is All You Need', arXiv:1706.03762, https://arxiv.org/abs/1706.03762"];
const gptSrc = [
  "Radford et al., 2018, 'Improving Language Understanding by Generative Pre-Training', OpenAI.",
  "Radford et al., 2019, 'Language Models are Unsupervised Multitask Learners', OpenAI.",
  "Brown et al., 2020, 'Language Models are Few-Shot Learners', arXiv:2005.14165, https://arxiv.org/abs/2005.14165",
];

async function main() {
  await fs.mkdir(TMP, { recursive: true });
  const formulaBlob = await readImageBlob(FORMULA_IMG);
  const gdFormulaBlob = await readImageBlob(GD_FORMULA_IMG);
  const btFormulaBlob = await readImageBlob(BT_FORMULA_IMG);
  const partialWBlob = await readImageBlob(PARTIAL_W_IMG);
  const partialBBlob = await readImageBlob(PARTIAL_B_IMG);
  const partialBasicBlob = await readImageBlob(PARTIAL_BASIC_IMG);
  const p = Presentation.create({ slideSize: { width: W, height: H } });

  {
    const s = p.slides.add();
    s.background.fill = C.white;
    addText(s, "From Linear Regression to GPT", { left: 70, top: 56, width: 870, height: 62 }, { fontSize: 50, bold: true });
    addText(s, "Roadmap", { left: 74, top: 128, width: 260, height: 40 }, { fontSize: 30, bold: true, color: C.blue });
    addText(s, "From simple prediction to GPT", { left: 280, top: 132, width: 520, height: 34 }, { fontSize: 23, color: C.muted });
    box(s, "", { left: 70, top: 215, width: 1140, height: 350 }, { fill: "#F8FAFC", line: "#E2E8F0", radius: "rounded-lg" });
    line(s, 160, 390, 1120, 390, { color: C.line, width: 5 });
    const sections = [
      ["01", "Foundations", "MSE, partials,\ngradients", "Slides 2-15", C.blue],
      ["02", "Language", "tokens and\nembeddings", "Slides 16-20", C.green],
      ["03", "Attention", "query, key,\nvalue", "Slides 21-28", C.orange],
      ["04", "Transformer", "blocks and\nmasking", "Slides 29-32", C.purple],
      ["05", "GPT", "next-token\nprediction", "Slides 33-38", C.red],
    ];
    sections.forEach(([num, head, body, range, color], i) => {
      const x = 170 + i * 232;
      circle(s, num, x, 390, 42, { fill: i === 0 ? C.panel2 : C.white, line: color, lineWidth: 3, fontSize: 22, color });
      addText(s, head, { left: x - 95, top: 455, width: 190, height: 30 }, { fontSize: 24, bold: true, alignment: "center" });
      addText(s, body, { left: x - 96, top: 492, width: 192, height: 52 }, { fontSize: 18, color: C.ink, alignment: "center" });
      addText(s, range, { left: x - 80, top: 300, width: 160, height: 26 }, { fontSize: 15, bold: true, color: C.muted, alignment: "center" });
    });
    addText(s, "Goal: understand why attention made GPT possible", { left: 205, top: 608, width: 870, height: 40 }, { fontSize: 28, bold: true, color: C.blue, alignment: "center" });
    footer(s, 1);
    notes(s, ["This slide now functions as the table of contents: foundations, language representation, attention, Transformer, and GPT."]);
  }

  addSlide(p, "Start with a bad line", s => {
    regressionMseSlide(s, {
      w: -0.25,
      b: 5.3,
      mseFill: "#FEF2F2",
      mseLine: C.red,
      caption: "large errors",
      residualWidth: 4,
      labelTop: 450,
    }, formulaBlob);
  }, common, "Linear regression");

  addSlide(p, "Change w and b to lower MSE", s => {
    regressionMseSlide(s, {
      w: 0.55,
      b: 1.25,
      mseFill: "#FEF3C7",
      mseLine: C.orange,
      caption: "closer line",
      residualWidth: 3,
      labelTop: 390,
    }, formulaBlob);
  }, common, "Linear regression");

  addSlide(p, "The fitted line makes MSE small", s => {
    regressionMseSlide(s, {
      w: 0.80,
      b: 0.35,
      mseFill: "#DCFCE7",
      mseLine: C.green,
      caption: "small errors",
      residualWidth: 2,
      labelTop: 348,
    }, formulaBlob);
  }, common, "Linear regression");

  addSlide(p, "Example: attendance predicts grades", s => {
    attendanceGradeExample(s);
  }, common, "Linear regression");

  addSlide(p, "Partial derivatives: one variable at a time", s => {
    partialDerivativeSlide(s, partialBasicBlob);
  }, common, "Partial derivatives");

  addSlide(p, "Gradient descent changes w and b", s => {
    addFormulaImage(s, gdFormulaBlob, { left: 190, top: 165, width: 845, height: 70 }, "Gradient descent update formula");
    const curve = lossCurve(s, { points: [{ x: 8.0, label: "", fill: C.red, line: C.red }] });
    tangentLine(s, curve, 8.0, { color: C.purple });
    stepArrow(s, curve, 8.0, 6.7, { color: C.orange });
    stepArrow(s, curve, 6.7, 5.7, { color: C.green });
    addText(s, "tangent line", { left: 925, top: 382, width: 170, height: 28 }, { fontSize: 22, bold: true, color: C.purple, alignment: "center" });
    addText(s, "negative gradient", { left: 850, top: 305, width: 210, height: 30 }, { fontSize: 23, bold: true, color: C.orange, alignment: "center" });
    addText(s, "lower MSE", { left: 785, top: 500, width: 160, height: 30 }, { fontSize: 23, bold: true, color: C.green, alignment: "center" });
  }, common, "Gradient descent");

  addSlide(p, "Learning rate controls the step", s => {
    const curve = lossCurve(s, { points: [{ x: 8.2, fill: C.blue, line: C.blue }] });
    stepArrow(s, curve, 8.2, 7.8, { color: C.green, dot: 3.5, r: 10 });
    stepArrow(s, curve, 8.2, 6.2, { color: C.orange, dot: 4.2, r: 12 });
    stepArrow(s, curve, 8.2, 3.0, { color: C.red, dot: 4.2, r: 12 });
    box(s, "small α", { left: 995, top: 300, width: 140, height: 54 }, { fill: "#DCFCE7", line: C.green, fontSize: 23, color: C.green });
    box(s, "good α", { left: 995, top: 375, width: 140, height: 54 }, { fill: "#FEF3C7", line: C.orange, fontSize: 23, color: C.orange });
    box(s, "too large", { left: 995, top: 450, width: 140, height: 54 }, { fill: "#FEF2F2", line: C.red, fontSize: 23, color: C.red });
    addText(s, "α is the learning rate", { left: 320, top: 182, width: 560, height: 40 }, { fontSize: 32, bold: true, color: C.blue, alignment: "center" });
  }, common, "Learning rate");

  addSlide(p, "Backtracking line search shrinks bad steps", s => {
    addFormulaImage(s, btFormulaBlob, { left: 180, top: 160, width: 850, height: 78 }, "Backtracking line search rule");
    const curve = lossCurve(s, { points: [{ x: 8.0, fill: C.blue, line: C.blue }] });
    stepArrow(s, curve, 8.0, 2.6, { color: C.red, dot: 4.1, r: 12 });
    stepArrow(s, curve, 8.0, 5.7, { color: C.green, dot: 4.1, r: 12 });
    addText(s, "reject", { left: 470, top: 310, width: 120, height: 28 }, { fontSize: 24, bold: true, color: C.red, alignment: "center" });
    addText(s, "accept", { left: 760, top: 420, width: 120, height: 28 }, { fontSize: 24, bold: true, color: C.green, alignment: "center" });
    box(s, "α → α/2", { left: 995, top: 380, width: 145, height: 62 }, { fill: C.panel2, line: C.blue, fontSize: 27, color: C.blue });
  }, common, "Line search");

  addSlide(p, "The chain rule connects causes", s => {
    const xs = [160, 365, 570, 775, 980];
    ["x", "z = wx + b", "a = f(z)", "prediction", "loss"].forEach((t, i) => {
      box(s, t, { left: xs[i], top: 340, width: i === 1 ? 170 : 150, height: 76 }, { fill: i === 4 ? "#FEF2F2" : C.white, line: i === 4 ? C.red : C.line, fontSize: 22 });
      if (i < 4) arrow(s, xs[i] + (i === 1 ? 185 : 165), 356, 70, 42, "");
    });
    addText(s, "dL/dw = dL/dp * dp/da * da/dz * dz/dw", { left: 195, top: 500, width: 870, height: 50 }, { fontSize: 32, bold: true, color: C.blue, alignment: "center" });
  }, common, "Calculus");

  addSlide(p, "Matrices move many numbers at once", s => {
    const drawGrid = (x, y, rows, cols, label, fill) => {
      for (let r=0;r<rows;r++) for (let c=0;c<cols;c++) box(s, "", { left: x+c*42, top: y+r*42, width: 34, height: 34 }, { fill, line: C.line, radius: "rounded-sm" });
      addText(s, label, { left: x, top: y+rows*42+12, width: cols*42, height: 32 }, { fontSize: 28, bold: true, alignment: "center" });
    };
    drawGrid(210, 255, 4, 3, "X", C.white);
    addText(s, "x", { left: 375, top: 318, width: 40, height: 40 }, { fontSize: 40, bold: true, alignment: "center" });
    drawGrid(455, 215, 3, 5, "W", C.panel2);
    addText(s, "=", { left: 705, top: 318, width: 40, height: 40 }, { fontSize: 40, bold: true, alignment: "center" });
    drawGrid(785, 255, 4, 5, "Y", "#DCFCE7");
  }, common, "Parallel math");

  addSlide(p, "A neuron is a tiny calculator", s => {
    [260, 360, 460].forEach((y, i) => {
      circle(s, `x${i+1}`, 230, y, 34, { fill: C.white, line: C.line });
      line(s, 265, y, 545, 360, { color: C.line, width: 3 });
      addText(s, `w${i+1}`, { left: 355, top: y-24, width: 50, height: 24 }, { fontSize: 18, color: C.muted, alignment: "center" });
    });
    circle(s, "sum", 600, 360, 70, { fill: C.panel2, line: C.blue, fontSize: 28 });
    arrow(s, 690, 338, 120, 44, "f");
    circle(s, "output", 900, 360, 58, { fill: "#DCFCE7", line: C.green, fontSize: 24 });
  }, common, "Neural networks");

  addSlide(p, "A network stacks many neurons", s => {
    const layers = [[230,3],[480,5],[740,5],[1000,2]];
    layers.forEach(([x,n], li) => {
      for (let i=0;i<n;i++) circle(s, "", x, 255+i*(260/(n-1||1)), 24, { fill: li === 3 ? "#DCFCE7" : C.white, line: li === 0 ? C.muted : C.blue });
      if (li < layers.length-1) {
        const [nx, nn] = layers[li+1];
        for (let i=0;i<n;i++) for (let j=0;j<nn;j++) line(s, x+24, 255+i*(260/(n-1||1)), nx-24, 255+j*(260/(nn-1||1)), { color: "#CBD5E1/45", width: 1 });
      }
    });
    ["input","hidden","hidden","output"].forEach((t,i) => addText(s, t, { left: layers[i][0]-65, top: 560, width: 130, height: 28 }, { fontSize: 20, color: C.muted, alignment: "center" }));
  }, common, "Depth");

  addSlide(p, "Backprop sends error backward", s => {
    ["input","hidden","hidden","output","loss"].forEach((t,i) => box(s, t, { left: 130+i*215, top: 320, width: 145, height: 78 }, { fill: i === 4 ? "#FEF2F2" : C.white, line: i === 4 ? C.red : C.line }));
    for (let i=0;i<4;i++) arrow(s, 285+i*215, 338, 82, 40, "", { fill: "#DBEAFE" });
    for (let i=3;i>=0;i--) leftArrow(s, 165+i*215, 455, 80, 40, "", { fill: "#FFE4E6", line: C.red });
    addText(s, "forward", { left: 500, top: 250, width: 160, height: 36 }, { fontSize: 26, bold: true, color: C.blue, alignment: "center" });
    addText(s, "gradients", { left: 500, top: 512, width: 180, height: 36 }, { fontSize: 26, bold: true, color: C.red, alignment: "center" });
  }, common, "Training");

  addSlide(p, "Training is a loop", s => {
    const cx = 640, cy = 380;
    [["predict",cx,210],["measure",840,380],["backprop",640,550],["update",440,380]].forEach(([t,x,y], i) => circle(s, t, x, y, 62, { fill: i === 1 ? "#FEF2F2" : C.white, line: i === 1 ? C.red : C.blue, fontSize: 22 }));
    arrow(s, 680, 238, 135, 58, "", { fill: "#DBEAFE" });
    leftArrow(s, 660, 428, 120, 58, "", { fill: "#DBEAFE" });
    arrow(s, 470, 428, 120, 58, "", { fill: "#DBEAFE" });
    leftArrow(s, 333, 238, 135, 58, "", { fill: "#DBEAFE" });
    addText(s, "repeat", { left: 560, top: 360, width: 160, height: 42 }, { fontSize: 34, bold: true, alignment: "center" });
  }, common, "Loop");

  addSlide(p, "Words become vectors", s => {
    ["The","cat","sat"].forEach((t,i) => box(s, t, { left: 170+i*140, top: 320, width: 95, height: 58 }, { fill: C.white }));
    arrow(s, 575, 325, 125, 50, "");
    [0.75,0.25,0.9,0.4,0.6].forEach((v,i) => box(s, "", { left: 780+i*55, top: 390-v*170, width: 35, height: v*170 }, { fill: i%2 ? C.sky : C.blue, line: "none", radius: "rounded-sm" }));
    addText(s, "[0.75, 0.25, ...]", { left: 765, top: 430, width: 310, height: 40 }, { fontSize: 28, bold: true, alignment: "center" });
  }, common, "Embeddings");

  addSlide(p, "Meaning becomes geometry", s => {
    line(s, 195, 535, 1050, 535, { color: C.line });
    line(s, 195, 535, 195, 205, { color: C.line });
    [["king",370,285],["queen",420,250],["man",330,380],["woman",435,365],["cat",760,330],["dog",820,300],["car",930,475]].forEach(([t,x,y]) => {
      circle(s, "", x, y, 11, { fill: t === "car" ? C.orange : C.blue, line: t === "car" ? C.orange : C.blue });
      addText(s, t, { left: x+14, top: y-15, width: 90, height: 24 }, { fontSize: 18, color: C.ink });
    });
  }, common, "Embeddings");

  addSlide(p, "Order changes the meaning", s => {
    const sent = (y, a, b) => {
      box(s, a, { left: 250, top: y, width: 150, height: 62 }, { fill: C.panel2, line: C.blue });
      arrow(s, 430, y+10, 210, 42, "chased", { fontSize: 20 });
      box(s, b, { left: 680, top: y, width: 150, height: 62 }, { fill: "#DCFCE7", line: C.green });
    };
    sent(280, "dog", "cat");
    sent(430, "cat", "dog");
  }, common, "Sequences");

  addSlide(p, "RNNs read one token at a time", s => {
    ["The","cat","sat","there"].forEach((t,i) => {
      box(s, t, { left: 165+i*230, top: 355, width: 110, height: 60 }, { fill: C.white });
      circle(s, "h", 220+i*230, 255, 45, { fill: C.panel2, line: C.blue, fontSize: 28 });
      line(s, 220+i*230, 310, 220+i*230, 355, { color: C.line, width: 3 });
      if (i < 3) arrow(s, 278+i*230, 236, 130, 38, "");
    });
  }, common, "Before attention");

  addSlide(p, "Long memory fades", s => {
    ["The","animal","near","the","river","was","tired"].forEach((t,i) => box(s, t, { left: 110+i*155, top: 365, width: 105, height: 54 }, { fill: C.white, fontSize: 18 }));
    line(s, 160, 335, 1040, 335, { color: C.red, width: 5 });
    [160,315,470,625,780,935,1040].forEach((x,i) => circle(s, "", x, 335, 10, { fill: i < 2 ? C.red : "#FCA5A5", line: i < 2 ? C.red : "#FCA5A5" }));
    addText(s, "early context must survive many steps", { left: 325, top: 245, width: 630, height: 40 }, { fontSize: 30, bold: true, color: C.red, alignment: "center" });
  }, common, "Limit");

  addSlide(p, "Attention lets words look around", s => {
    const xs = [220, 390, 560, 730, 900];
    xs.forEach((x,i) => circle(s, ["The","cat","sat","on","mat"][i], x, 380, 46, { fill: i === 1 ? C.panel2 : C.white, line: i === 1 ? C.blue : C.line, fontSize: 18 }));
    xs.forEach((x,i) => { if (i !== 1) line(s, 390, 380, x, 380, { color: i === 4 ? C.orange : "#94A3B8", width: i === 4 ? 5 : 2 }); });
    addText(s, "one token compares with all others", { left: 340, top: 245, width: 600, height: 40 }, { fontSize: 30, bold: true, color: C.blue, alignment: "center" });
  }, transformerSrc, "Attention");

  addSlide(p, "Query, key, value", s => {
    circle(s, "Q", 270, 360, 70, { fill: C.panel2, line: C.blue, fontSize: 40 });
    circle(s, "K", 640, 265, 70, { fill: "#FEF3C7", line: C.orange, fontSize: 40 });
    circle(s, "V", 640, 455, 70, { fill: "#DCFCE7", line: C.green, fontSize: 40 });
    arrow(s, 355, 334, 185, 48, "match");
    arrow(s, 720, 435, 180, 48, "carry");
    box(s, "context", { left: 925, top: 405, width: 160, height: 80 }, { fill: C.white, fontSize: 26 });
  }, transformerSrc, "Mechanism");

  addSlide(p, "Scores come from dot products", s => {
    box(s, "Q", { left: 220, top: 315, width: 110, height: 90 }, { fill: C.panel2, line: C.blue, fontSize: 40 });
    addText(s, "dot", { left: 380, top: 338, width: 80, height: 38 }, { fontSize: 30, bold: true, alignment: "center" });
    box(s, "K", { left: 500, top: 315, width: 110, height: 90 }, { fill: "#FEF3C7", line: C.orange, fontSize: 40 });
    addText(s, "=", { left: 650, top: 338, width: 80, height: 38 }, { fontSize: 38, bold: true, alignment: "center" });
    bars(s, 770, 330, [40,75,25,95,55], C.blue);
    addText(s, "relevance", { left: 740, top: 455, width: 290, height: 36 }, { fontSize: 28, bold: true, alignment: "center" });
  }, transformerSrc, "Mechanism");

  addSlide(p, "Softmax makes weights", s => {
    bars(s, 260, 300, [35,80,20,95,45], C.orange);
    arrow(s, 535, 330, 170, 55, "softmax", { fontSize: 20 });
    bars(s, 800, 300, [25,70,15,90,40], C.green);
    addText(s, "raw scores", { left: 205, top: 460, width: 300, height: 34 }, { fontSize: 25, bold: true, alignment: "center" });
    addText(s, "weights sum to 1", { left: 730, top: 460, width: 350, height: 34 }, { fontSize: 25, bold: true, alignment: "center" });
  }, transformerSrc, "Mechanism");

  addSlide(p, "Values are mixed into context", s => {
    [0.15,0.55,0.1,0.2].forEach((w,i) => {
      box(s, `V${i+1}`, { left: 180+i*155, top: 315, width: 90, height: 80 }, { fill: "#DCFCE7", line: C.green, fontSize: 28 });
      addText(s, String(w), { left: 180+i*155, top: 410, width: 90, height: 30 }, { fontSize: 22, color: C.blue, bold: true, alignment: "center" });
    });
    arrow(s, 790, 335, 120, 45, "sum");
    box(s, "context vector", { left: 940, top: 300, width: 170, height: 110 }, { fill: C.panel2, line: C.blue, fontSize: 24 });
  }, transformerSrc, "Mechanism");

  addSlide(p, "Self-attention updates every token", s => {
    const xs = [245, 395, 545, 695, 845, 995];
    xs.forEach((x,i) => circle(s, String(i+1), x, 300, 32, { fill: C.white, line: C.blue }));
    xs.forEach((x,i) => circle(s, String(i+1)+"'", x, 480, 32, { fill: C.panel2, line: C.blue }));
    xs.forEach((x,i) => line(s, x, 335, x, 445, { color: C.line, width: 2 }));
    for (let i=0;i<xs.length;i++) for (let j=0;j<xs.length;j+=2) line(s, xs[i], 300, xs[j], 480, { color: "#CBD5E1/35", width: 1 });
  }, transformerSrc, "Self-attention");

  addSlide(p, "Many heads learn different links", s => {
    [["syntax",250,C.blue],["reference",515,C.green],["topic",780,C.orange],["position",1045,C.purple]].forEach(([t,x,c]) => {
      circle(s, "", x, 330, 80, { fill: `${c}/12`, line: c, lineWidth: 3 });
      line(s, x-45, 330, x+45, 330, { color: c, width: 4 });
      line(s, x, 285, x, 375, { color: c, width: 4 });
      addText(s, t, { left: x-75, top: 440, width: 150, height: 30 }, { fontSize: 22, bold: true, alignment: "center" });
    });
  }, transformerSrc, "Multi-head");

  addSlide(p, "Position is added to meaning", s => {
    box(s, "token vector", { left: 250, top: 310, width: 210, height: 90 }, { fill: C.panel2, line: C.blue });
    addText(s, "+", { left: 505, top: 330, width: 40, height: 40 }, { fontSize: 42, bold: true, alignment: "center" });
    box(s, "position vector", { left: 585, top: 310, width: 230, height: 90 }, { fill: "#FEF3C7", line: C.orange });
    addText(s, "=", { left: 850, top: 330, width: 40, height: 40 }, { fontSize: 42, bold: true, alignment: "center" });
    box(s, "ordered meaning", { left: 925, top: 310, width: 230, height: 90 }, { fill: "#DCFCE7", line: C.green });
  }, transformerSrc, "Order");

  addSlide(p, "A Transformer block mixes and computes", s => {
    box(s, "input", { left: 210, top: 485, width: 170, height: 54 }, { fill: C.white });
    box(s, "self-attention", { left: 445, top: 405, width: 260, height: 80 }, { fill: C.panel2, line: C.blue, fontSize: 26 });
    box(s, "feed-forward", { left: 775, top: 285, width: 260, height: 80 }, { fill: "#DCFCE7", line: C.green, fontSize: 26 });
    box(s, "output", { left: 915, top: 190, width: 170, height: 54 }, { fill: C.white });
    line(s, 380, 512, 445, 445, { color: C.line, width: 3 });
    line(s, 705, 445, 775, 325, { color: C.line, width: 3 });
    line(s, 1035, 325, 1000, 245, { color: C.line, width: 3 });
    addText(s, "residual + norm around each step", { left: 390, top: 555, width: 520, height: 34 }, { fontSize: 25, bold: true, color: C.muted, alignment: "center" });
  }, transformerSrc, "Block");

  addSlide(p, "The encoder reads the whole input", s => {
    ["I","love","machine","learning"].forEach((t,i) => box(s, t, { left: 140+i*136, top: 520, width: i === 3 ? 128 : 96, height: 50 }, { fill: C.white, fontSize: i === 3 ? 17 : 18 }));
    [410,330,250].forEach((y,i) => box(s, `encoder ${i+1}`, { left: 680, top: y, width: 260, height: 70 }, { fill: i === 1 ? C.panel2 : C.white, line: C.blue, fontSize: 24 }));
    arrow(s, 565, 460, 95, 38, "");
    addText(s, "contextual representations", { left: 620, top: 170, width: 410, height: 38 }, { fontSize: 28, bold: true, color: C.blue, alignment: "center" });
  }, transformerSrc, "Encoder");

  addSlide(p, "The decoder hides the future", s => {
    const x0=360,y0=240,sz=46;
    for (let r=0;r<6;r++) for (let c=0;c<6;c++) box(s, "", { left:x0+c*sz, top:y0+r*sz, width:sz-4, height:sz-4 }, { fill: c<=r ? C.panel2 : "#FEE2E2", line: C.white, radius:"rounded-sm" });
    addText(s, "visible", { left: 250, top: 500, width: 160, height: 30 }, { fontSize: 24, bold: true, color: C.blue, alignment:"center" });
    addText(s, "masked", { left: 780, top: 500, width: 160, height: 30 }, { fontSize: 24, bold: true, color: C.red, alignment:"center" });
  }, transformerSrc, "Decoder");

  addSlide(p, "Attention replaced recurrence", s => {
    box(s, "RNN\nstep by step", { left: 180, top: 300, width: 290, height: 130 }, { fill: C.white, fontSize: 30 });
    addText(s, "->", { left: 585, top: 340, width: 70, height: 45 }, { fontSize: 44, bold: true, color: C.blue, alignment: "center" });
    box(s, "Transformer\ndirect comparison", { left: 770, top: 300, width: 330, height: 130 }, { fill: C.panel2, line: C.blue, fontSize: 30 });
  }, transformerSrc, "2017");

  addSlide(p, "GPT keeps the decoder idea", s => {
    [500,405,310,215].forEach((y,i) => box(s, `decoder block ${4-i}`, { left: 455, top: y, width: 310, height: 58 }, { fill: i%2 ? C.white : C.panel2, line: C.blue, fontSize: 23 }));
    box(s, "prompt tokens", { left: 170, top: 500, width: 200, height: 62 }, { fill: C.white });
    arrow(s, 385, 507, 65, 45, "");
    arrow(s, 770, 225, 145, 45, "");
    box(s, "next token", { left: 940, top: 215, width: 180, height: 62 }, { fill: "#DCFCE7", line: C.green, fontSize: 25 });
  }, gptSrc, "GPT");

  addSlide(p, "Training target: the next token", s => {
    ["The","capital","of","France","is"].forEach((t,i) => box(s, t, { left: 150+i*160, top: 330, width: 125, height: 62 }, { fill: C.white, fontSize: 22 }));
    arrow(s, 940, 338, 95, 45, "");
    box(s, "Paris", { left: 1060, top: 330, width: 125, height: 62 }, { fill: "#DCFCE7", line: C.green, fontSize: 24 });
    addText(s, "predict this", { left: 1015, top: 440, width: 210, height: 34 }, { fontSize: 26, bold: true, color: C.green, alignment: "center" });
  }, gptSrc, "Objective");

  addSlide(p, "Pretraining compresses patterns into weights", s => {
    box(s, "text", { left: 150, top: 310, width: 155, height: 90 }, { fill: C.white });
    arrow(s, 330, 334, 110, 45, "tokens", { fontSize: 18 });
    box(s, "batches", { left: 470, top: 310, width: 170, height: 90 }, { fill: C.white });
    arrow(s, 665, 334, 110, 45, "train", { fontSize: 18 });
    box(s, "Transformer", { left: 805, top: 295, width: 210, height: 120 }, { fill: C.panel2, line: C.blue, fontSize: 30 });
    arrow(s, 1035, 334, 90, 45, "");
    circle(s, "weights", 1155, 357, 54, { fill: "#DCFCE7", line: C.green, fontSize: 20 });
  }, gptSrc, "Scale");

  addSlide(p, "Inference repeats the same move", s => {
    box(s, "prompt", { left: 200, top: 330, width: 160, height: 70 }, { fill: C.white });
    arrow(s, 385, 344, 125, 45, "");
    box(s, "model", { left: 540, top: 300, width: 200, height: 130 }, { fill: C.panel2, line: C.blue, fontSize: 32 });
    arrow(s, 765, 344, 125, 45, "");
    box(s, "new token", { left: 920, top: 330, width: 170, height: 70 }, { fill: "#DCFCE7", line: C.green });
    line(s, 1005, 405, 285, 500, { color: C.green, width: 4 });
    addText(s, "append and repeat", { left: 485, top: 520, width: 300, height: 35 }, { fontSize: 27, bold: true, color: C.green, alignment: "center" });
  }, gptSrc, "Generation");

  addSlide(p, "Sampling changes the voice", s => {
    addText(s, "low temperature", { left: 230, top: 230, width: 250, height: 32 }, { fontSize: 25, bold: true, alignment: "center" });
    bars(s, 270, 365, [95,25,10,8,5], C.blue);
    addText(s, "high temperature", { left: 760, top: 230, width: 270, height: 32 }, { fontSize: 25, bold: true, alignment: "center" });
    bars(s, 800, 365, [55,48,43,35,30], C.orange);
  }, gptSrc, "Decoding");

  addSlide(p, "The whole path is one learning story", s => {
    const items = [["line",170],["network",330],["backprop",500],["attention",690],["Transformer",890],["GPT",1060]];
    line(s, 150, 380, 1080, 380, { color: C.line, width: 4 });
    items.forEach(([t,x],i) => {
      circle(s, String(i+1), x, 380, 35, { fill: i >= 3 ? C.panel2 : C.white, line: i === 5 ? C.green : C.blue });
      addText(s, t, { left: x-75, top: 445, width: 150, height: 30 }, { fontSize: 21, bold: true, alignment: "center" });
    });
    addText(s, "learn from error -> compare tokens -> predict the next token", { left: 215, top: 245, width: 850, height: 44 }, { fontSize: 32, bold: true, color: C.blue, alignment: "center" });
  }, [...transformerSrc, ...gptSrc], "Recap");

  for (const [index, slide] of p.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    const png = await p.export({ slide, format: "png", scale: 1 });
    await fs.writeFile(path.join(TMP, `${stem}.png`), Buffer.from(await png.arrayBuffer()));
  }
  const montage = await p.export({ format: "webp", montage: true, scale: 1 });
  await fs.writeFile(path.join(TMP, "montage.webp"), Buffer.from(await montage.arrayBuffer()));
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(OUT);
  console.log(OUT);
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
