import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const C = {
  ink: "#0B0F19",
  muted: "#667085",
  line: "#CBD5E1",
  axis: "#94A3B8",
  panel: "#F8FAFC",
  blue: "#2563EB",
  green: "#10B981",
  orange: "#F97316",
  red: "#EF4444",
  purple: "#7C3AED",
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
  const s = slide.shapes.add({
    geometry: opts.geometry ?? "roundRect",
    position: pos,
    fill: opts.fill ?? C.white,
    line: { style: "solid", fill: opts.line ?? C.line, width: opts.lineWidth ?? 1.5 },
    borderRadius: "rounded-xl",
  });
  if (text) {
    s.text = text;
    s.text.style = {
      fontSize: opts.fontSize ?? 24,
      bold: opts.bold ?? true,
      color: opts.color ?? C.ink,
      alignment: opts.alignment ?? "center",
    };
  }
  return s;
}

function line(slide, x1, y1, x2, y2, opts = {}) {
  return slide.shapes.add({
    geometry: "line",
    position: {
      left: Math.min(x1, x2),
      top: Math.min(y1, y2),
      width: Math.abs(x2 - x1) || 1,
      height: Math.abs(y2 - y1) || 1,
    },
    fill: "none",
    line: {
      style: opts.style ?? "solid",
      fill: opts.color ?? C.line,
      width: opts.width ?? 2,
      endArrowType: opts.endArrowType,
    },
  });
}

function circle(slide, x, y, r, color) {
  return slide.shapes.add({
    geometry: "ellipse",
    position: { left: x - r, top: y - r, width: r * 2, height: r * 2 },
    fill: color,
    line: { style: "solid", fill: color, width: 1 },
  });
}

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function drawLandscape(slide) {
  const left = 112;
  const top = 184;
  const w = 512;
  const h = 344;
  box(slide, "", { left, top, width: w, height: h }, { fill: C.white, line: "#E2E8F0" });
  addText(slide, "loss landscape", { left: left + 28, top: top + 22, width: 230, height: 30 }, { fontSize: 24, bold: true });

  const pts = [];
  for (let i = 0; i <= 90; i++) {
    const x = left + 48 + (i / 90) * 415;
    const t = i / 90;
    const y = top + 250 - Math.sin(t * Math.PI * 3.1) * 44 - t * 68 + Math.pow(t - 0.76, 2) * 145;
    pts.push([x, y]);
  }
  for (let i = 0; i < pts.length - 1; i++) {
    const [x1, y1] = pts[i];
    const [x2, y2] = pts[i + 1];
    line(slide, x1, y1, x2, y2, { color: C.axis, width: 2 });
  }

  const path = [
    [182, 370], [240, 300], [292, 356], [352, 260], [415, 310], [486, 226], [556, 252],
  ];
  for (let i = 0; i < path.length - 1; i++) {
    line(slide, path[i][0], path[i][1], path[i + 1][0], path[i + 1][1], { color: C.orange, width: 2.4, endArrowType: "triangle" });
  }
  for (let i = 0; i < path.length; i++) circle(slide, path[i][0], path[i][1], i === path.length - 1 ? 7 : 5, i === path.length - 1 ? C.green : C.orange);
  addText(slide, "early: explore", { left: 168, top: 480, width: 180, height: 26 }, { fontSize: 19, bold: true, color: C.orange, alignment: "center" });
  addText(slide, "late: settle", { left: 430, top: 480, width: 180, height: 26 }, { fontSize: 19, bold: true, color: C.green, alignment: "center" });
}

function drawSchedule(slide) {
  const left = 700;
  const top = 184;
  const w = 420;
  const h = 344;
  box(slide, "", { left, top, width: w, height: h }, { fill: C.white, line: "#E2E8F0" });
  addText(slide, "annealing", { left: left + 28, top: top + 22, width: 200, height: 30 }, { fontSize: 24, bold: true });
  addText(slide, "temperature / learning rate", { left: left + 28, top: top + 58, width: 290, height: 26 }, { fontSize: 18, color: C.muted });

  line(slide, left + 54, top + 264, left + 356, top + 264, { color: C.axis, width: 2 });
  line(slide, left + 54, top + 264, left + 54, top + 102, { color: C.axis, width: 2 });
  addText(slide, "time", { left: left + 168, top: top + 282, width: 80, height: 24 }, { fontSize: 17, color: C.muted, alignment: "center" });
  addText(slide, "high", { left: left + 20, top: top + 104, width: 56, height: 24 }, { fontSize: 16, color: C.muted, alignment: "right" });
  addText(slide, "low", { left: left + 20, top: top + 244, width: 56, height: 24 }, { fontSize: 16, color: C.muted, alignment: "right" });

  const curve = [];
  for (let i = 0; i <= 56; i++) {
    const t = i / 56;
    curve.push([left + 76 + t * 252, top + 126 + (1 - Math.exp(-3.2 * t)) * 124]);
  }
  for (let i = 0; i < curve.length - 1; i++) {
    line(slide, curve[i][0], curve[i][1], curve[i + 1][0], curve[i + 1][1], { color: C.blue, width: 2.5 });
  }
  circle(slide, curve[0][0], curve[0][1], 6, C.orange);
  circle(slide, curve[curve.length - 1][0], curve[curve.length - 1][1], 6, C.green);
}

async function renumberFooters(presentation) {
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 320000 });
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    const b = item.bbox || [];
    if (b.length === 4 && Math.abs(b[0] - 1166) < 4 && Math.abs(b[1] - 660) < 8 && item.slide) {
      const shape = presentation.resolve(item.id);
      shape.text = String(item.slide).padStart(2, "0");
      shape.text.style = { fontSize: 13, color: C.muted, alignment: "right" };
    }
  }
}

async function main() {
  const p = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const after = p.slides.getItem(10);
  const slide = p.slides.insert({ after }).slide;
  slide.background.fill = C.white;
  addText(slide, "Optimization", { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, "Annealing starts wide, then settles", { left: 64, top: 62, width: 1080, height: 62 }, { fontSize: 40, bold: true });
  box(slide, "", { left: 74, top: 154, width: 1132, height: 474 }, { fill: C.panel, line: "#E2E8F0" });
  drawLandscape(slide);
  drawSchedule(slide);
  addText(slide, "same data, different path", { left: 376, top: 570, width: 520, height: 34 }, { fontSize: 27, bold: true, color: C.ink, alignment: "center" });
  footer(slide, 12);
  slide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on simulated annealing and learning-rate scheduling concepts. Annealing means reducing a temperature-like control over time: early search is wider, later updates become more stable.");
  slide.speakerNotes.setVisible(true);
  await renumberFooters(p);
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log(PPTX);
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
