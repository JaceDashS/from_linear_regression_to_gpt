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

function dot(slide, x, y, r, color) {
  return slide.shapes.add({
    geometry: "ellipse",
    position: { left: x - r, top: y - r, width: r * 2, height: r * 2 },
    fill: color,
    line: { style: "solid", fill: color, width: 1 },
  });
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

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function drawLandscape(slide) {
  const left = 112, top = 184, w = 500, h = 356;
  box(slide, "", { left, top, width: w, height: h }, { fill: C.white, line: "#E2E8F0" });
  addText(slide, "loss surface", { left: left + 30, top: top + 26, width: 220, height: 30 }, { fontSize: 25, bold: true });

  // Smooth valley curve as dots, avoiding fragile diagonal vector paths.
  const pts = [];
  for (let i = 0; i <= 58; i++) {
    const t = i / 58;
    const x = left + 72 + t * 360;
    const y = top + 252 - Math.cos((t - 0.5) * Math.PI) * 72 + Math.sin(t * Math.PI * 2) * 8;
    pts.push([x, y]);
    dot(slide, x, y, 2.6, "#94A3B8");
  }

  const path = [
    [left + 92, top + 88],
    [left + 182, top + 152],
    [left + 250, top + 204],
    [left + 302, top + 236],
    [left + 340, top + 248],
    [left + 370, top + 251],
  ];
  for (const [x, y] of path.slice(0, 3)) dot(slide, x, y, 9, C.orange);
  for (const [x, y] of path.slice(3)) dot(slide, x, y, 8, C.green);
  dot(slide, path[0][0], path[0][1], 16, C.orange);
  dot(slide, path[path.length - 1][0], path[path.length - 1][1], 16, C.green);

  // Use straight horizontal arrows only for stable rendering.
  addText(slide, "large moves", { left: left + 56, top: top + 312, width: 160, height: 28 }, { fontSize: 20, bold: true, color: C.orange, alignment: "center" });
  addText(slide, "small moves", { left: left + 276, top: top + 312, width: 160, height: 28 }, { fontSize: 20, bold: true, color: C.green, alignment: "center" });
  line(slide, left + 214, top + 326, left + 274, top + 326, { color: C.line, width: 2.6, endArrowType: "triangle" });
}

function drawSchedule(slide) {
  const left = 688, top = 184, w = 480, h = 356;
  box(slide, "", { left, top, width: w, height: h }, { fill: C.white, line: "#E2E8F0" });
  addText(slide, "annealing schedule", { left: left + 30, top: top + 26, width: 270, height: 30 }, { fontSize: 25, bold: true });
  addText(slide, "temperature / learning rate", { left: left + 30, top: top + 64, width: 300, height: 26 }, { fontSize: 18, color: C.muted });

  const ax0 = left + 70, ay0 = top + 274;
  const ax1 = left + 400, ay1 = top + 110;
  line(slide, ax0, ay0, ax0, ay1, { color: C.axis, width: 2.2 });
  line(slide, ax0, ay0, ax1, ay0, { color: C.axis, width: 2.2 });
  addText(slide, "high", { left: left + 22, top: ay1 - 12, width: 52, height: 24 }, { fontSize: 16, color: C.muted, alignment: "right" });
  addText(slide, "low", { left: left + 22, top: ay0 - 14, width: 52, height: 24 }, { fontSize: 16, color: C.muted, alignment: "right" });
  addText(slide, "time", { left: left + 210, top: ay0 + 30, width: 80, height: 24 }, { fontSize: 17, color: C.muted, alignment: "center" });

  for (let i = 0; i <= 52; i++) {
    const t = i / 52;
    const x = ax0 + 30 + t * 260;
    const y = ay1 + 28 + (1 - Math.exp(-3.2 * t)) * 120;
    dot(slide, x, y, 3.2, C.blue);
  }
  dot(slide, ax0 + 30, ay1 + 28, 9, C.orange);
  dot(slide, ax0 + 290, ay1 + 148, 9, C.green);

  addText(slide, "explore", { left: left + 88, top: top + 288, width: 110, height: 28 }, { fontSize: 20, bold: true, color: C.orange, alignment: "center" });
  addText(slide, "settle", { left: left + 280, top: top + 288, width: 110, height: 28 }, { fontSize: 20, bold: true, color: C.green, alignment: "center" });
}

async function renumberFooters(presentation) {
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 360000 });
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
  const old = p.slides.getItem(12); // slide 13
  old.delete();
  const slide = p.slides.insert({ after: p.slides.getItem(11) }).slide;
  slide.background.fill = C.white;
  addText(slide, "Optimization", { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, "Annealing reduces randomness over time", { left: 64, top: 62, width: 1080, height: 62 }, { fontSize: 40, bold: true });
  box(slide, "", { left: 74, top: 154, width: 1132, height: 474 }, { fill: C.panel, line: "#E2E8F0" });
  drawLandscape(slide);
  drawSchedule(slide);
  addText(slide, "early search is wide; later updates become stable", { left: 306, top: 570, width: 680, height: 34 }, { fontSize: 26, bold: true, color: C.ink, alignment: "center" });
  footer(slide, 13);
  slide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on simulated annealing and learning-rate scheduling concepts. Annealing reduces a temperature-like control over time: early search is wider, later updates become more stable.");
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
