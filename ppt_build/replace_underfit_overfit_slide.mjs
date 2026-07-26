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
      fontSize: opts.fontSize ?? 22,
      bold: opts.bold ?? true,
      color: opts.color ?? C.ink,
      alignment: opts.alignment ?? "center",
    };
  }
  return s;
}

function circle(slide, x, y, r, color) {
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
      beginArrowType: opts.beginArrowType,
      endArrowType: opts.endArrowType,
    },
  });
}

function title(slide) {
  addText(slide, "Regression", { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, "Train, validation, and test reveal fit quality", { left: 64, top: 62, width: 1080, height: 62 }, { fontSize: 40, bold: true });
}

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function drawPipeline(slide) {
  const y = 154;
  const items = [
    { x: 122, title: "train", sub: "fit weights", color: C.blue },
    { x: 482, title: "validation", sub: "choose model", color: C.green },
    { x: 842, title: "test", sub: "final check", color: C.purple },
  ];
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    box(slide, item.title, { left: item.x, top: y, width: 250, height: 72 }, { fill: C.white, line: item.color, lineWidth: 2.2, fontSize: 27, color: item.color });
    addText(slide, item.sub, { left: item.x, top: y + 82, width: 250, height: 28 }, { fontSize: 18, color: C.muted, alignment: "center" });
    if (i < items.length - 1) {
      line(slide, item.x + 272, y + 36, items[i + 1].x - 22, y + 36, { color: C.line, width: 3, endArrowType: "triangle" });
    }
  }
}

function drawCurve(slide, points, color) {
  for (let i = 0; i < points.length - 1; i++) {
    const [x1, y1] = points[i];
    const [x2, y2] = points[i + 1];
    for (let j = 0; j <= 10; j++) {
      const t = j / 10;
      circle(slide, x1 + (x2 - x1) * t, y1 + (y2 - y1) * t, 2.7, color);
    }
  }
}

function drawErrorGraph(slide) {
  const left = 120;
  const top = 328;
  const w = 1035;
  const h = 265;
  box(slide, "", { left: 74, top: 288, width: 1132, height: 348 }, { fill: C.panel, line: "#E2E8F0" });

  line(slide, left, top + h, left + w, top + h, { color: C.axis, width: 2 });
  line(slide, left, top + h, left, top, { color: C.axis, width: 2 });
  addText(slide, "model complexity", { left: 494, top: 605, width: 300, height: 28 }, { fontSize: 18, color: C.muted, alignment: "center" });
  addText(slide, "error", { left: 78, top: 316, width: 70, height: 28 }, { fontSize: 18, color: C.muted, alignment: "center" });

  const train = [
    [150, 394], [300, 430], [455, 470], [610, 510], [780, 542], [950, 564], [1115, 578],
  ];
  const valid = [
    [150, 372], [300, 400], [455, 436], [610, 458], [780, 438], [950, 386], [1115, 330],
  ];
  drawCurve(slide, train, C.blue);
  drawCurve(slide, valid, C.orange);

  addText(slide, "train error", { left: 938, top: 552, width: 150, height: 28 }, { fontSize: 20, bold: true, color: C.blue });
  addText(slide, "validation error", { left: 935, top: 336, width: 190, height: 28 }, { fontSize: 20, bold: true, color: C.orange });

  const zones = [
    { x: 195, label: "underfitting", sub: "both errors high", color: C.red },
    { x: 565, label: "best fit", sub: "validation lowest", color: C.green },
    { x: 912, label: "overfitting", sub: "gap opens", color: C.orange },
  ];
  for (const z of zones) {
    line(slide, z.x, top + 8, z.x, top + h, { color: z.color, width: 1.3, style: "dash" });
    addText(slide, z.label, { left: z.x - 115, top: 294, width: 230, height: 30 }, { fontSize: 23, bold: true, color: z.color, alignment: "center" });
    addText(slide, z.sub, { left: z.x - 115, top: 566, width: 230, height: 28 }, { fontSize: 17, color: C.muted, alignment: "center" });
  }
}

async function renumberFooters(presentation) {
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 240000 });
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
  const oldSlide8 = p.slides.getItem(7);
  oldSlide8.delete();
  const after = p.slides.getItem(6);
  const slide = p.slides.insert({ after }).slide;
  slide.background.fill = C.white;
  title(slide);
  drawPipeline(slide);
  drawErrorGraph(slide);
  footer(slide, 8);
  slide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard machine learning train/validation/test evaluation concepts.");
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
