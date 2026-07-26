import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const C = {
  ink: "#0B0F19",
  muted: "#667085",
  line: "#CBD5E1",
  panel: "#F8FAFC",
  blue: "#2563EB",
  green: "#10B981",
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

function dot(slide, x, y, color) {
  return slide.shapes.add({
    geometry: "ellipse",
    position: { left: x - 8, top: y - 8, width: 16, height: 16 },
    fill: color,
    line: { style: "solid", fill: color, width: 1 },
  });
}

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function drawDataPool(slide) {
  box(slide, "", { left: 86, top: 192, width: 298, height: 352 }, { fill: C.white, line: "#E2E8F0" });
  addText(slide, "one dataset", { left: 112, top: 214, width: 246, height: 32 }, { fontSize: 26, bold: true, alignment: "center" });
  const colors = [C.blue, C.green, C.purple];
  const pts = [
    [146, 292, 0], [190, 342, 0], [244, 284, 0], [301, 378, 0], [169, 431, 0],
    [260, 464, 0], [332, 312, 0], [216, 400, 0], [326, 461, 0],
    [151, 379, 1], [224, 318, 1], [288, 430, 1],
    [337, 388, 2], [196, 474, 2],
  ];
  for (const [x, y, c] of pts) dot(slide, x, y, colors[c]);
}

function drawSplit(slide) {
  const cards = [
    { x: 480, title: "train", color: C.blue, size: "largest", job: "learn weights" },
    { x: 724, title: "validation", color: C.green, size: "medium", job: "choose model" },
    { x: 968, title: "test", color: C.purple, size: "hidden", job: "final estimate" },
  ];
  for (const card of cards) {
    box(slide, "", { left: card.x, top: 192, width: 188, height: 352 }, { fill: C.white, line: card.color, lineWidth: 2.2 });
    addText(slide, card.title, { left: card.x + 12, top: 224, width: 164, height: 34 }, { fontSize: 29, bold: true, color: card.color, alignment: "center" });
    addText(slide, card.size, { left: card.x + 12, top: 284, width: 164, height: 30 }, { fontSize: 22, bold: true, color: C.ink, alignment: "center" });
    addText(slide, card.job, { left: card.x + 18, top: 348, width: 152, height: 56 }, { fontSize: 21, color: C.ink, alignment: "center" });
    addText(slide, card.title === "test" ? "use once" : "repeat use", { left: card.x + 18, top: 458, width: 152, height: 30 }, { fontSize: 17, color: C.muted, alignment: "center" });
  }
  line(slide, 402, 366, 458, 366, { color: C.line, width: 3, endArrowType: "triangle" });
}

async function renumberFooters(presentation) {
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 280000 });
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
  const after = p.slides.getItem(6);
  const slide = p.slides.insert({ after }).slide;
  slide.background.fill = C.white;
  addText(slide, "Regression", { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, "Split data before judging the model", { left: 64, top: 62, width: 1080, height: 62 }, { fontSize: 40, bold: true });
  box(slide, "", { left: 74, top: 154, width: 1132, height: 474 }, { fill: C.panel, line: "#E2E8F0" });
  drawDataPool(slide);
  drawSplit(slide);
  addText(slide, "Same data source, different jobs", { left: 388, top: 574, width: 510, height: 34 }, { fontSize: 27, bold: true, color: C.ink, alignment: "center" });
  footer(slide, 8);
  slide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard train/validation/test model evaluation practice.");
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
