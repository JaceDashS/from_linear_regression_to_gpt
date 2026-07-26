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

function circle(slide, x, y, r, color, stroke = color) {
  return slide.shapes.add({
    geometry: "ellipse",
    position: { left: x - r, top: y - r, width: r * 2, height: r * 2 },
    fill: color,
    line: { style: "solid", fill: stroke, width: 1.4 },
  });
}

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function modelStack(slide, x, y, color, layers, label, sub) {
  for (let i = 0; i < layers; i++) {
    box(slide, "", { left: x + i * 10, top: y - i * 13, width: 164, height: 56 }, {
      fill: i === layers - 1 ? C.white : "#EEF2F7",
      line: i === layers - 1 ? color : "#CBD5E1",
      lineWidth: i === layers - 1 ? 2.2 : 1.2,
    });
  }
  addText(slide, label, { left: x - 22, top: y + 94, width: 240, height: 32 }, { fontSize: 27, bold: true, color, alignment: "center" });
  addText(slide, sub, { left: x - 22, top: y + 134, width: 240, height: 30 }, { fontSize: 18, color: C.muted, alignment: "center" });
}

function token(slide, text, x, y, color) {
  return box(slide, text, { left: x, top: y, width: 92, height: 52 }, { fill: C.white, line: color, lineWidth: 1.8, fontSize: 19, color: C.ink });
}

function drawSlide(slide) {
  slide.background.fill = C.white;
  addText(slide, "Compression", { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, "Teacher-student distillation transfers behavior", { left: 64, top: 62, width: 1080, height: 62 }, { fontSize: 40, bold: true });
  box(slide, "", { left: 74, top: 154, width: 1132, height: 474 }, { fill: C.panel, line: "#E2E8F0" });

  modelStack(slide, 146, 270, C.purple, 5, "teacher", "large, expensive");
  line(slide, 360, 298, 470, 298, { color: C.line, width: 3, endArrowType: "triangle" });

  box(slide, "soft targets", { left: 486, top: 210, width: 284, height: 70 }, { fill: C.white, line: C.orange, lineWidth: 2.2, fontSize: 25, color: C.orange });
  addText(slide, "probabilities carry extra signal", { left: 458, top: 292, width: 340, height: 28 }, { fontSize: 18, color: C.muted, alignment: "center" });
  token(slide, "cat .62", 440, 365, C.orange);
  token(slide, "dog .21", 548, 365, C.orange);
  token(slide, "fox .09", 656, 365, C.orange);
  token(slide, "...", 764, 365, C.orange);

  line(slide, 824, 298, 918, 298, { color: C.line, width: 3, endArrowType: "triangle" });
  modelStack(slide, 942, 296, C.green, 3, "student", "smaller, faster");

  addText(slide, "learn the teacher's output distribution", { left: 346, top: 512, width: 590, height: 34 }, { fontSize: 27, bold: true, color: C.ink, alignment: "center" });
  circle(slide, 1010, 223, 9, C.green);
  circle(slide, 1040, 223, 9, C.green);
  circle(slide, 1070, 223, 9, C.green);
  addText(slide, "deployment", { left: 948, top: 190, width: 190, height: 26 }, { fontSize: 18, color: C.muted, alignment: "center" });
}

async function main() {
  const p = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  while (p.slides.items.length > 48) {
    p.slides.getItem(p.slides.items.length - 1).delete();
  }
  const slide = p.slides.add();
  drawSlide(slide);
  footer(slide, 49);
  slide.speakerNotes.textFrame.setText("[Sources]\nDistilling the Knowledge in a Neural Network: https://arxiv.org/abs/1503.02531\nDistilBERT, a distilled version of BERT: https://arxiv.org/abs/1910.01108");
  slide.speakerNotes.setVisible(true);
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log(PPTX);
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
