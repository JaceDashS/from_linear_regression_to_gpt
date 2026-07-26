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

function drawCard(slide, cfg) {
  box(slide, "", { left: cfg.x, top: 186, width: 316, height: 326 }, { fill: C.white, line: cfg.color, lineWidth: 2.2 });
  addText(slide, cfg.title, { left: cfg.x + 24, top: 216, width: 268, height: 38 }, { fontSize: 31, bold: true, color: cfg.color, alignment: "center" });
  addText(slide, cfg.when, { left: cfg.x + 28, top: 286, width: 260, height: 34 }, { fontSize: 22, bold: true, color: C.ink, alignment: "center" });
  addText(slide, cfg.job, { left: cfg.x + 36, top: 338, width: 244, height: 62 }, { fontSize: 20, color: C.ink, alignment: "center" });
  addText(slide, cfg.rule, { left: cfg.x + 28, top: 438, width: 260, height: 34 }, { fontSize: 18, color: C.muted, alignment: "center" });
}

async function rewriteSlide8(presentation) {
  const old = presentation.slides.getItem(7);
  old.delete();
  const slide = presentation.slides.insert({ after: presentation.slides.getItem(6) }).slide;
  slide.background.fill = C.white;
  addText(slide, "Regression", { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, "Train, validation, and test have different jobs", { left: 64, top: 62, width: 1080, height: 62 }, { fontSize: 40, bold: true });
  box(slide, "", { left: 74, top: 154, width: 1132, height: 474 }, { fill: C.panel, line: "#E2E8F0" });

  drawCard(slide, {
    x: 126,
    color: C.blue,
    title: "train",
    when: "during learning",
    job: "update weights\nusing loss",
    rule: "model can see this",
  });
  drawCard(slide, {
    x: 482,
    color: C.green,
    title: "validation",
    when: "during selection",
    job: "tune choices\ncompare models",
    rule: "guides decisions",
  });
  drawCard(slide, {
    x: 838,
    color: C.purple,
    title: "test",
    when: "after decisions",
    job: "estimate real-world\nperformance",
    rule: "use once at the end",
  });

  line(slide, 446, 348, 474, 348, { color: C.line, width: 3, endArrowType: "triangle" });
  line(slide, 802, 348, 830, 348, { color: C.line, width: 3, endArrowType: "triangle" });
  addText(slide, "Do not train on the exam", { left: 392, top: 552, width: 500, height: 36 }, { fontSize: 28, bold: true, color: C.red, alignment: "center" });
  footer(slide, 8);
  slide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard train/validation/test model evaluation practice.\nTrain set updates model weights. Validation set supports model selection and hyperparameter decisions. Test set is held out for the final estimate.");
  slide.speakerNotes.setVisible(true);
}

async function renumberFooters(presentation) {
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 300000 });
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
  await rewriteSlide8(p);
  await renumberFooters(p);
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log(PPTX);
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
