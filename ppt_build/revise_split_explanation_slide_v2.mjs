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
  orange: "#F97316",
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
    line: { style: opts.style ?? "solid", fill: opts.color ?? C.line, width: opts.width ?? 2, endArrowType: opts.endArrowType },
  });
}

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function card(slide, cfg) {
  box(slide, "", { left: cfg.x, top: 184, width: 314, height: 372 }, { fill: C.white, line: cfg.color, lineWidth: 2.2 });
  addText(slide, cfg.title, { left: cfg.x + 22, top: 212, width: 270, height: 36 }, { fontSize: 31, bold: true, color: cfg.color, alignment: "center" });
  addText(slide, cfg.role, { left: cfg.x + 24, top: 270, width: 266, height: 52 }, { fontSize: 23, bold: true, color: C.ink, alignment: "center" });

  const rows = [
    ["uses", cfg.uses],
    ["affects", cfg.affects],
    ["timing", cfg.timing],
  ];
  for (let i = 0; i < rows.length; i++) {
    const y = 354 + i * 54;
    addText(slide, rows[i][0], { left: cfg.x + 30, top: y, width: 74, height: 24 }, { fontSize: 16, bold: true, color: cfg.color });
    addText(slide, rows[i][1], { left: cfg.x + 106, top: y - 2, width: 178, height: 34 }, { fontSize: 17, color: C.ink });
    if (i < rows.length - 1) line(slide, cfg.x + 30, y + 38, cfg.x + 284, y + 38, { color: "#E2E8F0", width: 1 });
  }
}

async function rewriteSlide8(presentation) {
  const old = presentation.slides.getItem(7);
  old.delete();
  const slide = presentation.slides.insert({ after: presentation.slides.getItem(6) }).slide;
  slide.background.fill = C.white;
  addText(slide, "Regression", { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, "Data is split to separate learning from judging", { left: 64, top: 62, width: 1080, height: 62 }, { fontSize: 40, bold: true });
  box(slide, "", { left: 74, top: 154, width: 1132, height: 474 }, { fill: C.panel, line: "#E2E8F0" });

  card(slide, {
    x: 122,
    title: "train",
    color: C.blue,
    role: "learn the model",
    uses: "examples + labels",
    affects: "weights directly",
    timing: "every update",
  });
  card(slide, {
    x: 483,
    title: "validation",
    color: C.green,
    role: "choose the setup",
    uses: "held-out examples",
    affects: "choices indirectly",
    timing: "during development",
  });
  card(slide, {
    x: 844,
    title: "test",
    color: C.purple,
    role: "estimate final score",
    uses: "unseen examples",
    affects: "nothing",
    timing: "after all choices",
  });

  line(slide, 444, 370, 474, 370, { color: C.line, width: 3, endArrowType: "triangle" });
  line(slide, 805, 370, 835, 370, { color: C.line, width: 3, endArrowType: "triangle" });
  addText(slide, "validation is a development check; test is the final check", { left: 292, top: 580, width: 700, height: 30 }, { fontSize: 23, bold: true, color: C.orange, alignment: "center" });
  footer(slide, 8);
  slide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard train/validation/test model evaluation practice.\nTrain data is used to fit parameters. Validation data is used for model selection and hyperparameter tuning, so it can influence the development process. Test data is held out until the end to estimate final generalization.");
  slide.speakerNotes.setVisible(true);
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
