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

function arrow(slide, x, y, w, h, text = "", opts = {}) {
  const s = slide.shapes.add({
    geometry: "rightArrow",
    position: { left: x, top: y, width: w, height: h },
    fill: opts.fill ?? "#E0F2FE",
    line: { style: "solid", fill: opts.line ?? C.blue, width: 1.5 },
  });
  if (text) {
    s.text = text;
    s.text.style = { fontSize: opts.fontSize ?? 18, bold: true, color: opts.color ?? C.ink, alignment: "center" };
  }
  return s;
}

function frame(slide) {
  box(slide, "", { left: 64, top: 142, width: 1152, height: 500 }, { fill: C.panel, line: "#E2E8F0" });
}

function title(slide, heading, kicker) {
  addText(slide, kicker, { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, heading, { left: 64, top: 62, width: 1040, height: 62 }, { fontSize: 42, bold: true });
}

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function notes(slide) {
  slide.speakerNotes.textFrame.setText([
    "[Sources]",
    "Geva et al., 2021, 'Transformer Feed-Forward Layers Are Key-Value Memories', arXiv:2012.14913, https://arxiv.org/abs/2012.14913",
    "Ndubuaku et al., 2026, 'A Controlled Study of Attention-Only Transformers', arXiv:2607.18363v1, https://arxiv.org/html/2607.18363v1",
  ].join("\n"));
  slide.speakerNotes.setVisible(true);
}

function addSlide(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  title(slide, "FFN helps, but the full reason is not settled", "Transformer");
  frame(slide);

  box(slide, "attention", { left: 135, top: 310, width: 170, height: 78 }, { fill: "#E0F2FE", line: C.blue, color: C.blue, fontSize: 26 });
  arrow(slide, 330, 328, 110, 42, "mix");
  box(slide, "FFN", { left: 465, top: 285, width: 190, height: 128 }, { fill: "#FEF3C7", line: C.orange, color: C.orange, fontSize: 38 });
  arrow(slide, 680, 328, 110, 42, "transform");
  box(slide, "features", { left: 815, top: 310, width: 170, height: 78 }, { fill: "#DCFCE7", line: C.green, color: C.green, fontSize: 26 });

  box(slide, "nonlinear feature transform", { left: 175, top: 470, width: 275, height: 68 }, { fill: C.white, line: C.line, fontSize: 22 });
  box(slide, "key-value memory evidence", { left: 505, top: 470, width: 275, height: 68 }, { fill: C.white, line: C.line, fontSize: 22 });
  box(slide, "not a single settled cause", { left: 835, top: 470, width: 275, height: 68 }, { fill: "#FEF2F2", line: C.red, color: C.red, fontSize: 22 });

  addText(slide, "Known: FFNs store and combine useful patterns.", { left: 220, top: 190, width: 840, height: 34 }, { fontSize: 27, bold: true, color: C.blue, alignment: "center" });
  addText(slide, "Open question: how much is form, capacity, or compute efficiency?", { left: 170, top: 580, width: 940, height: 34 }, { fontSize: 25, bold: true, color: C.red, alignment: "center" });
  footer(slide, presentation.slides.items.length);
  notes(slide);
}

async function main() {
  const presentation = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  addSlide(presentation);
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(PPTX);
  console.log(PPTX);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
