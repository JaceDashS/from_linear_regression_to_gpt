import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx";
const CANDIDATE = "C:/docs/from_linear_regression_to_gpt/tmp/attention_dot_product_update/candidate.pptx";
const RENDER_DIR = "C:/docs/from_linear_regression_to_gpt/tmp/attention_dot_product_update/render";

const C = {
  ink: "#0B0F19",
  muted: "#667085",
  blue: "#2563EB",
};

function findText(slide, text) {
  const shape = slide.shapes.items.find((item) => item.text?.toString().trim() === text);
  if (!shape) throw new Error(`Missing inherited text: ${text}`);
  return shape;
}

function addText(slide, text, position, style = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name: style.name,
    position,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    fontSize: style.fontSize ?? 22,
    bold: style.bold ?? false,
    color: style.color ?? C.ink,
    alignment: style.alignment ?? "center",
    verticalAlignment: style.verticalAlignment ?? "middle",
  };
  return shape;
}

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));

// Slide 32: explain the dot product and the first cross-attention query.
const dotSlide = presentation.slides.getItem(31);
findText(dotSlide, "Scores come from dot products").text = "Dot product turns Q and K into scores";
const dotLabel = findText(dotSlide, "dot");
dotLabel.text = "·";
dotLabel.text.style = { fontSize: 38, bold: true, color: C.ink, alignment: "center" };
findText(dotSlide, "relevance").text = "one score per key";

addText(dotSlide, "score = QKᵀ", { left: 220, top: 178, width: 780, height: 52 }, {
  name: "qk-score-formula",
  fontSize: 32,
  bold: true,
  color: C.blue,
});
addText(
  dotSlide,
  "Dot product: multiply matching dimensions, then add them.\n<BOS> query · each encoder key → one score for each source token",
  { left: 170, top: 510, width: 940, height: 78 },
  { name: "dot-product-explanation", fontSize: 21, color: C.muted },
);

// Slide 33: show scaling, Softmax, and the complete Attention expression.
const softmaxSlide = presentation.slides.getItem(32);
findText(softmaxSlide, "Softmax makes weights").text = "Scale scores, then Softmax makes weights";
const softmaxArrow = findText(softmaxSlide, "softmax");
softmaxArrow.text = "scale + softmax";
softmaxArrow.text.style = { fontSize: 17, bold: true, color: C.ink, alignment: "center" };
findText(softmaxSlide, "raw scores").text = "QKᵀ / √dₖ";
findText(softmaxSlide, "weights sum to 1").text = "attention weights (sum = 1)";

addText(
  softmaxSlide,
  "Attention(Q, K, V) = Softmax(QKᵀ / √dₖ)V",
  { left: 150, top: 178, width: 980, height: 52 },
  { name: "scaled-dot-product-attention-formula", fontSize: 30, bold: true, color: C.blue },
);
addText(
  softmaxSlide,
  "Dividing by √dₖ keeps large vectors from making the scores too extreme.",
  { left: 210, top: 520, width: 860, height: 48 },
  { name: "scaling-explanation", fontSize: 20, color: C.muted },
);

await fs.mkdir(RENDER_DIR, { recursive: true });
for (const slideNumber of [32, 33]) {
  const slide = presentation.slides.getItem(slideNumber - 1);
  await writeBlob(`${RENDER_DIR}/slide-${slideNumber}.png`, await slide.export({ format: "png", scale: 2 }));
  await fs.writeFile(
    `${RENDER_DIR}/slide-${slideNumber}.layout.json`,
    await (await slide.export({ format: "layout" })).text(),
    "utf8",
  );
}

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(CANDIDATE);
console.log(JSON.stringify({ candidate: CANDIDATE, slides: presentation.slides.items.length }));
