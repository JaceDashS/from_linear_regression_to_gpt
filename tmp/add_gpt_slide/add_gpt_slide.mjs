import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const STARTER = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_gpt_slide\\template-starter.pptx";
const IMAGE = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_gpt_slide\\assets\\gpt.jpg";
const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed_cnn_added_rnn_gru_added_seq2seq_added_transformer_added_gpt_added.pptx";
const TMP = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_gpt_slide";
const RENDER = `${TMP}\\final-render`;
const C = { ink: "#0B0F19", muted: "#667085", blue: "#2563EB" };

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(STARTER));
const slide = presentation.slides.getItem(41);
const inheritedShapes = [...slide.shapes.items];
if (inheritedShapes.length < 18) throw new Error(`Unexpected source-slide shape count: ${inheritedShapes.length}`);

const keep = new Set([4, 5, 6, 9, 10]);
for (let index = 3; index < inheritedShapes.length - 1; index += 1) {
  if (!keep.has(index)) inheritedShapes[index].delete();
}
for (const image of [...slide.images.items]) image.delete();

inheritedShapes[0].text = "GPT architecture";
inheritedShapes[0].text.style = { fontSize: 16, bold: true, color: C.blue, alignment: "left" };
inheritedShapes[1].text = "GPT repeats a decoder-only Transformer block";
inheritedShapes[1].text.style = { fontSize: 42, bold: true, color: C.ink, alignment: "left" };

const stackTitle = inheritedShapes[5];
const stackBody = inheritedShapes[6];
const outputTitle = inheritedShapes[9];
const outputBody = inheritedShapes[10];

stackTitle.position = { left: 850, top: 242, width: 285, height: 34 };
stackTitle.text = "Decoder-only stack";
stackTitle.text.style = { fontSize: 26, bold: true, color: C.blue, alignment: "left" };
stackBody.position = { left: 850, top: 284, width: 300, height: 118 };
stackBody.text = "masked self-attention\nfeed-forward network\nresidual + LayerNorm";
stackBody.text.style = { fontSize: 20, color: C.muted, alignment: "left", verticalAlignment: "top" };

outputTitle.position = { left: 850, top: 430, width: 285, height: 34 };
outputTitle.text = "Next-token output";
outputTitle.text.style = { fontSize: 26, bold: true, color: C.blue, alignment: "left" };
outputBody.position = { left: 850, top: 472, width: 300, height: 72 };
outputBody.text = "Linear + Softmax turns the final state into a token distribution.";
outputBody.text.style = { fontSize: 19, color: C.muted, alignment: "left", verticalAlignment: "top" };

const bytes = await fs.readFile(IMAGE);
slide.images.add({
  blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  contentType: "image/jpeg",
  alt: "GPT decoder-only Transformer architecture showing input embeddings, positional encoding, repeated masked multi-head attention and feed-forward blocks, and next-token output",
  fit: "contain",
  position: { left: 235, top: 160, width: 430, height: 472 },
});

slide.speakerNotes.textFrame.setText([
  "[Sources]",
  "GPT architecture image supplied by the user: https://img.sbbit.jp/article/image/114721/l_bit202308301441497020.jpg",
].join("\n"));
slide.speakerNotes.setVisible(true);

const roadmapRanges = new Map([["Slides 42-56", "Slides 42-57"]]);
for (const shape of presentation.slides.getItem(0).shapes.items) {
  const current = shape.text?.toString().trim();
  if (roadmapRanges.has(current)) shape.text = roadmapRanges.get(current);
}

for (let index = 0; index < presentation.slides.items.length; index += 1) {
  const currentSlide = presentation.slides.getItem(index);
  for (const shape of currentSlide.shapes.items) {
    const pos = shape.position;
    if (pos && pos.left >= 1140 && pos.top >= 640 && pos.width <= 80 && pos.height <= 40) {
      shape.text = String(index + 1).padStart(2, "0");
      shape.text.style = { fontSize: 13, color: C.muted, alignment: "right" };
    }
  }
}

await fs.mkdir(`${RENDER}\\layouts`, { recursive: true });
for (let index = 0; index < presentation.slides.items.length; index += 1) {
  const currentSlide = presentation.slides.getItem(index);
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(`${RENDER}\\${stem}.png`, await currentSlide.export({ format: "png", scale: 1 }));
  await fs.writeFile(`${RENDER}\\layouts\\${stem}.layout.json`, await (await currentSlide.export({ format: "layout" })).text(), "utf8");
}

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(FINAL);
console.log(`slides=${presentation.slides.items.length}`);
console.log(`saved=${FINAL}`);
