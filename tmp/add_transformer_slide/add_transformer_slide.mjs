import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const STARTER = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_transformer_slide\\template-starter.pptx";
const IMAGE = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_transformer_slide\\assets\\transformer.png";
const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed_cnn_added_rnn_gru_added_seq2seq_added_transformer_added.pptx";
const TMP = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_transformer_slide";
const RENDER = `${TMP}\\final-render`;

const C = { ink: "#0B0F19", muted: "#667085", blue: "#2563EB" };

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(STARTER));
const slide = presentation.slides.getItem(36);
const inheritedShapes = [...slide.shapes.items];
if (inheritedShapes.length < 18) throw new Error(`Unexpected source-slide shape count: ${inheritedShapes.length}`);

const keep = new Set([4, 5, 6, 9, 10]);
for (let index = 3; index < inheritedShapes.length - 1; index += 1) {
  if (!keep.has(index)) inheritedShapes[index].delete();
}
for (const image of [...slide.images.items]) image.delete();

slide.shapes.items[0].text = "Architecture";
slide.shapes.items[0].text.style = { fontSize: 16, bold: true, color: C.blue, alignment: "left" };
slide.shapes.items[1].text = "The Transformer connects encoder and decoder stacks";
slide.shapes.items[1].text.style = { fontSize: 42, bold: true, color: C.ink, alignment: "left" };

const encoderTitle = inheritedShapes[5];
const encoderBody = inheritedShapes[6];
const decoderTitle = inheritedShapes[9];
const decoderBody = inheritedShapes[10];

encoderTitle.position = { left: 856, top: 242, width: 250, height: 34 };
encoderTitle.text = "Encoder";
encoderTitle.text.style = { fontSize: 26, bold: true, color: C.blue, alignment: "left" };
encoderBody.position = { left: 856, top: 282, width: 260, height: 72 };
encoderBody.text = "self-attention\nfeed-forward";
encoderBody.text.style = { fontSize: 21, color: C.muted, alignment: "left", verticalAlignment: "top" };

decoderTitle.position = { left: 856, top: 374, width: 250, height: 34 };
decoderTitle.text = "Decoder";
decoderTitle.text.style = { fontSize: 26, bold: true, color: C.blue, alignment: "left" };
decoderBody.position = { left: 856, top: 414, width: 270, height: 100 };
decoderBody.text = "masked self-attention\ncross-attention\nfeed-forward";
decoderBody.text.style = { fontSize: 20, color: C.muted, alignment: "left", verticalAlignment: "top" };

const bytes = await fs.readFile(IMAGE);
slide.images.add({
  blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  contentType: "image/png",
  alt: "Transformer architecture with encoder and decoder stacks, multi-head attention, feed-forward blocks, residual normalization, embeddings, and positional encoding",
  fit: "contain",
  position: { left: 240, top: 158, width: 350, height: 470 },
});

slide.speakerNotes.textFrame.setText([
  "[Sources]",
  "Transformer architecture image supplied by the user: https://qiita-user-contents.imgix.net/https%3A%2F%2Fqiita-image-store.s3.ap-northeast-1.amazonaws.com%2F0%2F146659%2F90f26267-2445-f2dc-b0c6-80f384054ae0.png?ixlib=rb-4.0.0&auto=format&gif-q=60&q=75&s=17811178f02900b58011bbcc4220f974",
].join("\n"));
slide.speakerNotes.setVisible(true);

const roadmapRanges = new Map([
  ["Slides 37-40", "Slides 37-41"],
  ["Slides 41-55", "Slides 42-56"],
]);
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
