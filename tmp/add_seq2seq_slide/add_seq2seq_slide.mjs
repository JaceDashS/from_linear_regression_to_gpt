import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const STARTER = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_seq2seq_slide\\template-starter.pptx";
const IMAGE = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_seq2seq_slide\\assets\\seq2seq.png";
const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed_cnn_added_rnn_gru_added_seq2seq_added.pptx";
const TMP = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_seq2seq_slide";
const RENDER = `${TMP}\\final-render`;

const C = { ink: "#0B0F19", muted: "#667085", blue: "#2563EB" };

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(STARTER));
const slide = presentation.slides.getItem(27);
const inheritedShapes = [...slide.shapes.items];
if (inheritedShapes.length < 18) throw new Error(`Unexpected source-slide shape count: ${inheritedShapes.length}`);
const takeaway = inheritedShapes.at(-2);
for (let index = 3; index < inheritedShapes.length - 2; index += 1) inheritedShapes[index].delete();
for (const image of [...slide.images.items]) image.delete();

slide.shapes.items[0].text = "Encoder-decoder";
slide.shapes.items[0].text.style = { fontSize: 16, bold: true, color: C.blue, alignment: "left" };
slide.shapes.items[1].text = "Seq2Seq compresses the input into one context";
slide.shapes.items[1].text.style = { fontSize: 42, bold: true, color: C.ink, alignment: "left" };
takeaway.text = "One context vector must carry the whole input";
takeaway.text.style = { fontSize: 24, bold: true, color: C.ink, alignment: "center" };

const bytes = await fs.readFile(IMAGE);
slide.images.add({
  blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  contentType: "image/png",
  alt: "Seq2Seq encoder recurrent cells compressing input tokens into a context passed to a recurrent decoder",
  fit: "contain",
  position: { left: 138, top: 175, width: 1004, height: 370 },
});

slide.speakerNotes.textFrame.setText([
  "[Sources]",
  "Seq2Seq encoder-decoder image supplied by the user: https://blog.octopt.com/wp-content/uploads/2020/04/Seq2Seq.png",
].join("\n"));
slide.speakerNotes.setVisible(true);

const roadmapRanges = new Map([
  ["Slides 22-27", "Slides 22-28"],
  ["Slides 28-35", "Slides 29-36"],
  ["Slides 36-39", "Slides 37-40"],
  ["Slides 40-54", "Slides 41-55"],
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
