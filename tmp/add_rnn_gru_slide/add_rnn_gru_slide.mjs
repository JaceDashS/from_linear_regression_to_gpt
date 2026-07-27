import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const STARTER = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_rnn_gru_slide\\template-starter.pptx";
const IMAGE = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_rnn_gru_slide\\assets\\rnn_to_gru.jpg";
const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed_cnn_added_rnn_gru_added.pptx";
const TMP = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_rnn_gru_slide";
const RENDER = `${TMP}\\final-render`;

const C = { ink: "#0B0F19", muted: "#667085", blue: "#2563EB" };

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(STARTER));
const slide = presentation.slides.getItem(26);
const inheritedShapes = [...slide.shapes.items];
if (inheritedShapes.length < 18) throw new Error(`Unexpected source-slide shape count: ${inheritedShapes.length}`);
const takeaway = inheritedShapes.at(-2);
for (let index = 3; index < inheritedShapes.length - 2; index += 1) inheritedShapes[index].delete();
for (const image of [...slide.images.items]) image.delete();

slide.shapes.items[0].text = "Memory";
slide.shapes.items[0].text.style = { fontSize: 16, bold: true, color: C.blue, alignment: "left" };
slide.shapes.items[1].text = "LSTM and GRU add gates to the RNN";
slide.shapes.items[1].text.style = { fontSize: 42, bold: true, color: C.ink, alignment: "left" };
takeaway.text = "Gates decide what to keep, update, and forget";
takeaway.text.style = { fontSize: 24, bold: true, color: C.ink, alignment: "center" };

const bytes = await fs.readFile(IMAGE);
slide.images.add({
  blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  contentType: "image/png",
  alt: "Side-by-side RNN, LSTM, and GRU cells showing tanh paths and memory gates",
  fit: "contain",
  position: { left: 110, top: 205, width: 1060, height: 307 },
});

slide.speakerNotes.textFrame.setText([
  "[Sources]",
  "RNN, LSTM, and GRU comparison image supplied by the user: https://media.licdn.com/dms/image/v2/D5612AQGC0XlthYQKjg/article-cover_image-shrink_600_2000/article-cover_image-shrink_600_2000/0/1710956751324?e=2147483647&v=beta&t=cwk4Mpnc9zFtGX-fM1mnDOrDxBociQdD0iY3m-ILPEY",
].join("\n"));
slide.speakerNotes.setVisible(true);

const roadmapRanges = new Map([
  ["Slides 22-26", "Slides 22-27"],
  ["Slides 27-34", "Slides 28-35"],
  ["Slides 35-38", "Slides 36-39"],
  ["Slides 39-53", "Slides 40-54"],
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
