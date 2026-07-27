import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const STARTER = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_cnn_slide\\template-starter.pptx";
const IMAGE = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_cnn_slide\\assets\\cnn_banner.png";
const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed_cnn_added.pptx";
const TMP = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_cnn_slide";
const RENDER = `${TMP}\\final-render`;

const C = { ink: "#0B0F19", muted: "#667085", blue: "#2563EB" };

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(STARTER));
const slide = presentation.slides.getItem(18);

const inheritedShapes = [...slide.shapes.items];
if (inheritedShapes.length < 18) throw new Error(`Unexpected source-slide shape count: ${inheritedShapes.length}`);
for (let index = 3; index < inheritedShapes.length - 1; index += 1) inheritedShapes[index].delete();

slide.shapes.items[0].text = "Vision";
slide.shapes.items[0].text.style = { fontSize: 16, bold: true, color: C.blue, alignment: "left" };
slide.shapes.items[1].text = "CNNs reuse filters to find local patterns";
slide.shapes.items[1].text.style = { fontSize: 42, bold: true, color: C.ink, alignment: "left" };

const contentImage = slide.images.items.find((image) => image.frame?.width < 1000);
if (!contentImage) throw new Error("Could not locate the inherited content image on the duplicated slide.");
for (const image of [...slide.images.items]) {
  image.delete();
}
const bytes = await fs.readFile(IMAGE);
slide.images.add({
  blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  contentType: "image/png",
  alt: "CNN pipeline from image and shared kernels through feature maps, fully connected layers, and class probabilities",
  fit: "contain",
  position: { left: 172, top: 170, width: 936, height: 440 },
});

slide.speakerNotes.textFrame.setText([
  "[Sources]",
  "CNN overview and claims about local connectivity, feature maps, pooling, and parameter sharing: https://developersbreach.com/convolution-neural-network-deep-learning/",
  "CNN architecture image supplied by the user: https://i0.wp.com/developersbreach.com/wp-content/uploads/2020/08/cnn_banner.png?fit=1200%2C564&ssl=1",
].join("\n"));
slide.speakerNotes.setVisible(true);

const roadmapRanges = new Map([
  ["Slides 2-20", "Slides 2-21"],
  ["Slides 21-25", "Slides 22-26"],
  ["Slides 26-33", "Slides 27-34"],
  ["Slides 34-37", "Slides 35-38"],
  ["Slides 38-52", "Slides 39-53"],
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
