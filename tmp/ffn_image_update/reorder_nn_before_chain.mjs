import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx";
const CANDIDATE = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/reordered-candidate.pptx";
const OUT = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/reordered-render";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

function slideWithText(presentation, text) {
  const slide = presentation.slides.items.find((item) =>
    item.shapes.items.some((shape) => shape.text?.toString().trim() === text),
  );
  if (!slide) throw new Error(`Slide not found: ${text}`);
  return slide;
}

function replacePageNumber(slide, oldNumber, newNumber) {
  const shape = slide.shapes.items.find((item) => item.text?.toString().trim() === String(oldNumber));
  if (!shape) throw new Error(`Page number ${oldNumber} not found`);
  shape.text = String(newNumber);
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));

const neuron = slideWithText(presentation, "A neuron is a tiny calculator");
const network = slideWithText(presentation, "A network stacks many neurons");
const chainRule = slideWithText(presentation, "The chain rule connects causes");
const matrices = slideWithText(presentation, "Matrices move many numbers at once");

neuron.moveTo(15);
network.moveTo(16);

replacePageNumber(neuron, 18, 16);
replacePageNumber(network, 19, 17);
replacePageNumber(chainRule, 16, 18);
replacePageNumber(matrices, 17, 19);

await fs.mkdir(OUT, { recursive: true });
for (let i = 15; i <= 20; i += 1) {
  const slide = presentation.slides.getItem(i);
  const stem = `slide-${i + 1}`;
  await writeBlob(`${OUT}/${stem}.png`, await slide.export({ format: "png", scale: 1 }));
  await fs.writeFile(`${OUT}/${stem}.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");
}

const order = presentation.slides.items.slice(14, 21).map((slide, index) => ({
  slide: index + 15,
  title: slide.shapes.items.find((shape) => shape.position?.top === 62)?.text?.toString().trim() ?? "",
}));
await fs.writeFile(`${OUT}/order.json`, JSON.stringify(order, null, 2), "utf8");

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(CANDIDATE);
console.log(JSON.stringify({ candidate: CANDIDATE, slides: presentation.slides.items.length, order }));
