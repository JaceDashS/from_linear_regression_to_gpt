import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/reordered-normalized.pptx";
const OUT = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/reordered-normalized-verified";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
await fs.mkdir(OUT, { recursive: true });

const expected = [
  "A neuron is a tiny calculator",
  "A network stacks many neurons",
  "The chain rule connects causes",
  "Matrices move many numbers at once",
  "CNNs reuse filters to find local patterns",
  "Backprop sends error backward",
];

for (let offset = 0; offset < expected.length; offset += 1) {
  const index = 15 + offset;
  const slide = presentation.slides.getItem(index);
  const texts = slide.shapes.items.map((shape) => shape.text?.toString().trim()).filter(Boolean);
  if (!texts.includes(expected[offset])) throw new Error(`Unexpected slide ${index + 1}: ${texts.join(" | ")}`);
  if (!texts.includes(String(index + 1))) throw new Error(`Wrong page marker on slide ${index + 1}`);
  await writeBlob(`${OUT}/slide-${index + 1}.png`, await slide.export({ format: "png", scale: 1 }));
  await fs.writeFile(`${OUT}/slide-${index + 1}.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");
}

console.log(JSON.stringify({ slides: presentation.slides.items.length, verified: expected }));
