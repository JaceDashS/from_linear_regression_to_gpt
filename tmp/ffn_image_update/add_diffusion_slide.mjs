import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx";
const IMAGE = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/diffusion-process.png";
const CANDIDATE = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/diffusion-candidate.pptx";
const OUT = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/diffusion-render";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

function findText(slide, text) {
  const shape = slide.shapes.items.find((item) => item.text?.toString().trim() === text);
  if (!shape) throw new Error(`Missing text: ${text}`);
  return shape;
}

function replacePageNumber(slide, oldNumber, newNumber) {
  const shape = slide.shapes.items.find((item) => item.text?.toString().trim() === String(oldNumber));
  if (!shape) throw new Error(`Missing page number ${oldNumber}`);
  shape.text = String(newNumber);
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const cnnSlide = presentation.slides.items.find((slide) =>
  slide.shapes.items.some((shape) => shape.text?.toString().trim() === "CNNs reuse filters to find local patterns"),
);
if (!cnnSlide) throw new Error("CNN source slide not found");

const diffusion = cnnSlide.duplicate();
diffusion.moveTo(54);

findText(diffusion, "Vision").text = "Diffusion";
findText(diffusion, "CNNs reuse filters to find local patterns").text = "Diffusion learns to reverse noise";
replacePageNumber(diffusion, 20, 55);

const sourceImage = diffusion.images.items.find((image) => image.position?.width > 900 && image.position?.height > 400);
if (!sourceImage) throw new Error("Source image frame not found");
const bytes = await fs.readFile(IMAGE);
sourceImage.replace({
  blob: bytes,
  contentType: "image/png",
  alt: "Cat image progressively corrupted by noise, with reverse denoising direction shown below",
  fit: "cover",
});
sourceImage.position = { left: 105, top: 178, width: 1070, height: 410 };
sourceImage.crop = { left: 0, top: 0.14, right: 0, bottom: 0.14 };

const forward = diffusion.shapes.add({
  geometry: "textbox",
  name: "diffusion-forward-label",
  position: { left: 165, top: 168, width: 360, height: 36 },
  fill: "none",
  line: { style: "solid", fill: "none", width: 0 },
});
forward.text = "Forward: gradually add noise →";
forward.text.style = { fontSize: 20, bold: true, color: "#667085", alignment: "left" };

const reverse = diffusion.shapes.add({
  geometry: "textbox",
  name: "diffusion-reverse-label",
  position: { left: 670, top: 574, width: 460, height: 40 },
  fill: "none",
  line: { style: "solid", fill: "none", width: 0 },
});
reverse.text = "← Generation: predict and remove noise step by step";
reverse.text.style = { fontSize: 20, bold: true, color: "#2563EB", alignment: "right" };

diffusion.speakerNotes.textFrame.setText([
  "[Sources]",
  "User-supplied diffusion image: https://miro.medium.com/v2/resize:fit:1400/1*yvClU5LgylulNWLPNpOCfA.png",
  "Ho et al., 2020, 'Denoising Diffusion Probabilistic Models': https://arxiv.org/abs/2006.11239",
].join("\n"));
diffusion.speakerNotes.setVisible(true);

for (const [oldNumber, newNumber] of [[55, 56], [56, 57], [57, 58], [58, 59]]) {
  const slide = presentation.slides.getItem(newNumber - 1);
  replacePageNumber(slide, oldNumber, newNumber);
}

await fs.mkdir(OUT, { recursive: true });
for (let i = 54; i <= 56; i += 1) {
  const slide = presentation.slides.getItem(i);
  await writeBlob(`${OUT}/slide-${i + 1}.png`, await slide.export({ format: "png", scale: 1 }));
  await fs.writeFile(`${OUT}/slide-${i + 1}.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");
}

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(CANDIDATE);
console.log(JSON.stringify({ candidate: CANDIDATE, slides: presentation.slides.items.length }));
