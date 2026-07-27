import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx";
const CANDIDATE = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/depth2-candidate.pptx";
const OUT = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/depth2-render";

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

function findText(slide, text) {
  const shape = slide.shapes.items.find((item) => item.text?.toString().trim() === text);
  if (!shape) throw new Error(`Text not found: ${text}`);
  return shape;
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const neuron = slideWithText(presentation, "A neuron is a tiny calculator");
const network = slideWithText(presentation, "A network stacks many neurons");
const chainRule = slideWithText(presentation, "The chain rule connects causes");

const example = chainRule.duplicate();
example.moveTo(17);

findText(example, "Calculus").text = "Neural networks";
findText(example, "The chain rule connects causes").text = "Two sigmoid layers create a nonlinear regressor";
findText(example, "z = wx + b").text = "h₁ = σ(w₁x+b₁)";
findText(example, "a = f(z)").text = "h₂ = σ(w₂h₁+b₂)";
findText(example, "prediction").text = "ŷ = w₃h₂+b₃";

const loss = findText(example, "loss");
loss.delete();
const finalArrow = example.shapes.items.find((shape) => {
  const box = shape.position;
  return box && box.left >= 930 && box.top >= 340 && box.top <= 380;
});
if (!finalArrow) throw new Error("Final connector not found");
finalArrow.delete();

for (const image of [...example.images.items]) image.delete();

const labels = [
  ["input", 160, 430, 150],
  ["hidden 1\nwidth = 1", 365, 430, 170],
  ["hidden 2\nwidth = 1", 570, 430, 170],
  ["linear output", 775, 430, 170],
];
for (const [text, left, top, width] of labels) {
  const label = example.shapes.add({
    geometry: "textbox",
    position: { left, top, width, height: 56 },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  label.text = text;
  label.text.style = { fontSize: 18, bold: true, color: "#667085", alignment: "center", verticalAlignment: "middle" };
}

const sigmoid = example.shapes.add({
  geometry: "textbox",
  position: { left: 225, top: 515, width: 420, height: 60 },
  fill: "none",
  line: { style: "solid", fill: "none", width: 0 },
});
sigmoid.text = "hidden activation:  σ(z) = 1 / (1 + e⁻ᶻ)";
sigmoid.text.style = { fontSize: 22, bold: true, color: "#2563EB", alignment: "center", verticalAlignment: "middle" };

const regression = example.shapes.add({
  geometry: "textbox",
  position: { left: 675, top: 515, width: 390, height: 60 },
  fill: "none",
  line: { style: "solid", fill: "none", width: 0 },
});
regression.text = "regression output stays linear";
regression.text.style = { fontSize: 22, bold: true, color: "#EF4444", alignment: "center", verticalAlignment: "middle" };

findText(example, "18").text = "17";
example.speakerNotes.textFrame.setText([
  "[Sources]",
  "Educational synthesis based on standard neural-network and regression concepts.",
  "Hidden layers use sigmoid activations; the output layer remains linear for unconstrained regression.",
].join("\n"));
example.speakerNotes.setVisible(true);

neuron.delete();
findText(network, "17").text = "16";

await fs.mkdir(OUT, { recursive: true });
for (let i = 15; i <= 17; i += 1) {
  const slide = presentation.slides.getItem(i);
  await writeBlob(`${OUT}/slide-${i + 1}.png`, await slide.export({ format: "png", scale: 1 }));
  await fs.writeFile(`${OUT}/slide-${i + 1}.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");
}

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(CANDIDATE);
console.log(JSON.stringify({ candidate: CANDIDATE, slides: presentation.slides.items.length }));
