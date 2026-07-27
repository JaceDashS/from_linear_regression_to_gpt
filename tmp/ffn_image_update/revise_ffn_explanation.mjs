import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx";
const CANDIDATE = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/candidate-v2.pptx";
const OUT = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/render-v2";

const C = { ink: "#0B0F19", blue: "#2563EB", red: "#FF4D4F" };

function findText(slide, text) {
  const shape = slide.shapes.items.find((item) => item.text?.toString().trim() === text);
  if (!shape) throw new Error(`Missing inherited text: ${text}`);
  return shape;
}

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const slide = presentation.slides.getItem(53);

const title = findText(slide, "Some FFN neurons correlate with specific facts");
title.text = "Why FFNs help is still not fully settled";

const intro = findText(slide, "The FFN is a two-layer MLP after self-attention.");
intro.text = "Attention mixes tokens; the FFN transforms each token nonlinearly.";
intro.position = { left: 105, top: 174, width: 520, height: 64 };
intro.text.style = { fontSize: 21, bold: true, color: C.blue, alignment: "left", verticalAlignment: "middle" };

const known = findText(slide, "1  Hidden state enters the FFN");
known.text = "Known: FFNs detect and combine learned features.";
known.position = { left: 105, top: 258, width: 520, height: 68 };
known.text.style = { fontSize: 21, bold: true, color: C.ink, alignment: "left", verticalAlignment: "middle" };

const memory = findText(slide, "2  Intermediate neurons activate selectively");
memory.text = "Evidence: FFN weights can act like key–value parameter memory.";
memory.position = { left: 105, top: 348, width: 520, height: 72 };
memory.text.style = { fontSize: 19, bold: true, color: C.ink, alignment: "left", verticalAlignment: "middle" };

const ablation = findText(slide, "Evidence: activation correlates with fact expression");
ablation.text = "Ablations: reducing FFN capacity shifts work elsewhere and can cost more compute.";
ablation.position = { left: 105, top: 442, width: 520, height: 74 };
ablation.text.style = { fontSize: 19, bold: true, color: C.red, alignment: "left", verticalAlignment: "middle" };

const conclusion = findText(slide, "Caution: correlation does not prove that one fact lives in one neuron.");
conclusion.text = "Conclusion: the mechanism is clearer than the single cause of the performance gain.";
conclusion.position = { left: 105, top: 538, width: 520, height: 62 };
conclusion.text.style = { fontSize: 18, bold: true, color: C.red, alignment: "left", verticalAlignment: "middle" };

slide.speakerNotes.textFrame.setText([
  "[Sources]",
  "User-supplied FFN / knowledge-neuron figure: https://i0.wp.com/syncedreview.com/wp-content/uploads/2021/04/image-114.png?resize=652%2C863&ssl=1",
  "Dai et al., 2022, 'Knowledge Neurons in Pretrained Transformers': https://aclanthology.org/2022.acl-long.581/",
  "Smithline and Mascioli, 2026, 'Sparsity Moves Computation: How FFN Architecture Reshapes Attention in Small Transformers': https://arxiv.org/abs/2605.09403",
].join("\n"));
slide.speakerNotes.setVisible(true);

await fs.mkdir(OUT, { recursive: true });
await writeBlob(`${OUT}/slide-54.png`, await slide.export({ format: "png", scale: 2 }));
await fs.writeFile(`${OUT}/slide-54.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(CANDIDATE);
console.log(JSON.stringify({ candidate: CANDIDATE, slides: presentation.slides.items.length }));
