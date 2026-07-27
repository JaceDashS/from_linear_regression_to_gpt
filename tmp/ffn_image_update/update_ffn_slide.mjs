import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx";
const IMAGE = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/ffn-source.png";
const CANDIDATE = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/candidate.pptx";
const OUT = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/render";
const IMAGE_URL = "https://i0.wp.com/syncedreview.com/wp-content/uploads/2021/04/image-114.png?resize=652%2C863&ssl=1";

const C = {
  ink: "#0B0F19",
  muted: "#667085",
  blue: "#2563EB",
  red: "#FF4D4F",
};

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

findText(slide, "FFN helps, but the full reason is not settled").text =
  "Some FFN neurons correlate with specific facts";

// Remove the inherited abstract pipeline; keep and repurpose its explanatory slots.
for (const text of ["attention", "mix", "FFN", "transform", "features"]) {
  findText(slide, text).delete();
}

const intro = findText(slide, "Known: FFNs store and combine useful patterns.");
intro.text = "The FFN is a two-layer MLP after self-attention.";
intro.position = { left: 105, top: 180, width: 520, height: 48 };
intro.text.style = { fontSize: 24, bold: true, color: C.blue, alignment: "left" };

const step1 = findText(slide, "nonlinear feature transform");
step1.text = "1  Hidden state enters the FFN";
step1.position = { left: 105, top: 260, width: 520, height: 68 };
step1.text.style = { fontSize: 22, bold: true, color: C.ink, alignment: "left", verticalAlignment: "middle" };

const step2 = findText(slide, "key-value memory evidence");
step2.text = "2  Intermediate neurons activate selectively";
step2.position = { left: 105, top: 350, width: 520, height: 68 };
step2.text.style = { fontSize: 22, bold: true, color: C.ink, alignment: "left", verticalAlignment: "middle" };

const evidence = findText(slide, "not a single settled cause");
evidence.text = "Evidence: activation correlates with fact expression";
evidence.position = { left: 105, top: 440, width: 520, height: 72 };
evidence.text.style = { fontSize: 21, bold: true, color: C.red, alignment: "left", verticalAlignment: "middle" };

const caution = findText(slide, "Open question: how much is form, capacity, or compute efficiency?");
caution.text = "Caution: correlation does not prove that one fact lives in one neuron.";
caution.position = { left: 105, top: 545, width: 520, height: 48 };
caution.text.style = { fontSize: 19, bold: true, color: C.red, alignment: "left", verticalAlignment: "middle" };

const bytes = await fs.readFile(IMAGE);
slide.images.add({
  blob: bytes,
  contentType: "image/png",
  alt: "Knowledge-neuron attribution diagram showing a feed-forward network after self-attention",
  fit: "contain",
  crop: { left: 0, top: 0, right: 0, bottom: 0.11 },
  position: { left: 690, top: 162, width: 455, height: 460 },
});

slide.speakerNotes.textFrame.setText([
  "[Sources]",
  `User-supplied FFN / knowledge-neuron figure: ${IMAGE_URL}`,
  "Dai et al., 2022, 'Knowledge Neurons in Pretrained Transformers': https://aclanthology.org/2022.acl-long.581/",
].join("\n"));
slide.speakerNotes.setVisible(true);

await fs.mkdir(OUT, { recursive: true });
await writeBlob(`${OUT}/slide-54.png`, await slide.export({ format: "png", scale: 2 }));
await fs.writeFile(`${OUT}/slide-54.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(CANDIDATE);
console.log(JSON.stringify({ candidate: CANDIDATE, slides: presentation.slides.items.length }));
