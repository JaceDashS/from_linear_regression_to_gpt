import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx";
const CANDIDATE = "C:/docs/from_linear_regression_to_gpt/tmp/late_architecture_update/candidate.pptx";
const RENDER_DIR = "C:/docs/from_linear_regression_to_gpt/tmp/late_architecture_update/render";

const C = { ink: "#0B0F19", muted: "#667085" };

function findText(slide, text) {
  const shape = slide.shapes.items.find((item) => item.text?.toString().trim() === text);
  if (!shape) throw new Error(`Missing inherited text: ${text}`);
  return shape;
}

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));

// Slide 55: make the dLLM / Mercury 2 architecture explicit in the existing rail.
const dllmSlide = presentation.slides.getItem(54);
findText(dllmSlide, "Diffusion language models (dLLMs) denoise tokens").text =
  "dLLMs refine many token positions in parallel";
findText(dllmSlide, "REAL-WORLD dLLM").text = "MERCURY 2";
findText(dllmSlide, "Mercury 2 ↗").text = "Transformer dLLM";
findText(dllmSlide, "many tokens\nrefined in parallel").text =
  "noisy sequence\n→ parallel predictions";
findText(dllmSlide, "not one token\nat a time").text =
  "repeat refinement\n→ final text";

dllmSlide.speakerNotes.textFrame.setText([
  "[Sources]",
  "Inception, Introducing Mercury 2: https://www.inceptionlabs.ai/blog/introducing-mercury-2",
  "Khanna et al., 2025, 'Mercury: Ultra-Fast Language Models Based on Diffusion': https://arxiv.org/abs/2506.17298",
].join("\n"));
dllmSlide.speakerNotes.setVisible(true);

// Slide 56: label the draft–verify cycle as an explicit inference structure.
const specSlide = presentation.slides.getItem(55);
findText(specSlide, "Speculative decoding drafts, then verifies").text =
  "Speculative decoding uses a draft-and-verify loop";
findText(specSlide, "small draft model").text = "1  small draft model";
findText(specSlide, "candidate tokens").text = "2  propose token block";
findText(specSlide, "large target model").text = "3  large target model";
findText(specSlide, "parallel check").text = "verify block in one pass";
findText(specSlide, "accept prefix").text = "4a  accept prefix";
findText(specSlide, "reject / resample").text = "4b  reject + resample";

const loop = findText(specSlide, "same target behavior, fewer serial waits");
loop.text = "draft → verify → accept / resample → repeat\nsame target distribution, fewer serial waits";
loop.position = { left: 300, top: 542, width: 680, height: 70 };
loop.text.style = {
  fontSize: 19,
  bold: true,
  color: C.ink,
  alignment: "center",
  verticalAlignment: "middle",
};

specSlide.speakerNotes.textFrame.setText([
  "[Sources]",
  "Leviathan, Kalman, and Matias, 2023, 'Fast Inference from Transformers via Speculative Decoding': https://arxiv.org/abs/2211.17192",
].join("\n"));
specSlide.speakerNotes.setVisible(true);

await fs.mkdir(RENDER_DIR, { recursive: true });
for (const slideNumber of [55, 56]) {
  const slide = presentation.slides.getItem(slideNumber - 1);
  await writeBlob(`${RENDER_DIR}/slide-${slideNumber}.png`, await slide.export({ format: "png", scale: 2 }));
  await fs.writeFile(
    `${RENDER_DIR}/slide-${slideNumber}.layout.json`,
    await (await slide.export({ format: "layout" })).text(),
    "utf8",
  );
}

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(CANDIDATE);
console.log(JSON.stringify({ candidate: CANDIDATE, slides: presentation.slides.items.length }));
