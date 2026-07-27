import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx";
const IMAGE = "C:/docs/from_linear_regression_to_gpt/tmp/linkedin_intro_image/source-image.jpg";
const FINAL = "C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt_with_ai_landscape.pptx";
const OUT = "C:/docs/from_linear_regression_to_gpt/tmp/linkedin_intro_image/final-render";
const MAP = "C:/docs/from_linear_regression_to_gpt/tmp/linkedin_intro_image/template-frame-map.json";
const SOURCE_URL = "https://media.licdn.com/dms/image/v2/D4D22AQEwD7fMcLTAdQ/feedshare-shrink_800/B4DZd9bUoRG8Ag-/0/1750156005031?e=2147483647&v=beta&t=dCR6tdHvBihBUouwX60e8hqjI_j5cD386Zc8Bszzw_M";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

function byExactText(slide, text) {
  return slide.shapes.items.find((shape) => shape.text?.toString().trim() === text);
}

function setText(slide, from, to) {
  const shape = byExactText(slide, from);
  if (!shape) throw new Error(`Missing inherited text shape: ${from}`);
  shape.text = to;
  return shape;
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const originalCount = presentation.slides.items.length;

// Keep the roadmap ranges accurate after inserting the orientation slide at page 2.
const roadmap = presentation.slides.getItem(0);
for (const [from, to] of [
  ["Slides 2-21", "Slides 3-22"],
  ["Slides 22-28", "Slides 23-29"],
  ["Slides 29-36", "Slides 30-37"],
  ["Slides 37-41", "Slides 38-42"],
  ["Slides 42-57", "Slides 43-58"],
]) setText(roadmap, from, to);

// Duplicate the deck's established portrait-image + sidebar composition.
const slide = presentation.slides.getItem(36).duplicate();
slide.moveTo(1);

setText(slide, "Architecture", "Orientation");
setText(slide, "The Transformer connects encoder and decoder stacks", "This deck follows one path through the AI landscape");
setText(slide, "Encoder", "Starting point");
setText(slide, "self-attention\nfeed-forward", "linear regression\n→ machine learning\n→ neural networks");
setText(slide, "Decoder", "Destination");
setText(slide, "masked self-attention\ncross-attention\nfeed-forward", "deep learning\n→ Transformers\n→ GPT");

const inheritedImage = slide.images.items[0];
if (!inheritedImage) throw new Error("The duplicated source slide has no replaceable image frame.");
// Imported duplicates can share the same package media asset. Remove only the
// duplicate's image element and add independent bytes so the source slide stays intact.
inheritedImage.delete();
const bytes = await fs.readFile(IMAGE);
const image = slide.images.add({
  blob: bytes,
  contentType: "image/jpeg",
  alt: "Nested overview of artificial intelligence, machine learning, neural networks, deep learning, and generative AI",
  fit: "contain",
  position: { left: 208, top: 158, width: 432, height: 470 },
  geometry: "rect",
});

slide.speakerNotes.textFrame.setText([
  "[Sources]",
  `AI landscape image supplied by the user: ${SOURCE_URL}`,
  "The image is presented as an orientation graphic; labels are reproduced from the supplied source.",
].join("\n"));
slide.speakerNotes.setVisible(true);

// Renumber only the small bottom-right page marker on every slide.
for (let index = 0; index < presentation.slides.items.length; index += 1) {
  const current = presentation.slides.getItem(index);
  for (const shape of current.shapes.items) {
    const pos = shape.position;
    if (pos && pos.left >= 1140 && pos.top >= 640 && pos.width <= 80 && pos.height <= 40) {
      shape.text = String(index + 1).padStart(2, "0");
    }
  }
}

// Record the exact source-slide mapping for the edited deck.
const outputSlides = [];
for (let outputSlide = 1; outputSlide <= originalCount + 1; outputSlide += 1) {
  if (outputSlide === 2) {
    outputSlides.push({
      outputSlide,
      sourceSlide: 37,
      narrativeRole: "orientation map",
      reuseMode: "duplicate-slide",
      editTargets: [
        { sourceElementId: "sh/apsrid4r", action: "rewrite" },
        { sourceElementId: "sh/d07qt0bi", action: "rewrite" },
        { sourceElementId: "im/21czytcf", action: "replace-and-reposition" },
        { sourceElementId: "sh/fyx4369o", action: "rewrite" },
        { sourceElementId: "sh/alkvylo3", action: "rewrite" },
        { sourceElementId: "sh/7mtwjmxw", action: "rewrite" },
        { sourceElementId: "sh/mlofm18v", action: "rewrite" },
        { sourceElementId: "sh/tkfuxofm", action: "rewrite" },
      ],
    });
  } else {
    outputSlides.push({
      outputSlide,
      sourceSlide: outputSlide === 1 ? 1 : outputSlide - 1,
      narrativeRole: outputSlide === 1 ? "roadmap" : "preserved source content",
      reuseMode: "duplicate-slide",
      editTargets: outputSlide === 1 ? [
        { sourceElementId: "sh/toryx8rm", action: "rewrite" },
        { sourceElementId: "sh/8vu9o3il", action: "rewrite" },
        { sourceElementId: "sh/7ehcbetg", action: "rewrite" },
        { sourceElementId: "sh/udsfupcf", action: "rewrite" },
        { sourceElementId: "sh/v6torutg", action: "rewrite" },
      ] : [],
    });
  }
}
await fs.writeFile(MAP, JSON.stringify({ outputSlides, omittedSourceSlides: [] }, null, 2), "utf8");

await fs.mkdir(`${OUT}/layouts`, { recursive: true });
for (let index = 0; index < presentation.slides.items.length; index += 1) {
  const current = presentation.slides.getItem(index);
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(`${OUT}/${stem}.png`, await current.export({ format: "png", scale: 1 }));
  await fs.writeFile(`${OUT}/layouts/${stem}.layout.json`, await (await current.export({ format: "layout" })).text(), "utf8");
}
await writeBlob(`${OUT}/montage.webp`, await presentation.export({ format: "webp", montage: true, scale: 0.5 }));

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(FINAL);
console.log(JSON.stringify({ final: FINAL, slides: presentation.slides.items.length, rendered: OUT }));
