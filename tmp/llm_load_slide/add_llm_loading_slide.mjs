import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading.pptx";
const OUT = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\llm_load_slide\\final-render";

const C = {
  ink: "#0B0F19",
  muted: "#667085",
  line: "#CBD5E1",
  panel: "#F8FAFC",
  blue: "#2563EB",
  blueFill: "#E0F2FE",
  green: "#10B981",
  greenFill: "#DCFCE7",
  orange: "#F97316",
  orangeFill: "#FEF3C7",
  white: "#FFFFFF",
};

function addText(slide, text, position, style = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name: style.name,
    position,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    fontSize: style.fontSize ?? 22,
    bold: style.bold ?? false,
    color: style.color ?? C.ink,
    alignment: style.alignment ?? "left",
    verticalAlignment: style.verticalAlignment ?? "middle",
  };
  return shape;
}

function component(slide, { x, fill, line, label, title, detail }) {
  const frame = slide.shapes.add({
    geometry: "roundRect",
    name: `${label.toLowerCase()}-component`,
    position: { left: x, top: 242, width: 300, height: 230 },
    fill,
    line: { style: "solid", fill: line, width: 2 },
    borderRadius: "rounded-xl",
  });
  addText(slide, label, { left: x + 24, top: 262, width: 252, height: 28 }, {
    name: `${label.toLowerCase()}-label`, fontSize: 15, bold: true, color: line, alignment: "center",
  });
  addText(slide, title, { left: x + 22, top: 306, width: 256, height: 58 }, {
    name: `${label.toLowerCase()}-title`, fontSize: 27, bold: true, alignment: "center",
  });
  addText(slide, detail, { left: x + 28, top: 378, width: 244, height: 68 }, {
    name: `${label.toLowerCase()}-detail`, fontSize: 18, color: C.muted, alignment: "center",
  });
  return frame;
}

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));

// The added GPT slide extends the final roadmap range by one page.
for (const shape of presentation.slides.getItem(0).shapes.items) {
  if (shape.text?.toString().trim() === "Slides 38-51") shape.text = "Slides 38-52";
}

const sourceSlide = presentation.slides.getItem(39);
const slide = sourceSlide.duplicate();
slide.moveTo(40);

// Preserve the inherited header, panel, and footer; replace only the source diagram.
const inherited = [...slide.shapes.items];
for (let index = 3; index < inherited.length - 1; index += 1) inherited[index].delete();

slide.shapes.items[0].text = "Loading";
slide.shapes.items[0].text.style = { fontSize: 16, bold: true, color: C.blue, alignment: "left" };
slide.shapes.items[1].text = "An LLM needs three matching parts";
slide.shapes.items[1].text.style = { fontSize: 42, bold: true, color: C.ink, alignment: "left" };

addText(slide, "model", { left: 146, top: 180, width: 604, height: 34 }, {
  name: "model-group-label", fontSize: 18, bold: true, color: C.blue, alignment: "center",
});

component(slide, {
  x: 108, fill: C.blueFill, line: C.blue, label: "CONFIG", title: "Architecture",
  detail: "layers · hidden size\nattention heads",
});
addText(slide, "+", { left: 414, top: 320, width: 48, height: 54 }, {
  name: "model-plus", fontSize: 34, bold: true, color: C.muted, alignment: "center",
});
component(slide, {
  x: 468, fill: C.greenFill, line: C.green, label: "CHECKPOINT", title: "Trained weights",
  detail: "learned parameters\nthat store patterns",
});
addText(slide, "+", { left: 774, top: 320, width: 48, height: 54 }, {
  name: "tokenizer-plus", fontSize: 34, bold: true, color: C.muted, alignment: "center",
});
component(slide, {
  x: 828, fill: C.orangeFill, line: C.orange, label: "TEXT ↔ IDS", title: "Tokenizer",
  detail: "vocabulary · split rules\nspecial tokens",
});

addText(slide, "The tokenizer must match the checkpoint — otherwise token IDs mean something else.",
  { left: 174, top: 530, width: 932, height: 44 }, {
    name: "matching-takeaway", fontSize: 23, bold: true, color: C.ink, alignment: "center",
  });

slide.speakerNotes.textFrame.setText([
  "[Sources]",
  "Hugging Face Transformers, Loading models: https://huggingface.co/docs/transformers/en/models",
  "Hugging Face Transformers, Auto Classes: https://huggingface.co/docs/transformers/model_doc/auto",
].join("\n"));
slide.speakerNotes.setVisible(true);

// Renumber the visible page marker after insertion.
for (let index = 0; index < presentation.slides.items.length; index += 1) {
  const current = presentation.slides.getItem(index);
  for (const shape of current.shapes.items) {
    const pos = shape.position;
    if (pos && pos.left >= 1140 && pos.top >= 640 && pos.width <= 80 && pos.height <= 40) {
      shape.text = String(index + 1).padStart(2, "0");
      shape.text.style = { fontSize: 13, color: C.muted, alignment: "right" };
    }
  }
}

await fs.mkdir(`${OUT}\\layouts`, { recursive: true });
for (let index = 0; index < presentation.slides.items.length; index += 1) {
  const current = presentation.slides.getItem(index);
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(`${OUT}\\${stem}.png`, await current.export({ format: "png", scale: 1 }));
  const layout = await current.export({ format: "layout" });
  await fs.writeFile(`${OUT}\\layouts\\${stem}.layout.json`, await layout.text(), "utf8");
}
await writeBlob(`${OUT}\\montage.webp`, await presentation.export({ format: "webp", montage: true, scale: 1 }));

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(FINAL);
console.log(`saved=${FINAL}`);
console.log(`slides=${presentation.slides.items.length}`);
