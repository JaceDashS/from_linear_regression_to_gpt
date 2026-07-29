import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, PresentationFile } from "file:///C:/workspace/from_linear_regression_to_gpt/ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const root = "C:/workspace/from_linear_regression_to_gpt/tmp/ai_definition_update";
const sourcePptx = "C:/workspace/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx";
const outputPptx = path.join(root, "candidate.pptx");
const C = { ink: "#0B0F19", muted: "#667085", line: "#CBD5E1", blue: "#2563EB", green: "#10B981", orange: "#F97316", white: "#FFFFFF" };

function addText(slide, value, position, style = {}) {
  const shape = slide.shapes.add({ geometry: "textbox", position, fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  shape.text = value;
  shape.text.style = { fontSize: style.fontSize ?? 24, bold: style.bold ?? false, color: style.color ?? C.ink, alignment: style.alignment ?? "left" };
  return shape;
}

function addBox(slide, value, position, color) {
  const shape = slide.shapes.add({ geometry: "roundRect", position, fill: C.white, line: { style: "solid", fill: color, width: 2 }, borderRadius: "rounded-xl" });
  shape.text = value;
  shape.text.style = { fontSize: 24, bold: true, color, alignment: "center" };
  return shape;
}

function addArrow(slide, x1, x2, y) {
  return slide.shapes.add({ geometry: "line", position: { left: x1, top: y, width: x2 - x1, height: 1 }, fill: "none", line: { style: "solid", fill: C.line, width: 3, endArrowType: "triangle" } });
}

async function shiftRoadmapRanges(presentation) {
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 50000 });
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    const value = item.text ?? "";
    if (item.slide !== 1 || !/^Slides \d+-\d+$/.test(value)) continue;
    const [start, end] = value.match(/\d+/g).map(Number);
    presentation.resolve(item.id).text = `Slides ${start + 1}-${end + 1}`;
  }
}

async function renumberFooters(presentation) {
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 300000 });
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    const bbox = item.bbox ?? [];
    if (bbox.length !== 4 || Math.abs(bbox[0] - 1166) >= 4 || Math.abs(bbox[1] - 660) >= 8 || !item.slide) continue;
    const footer = presentation.resolve(item.id);
    footer.text = String(item.slide).padStart(2, "0");
    footer.text.style = { fontSize: 13, color: C.muted, alignment: "right" };
  }
}

async function saveBlob(blob, target) {
  await fs.writeFile(target, new Uint8Array(await blob.arrayBuffer()));
}

await fs.mkdir(root, { recursive: true });
const outputSlides = [
  { outputSlide: 1, sourceSlide: 1, narrativeRole: "roadmap", reuseMode: "preserve", editTargets: ["rewrite page-range labels"] },
  { outputSlide: 2, sourceSlide: 2, narrativeRole: "AI examples", reuseMode: "preserve", editTargets: ["rewrite footer page number"] },
  { outputSlide: 3, sourceSlide: 2, narrativeRole: "AI definition", reuseMode: "duplicate-slide", editTargets: ["replace local image progression with concise definition flow"] },
  ...Array.from({ length: 69 }, (_, index) => ({ outputSlide: index + 4, sourceSlide: index + 3, narrativeRole: "preserve existing narrative", reuseMode: "preserve", editTargets: ["rewrite footer page number"] })),
];
await fs.writeFile(path.join(root, "template-frame-map.json"), `${JSON.stringify({ outputSlides, omittedSourceSlides: [] }, null, 2)}\n`, "utf8");
const presentation = await PresentationFile.importPptx(await FileBlob.load(sourcePptx));
const definition = presentation.slides.getItem(1).duplicate();
definition.moveTo(2);

const local = await presentation.inspect({ kind: "textbox,shape,image", maxChars: 30000 });
for (const row of local.ndjson.split(/\r?\n/)) {
  if (!row.trim()) continue;
  const item = JSON.parse(row);
  if (item.slide !== 3) continue;
  const bbox = item.bbox ?? [];
  const isTopText = item.kind === "textbox" && bbox[1] < 150;
  const isFooter = item.kind === "textbox" && Math.abs((bbox[0] ?? 0) - 1166) < 4 && Math.abs((bbox[1] ?? 0) - 660) < 8;
  const isMainPanel = item.kind === "shape" && Math.abs((bbox[0] ?? 0) - 74) < 4 && Math.abs((bbox[1] ?? 0) - 164) < 4;
  if (!isTopText && !isFooter && !isMainPanel) presentation.resolve(item.id).delete();
}

addText(definition, "AI foundations", { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
addText(definition, "What is AI?", { left: 64, top: 62, width: 1080, height: 62 }, { fontSize: 40, bold: true });
addText(definition, "A simple definition", { left: 64, top: 122, width: 420, height: 28 }, { fontSize: 20, color: C.muted });
addText(definition, "AI enables computers to learn patterns from data\nand apply them to new situations.", { left: 140, top: 224, width: 1000, height: 104 }, { fontSize: 34, bold: true, alignment: "center" });
addArrow(definition, 390, 500, 430);
addArrow(definition, 780, 890, 430);
addBox(definition, "Data", { left: 200, top: 382, width: 190, height: 96 }, C.blue);
addBox(definition, "Learn patterns", { left: 500, top: 382, width: 280, height: 96 }, C.green);
addBox(definition, "Predict · Decide · Create", { left: 890, top: 382, width: 250, height: 96 }, C.orange);
addText(definition, "AI does not simply follow one fixed rule—it generalizes from examples.", { left: 230, top: 536, width: 820, height: 34 }, { fontSize: 22, color: C.muted, alignment: "center" });
addText(definition, "03", { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
definition.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis: AI systems learn patterns from data and apply them to prediction, decision-making, or generation.");
definition.speakerNotes.setVisible(true);

await shiftRoadmapRanges(presentation);
await renumberFooters(presentation);
await saveBlob(await presentation.export({ slide: definition, format: "png", scale: 2 }), path.join(root, "slide-3.png"));
await saveBlob(await presentation.export({ slide: definition, format: "layout" }), path.join(root, "slide-3.layout.json"));
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(outputPptx);
