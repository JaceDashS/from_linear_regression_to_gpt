import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, PresentationFile } from "file:///C:/workspace/from_linear_regression_to_gpt/ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const root = "C:/workspace/from_linear_regression_to_gpt/tmp/ai_agent_definition_update";
const sourcePptx = "C:/workspace/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx";
const outputPptx = path.join(root, "candidate.pptx");
const C = { ink: "#0B0F19", muted: "#667085", line: "#CBD5E1", blue: "#2563EB", green: "#10B981", orange: "#F97316", purple: "#7C3AED", white: "#FFFFFF" };

function addText(slide, value, position, style = {}) {
  const shape = slide.shapes.add({ geometry: "textbox", position, fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  shape.text = value;
  shape.text.style = { fontSize: style.fontSize ?? 20, bold: style.bold ?? false, color: style.color ?? C.ink, alignment: style.alignment ?? "left" };
  return shape;
}

function addNode(slide, value, position, color) {
  const shape = slide.shapes.add({ geometry: "roundRect", position, fill: C.white, line: { style: "solid", fill: color, width: 2.2 }, borderRadius: "rounded-xl" });
  shape.text = value;
  shape.text.style = { fontSize: 22, bold: true, color, alignment: "center" };
  return shape;
}

function addLine(slide, position, options = {}) {
  return slide.shapes.add({ geometry: "line", position, fill: "none", line: { style: options.style ?? "solid", fill: options.color ?? C.line, width: options.width ?? 2.5, beginArrowType: options.beginArrowType, endArrowType: options.endArrowType } });
}

function addDefinition(slide, label, description, left, top, color) {
  addText(slide, label, { left, top, width: 128, height: 26 }, { fontSize: 18, bold: true, color });
  addText(slide, description, { left: left + 132, top, width: 390, height: 28 }, { fontSize: 17, color: C.ink });
}

async function saveBlob(blob, target) {
  await fs.writeFile(target, new Uint8Array(await blob.arrayBuffer()));
}

await fs.mkdir(root, { recursive: true });
const presentation = await PresentationFile.importPptx(await FileBlob.load(sourcePptx));
const slide = presentation.slides.getItem(2);
const snapshot = await presentation.inspect({ kind: "textbox,shape,image", maxChars: 42000 });

for (const row of snapshot.ndjson.split(/\r?\n/)) {
  if (!row.trim()) continue;
  const item = JSON.parse(row);
  if (item.slide !== 3) continue;
  const bbox = item.bbox ?? [];
  const keepPanel = item.kind === "shape" && Math.abs((bbox[0] ?? 0) - 74) < 4 && Math.abs((bbox[1] ?? 0) - 164) < 4;
  const keepFooter = item.kind === "textbox" && Math.abs((bbox[0] ?? 0) - 1166) < 4 && Math.abs((bbox[1] ?? 0) - 660) < 8;
  if (!keepPanel && !keepFooter) presentation.resolve(item.id).delete();
}

addText(slide, "AI foundations", { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
addText(slide, "AI observes an environment and chooses actions", { left: 64, top: 62, width: 1120, height: 62 }, { fontSize: 40, bold: true });
addText(slide, "An intelligent-agent model", { left: 64, top: 122, width: 420, height: 28 }, { fontSize: 20, color: C.muted });
addText(slide, "An AI agent observes an environment and selects actions through a policy\nto achieve an objective.", { left: 125, top: 176, width: 1030, height: 54 }, { fontSize: 24, bold: true, alignment: "center" });

// Connectors are created before nodes so arrows remain behind the entities.
addLine(slide, { left: 300, top: 335, width: 48, height: 1 }, { endArrowType: "triangle" });
addLine(slide, { left: 538, top: 335, width: 48, height: 1 }, { endArrowType: "triangle" });
addLine(slide, { left: 776, top: 335, width: 48, height: 1 }, { endArrowType: "triangle" });
addLine(slide, { left: 1014, top: 335, width: 60, height: 1 }, { endArrowType: "triangle" });
addLine(slide, { left: 642, top: 286, width: 1, height: 20 }, { endArrowType: "triangle", color: C.purple });
addLine(slide, { left: 1074, top: 335, width: 1, height: 60 }, { color: C.line, width: 2 });
addLine(slide, { left: 205, top: 395, width: 869, height: 1 }, { beginArrowType: "triangle", color: C.line, width: 2 });
addLine(slide, { left: 205, top: 335, width: 1, height: 60 }, { color: C.line, width: 2 });

addNode(slide, "Environment", { left: 110, top: 306, width: 190, height: 58 }, C.blue);
addNode(slide, "Observation", { left: 348, top: 306, width: 190, height: 58 }, C.green);
addNode(slide, "Policy", { left: 586, top: 306, width: 190, height: 58 }, C.orange);
addNode(slide, "Action", { left: 824, top: 306, width: 190, height: 58 }, C.purple);
addNode(slide, "Objective", { left: 548, top: 242, width: 190, height: 44 }, C.purple);
addText(slide, "action changes the environment", { left: 455, top: 397, width: 370, height: 22 }, { fontSize: 15, color: C.muted, alignment: "center" });

addDefinition(slide, "Environment", "the world the AI interacts with", 110, 444, C.blue);
addDefinition(slide, "Observation", "information it receives: sensors, images, text", 110, 482, C.green);
addDefinition(slide, "Objective", "the criterion for a good result", 110, 520, C.purple);
addDefinition(slide, "Policy", "the rule or model that selects an action", 652, 444, C.orange);
addDefinition(slide, "Action", "the output: movement, classification, response", 652, 490, C.purple);

slide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on the conventional intelligent-agent formulation: environment, observation, objective, policy, and action.");
slide.speakerNotes.setVisible(true);

await fs.writeFile(path.join(root, "source-notes.txt"), "No external visual assets were added. The slide uses the conventional intelligent-agent formulation: environment, observation, objective, policy, and action.\n", "utf8");
await fs.writeFile(path.join(root, "deviation-log.txt"), "Slide 3: replaced the earlier short definition with the five-component intelligent-agent formulation requested by the user. The inherited background, main panel, and page marker were preserved.\n", "utf8");
await saveBlob(await presentation.export({ slide, format: "png", scale: 2 }), path.join(root, "slide-3.png"));
await saveBlob(await presentation.export({ slide, format: "layout" }), path.join(root, "slide-3.layout.json"));
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(outputPptx);
