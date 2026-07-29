import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, PresentationFile } from "file:///C:/workspace/from_linear_regression_to_gpt/ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

async function saveBlob(blob, target) {
  await fs.writeFile(target, new Uint8Array(await blob.arrayBuffer()));
}

const root = "C:/workspace/from_linear_regression_to_gpt/tmp/ai_intro_update";
const output = path.join(root, "verification");
await fs.mkdir(output, { recursive: true });
const presentation = await PresentationFile.importPptx(await FileBlob.load(path.join(root, "modified.pptx")));
for (const index of [0, 1, 2]) {
  const slide = presentation.slides.getItem(index);
  await saveBlob(await presentation.export({ slide, format: "png", scale: 2 }), path.join(output, `slide-${index + 1}.png`));
  await saveBlob(await presentation.export({ slide, format: "layout" }), path.join(output, `slide-${index + 1}.layout.json`));
}
const snapshot = await presentation.inspect({ kind: "slide,textbox,shape,image,notes", maxChars: 24000 });
await fs.writeFile(path.join(output, "inspect.ndjson"), snapshot.ndjson ?? "", "utf8");
console.log(JSON.stringify({ slides: presentation.slides.items.length }));
