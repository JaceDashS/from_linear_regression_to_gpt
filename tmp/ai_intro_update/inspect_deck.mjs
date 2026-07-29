import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, PresentationFile } from "file:///C:/workspace/from_linear_regression_to_gpt/ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

async function saveBlob(blob, target) {
  await fs.writeFile(target, new Uint8Array(await blob.arrayBuffer()));
}

const root = "C:/workspace/from_linear_regression_to_gpt/tmp/ai_intro_update";
const pptxPath = path.join(root, "source.pptx");
const output = path.join(root, "template-inspect");
await fs.mkdir(path.join(output, "slides"), { recursive: true });
await fs.mkdir(path.join(output, "layouts"), { recursive: true });

const presentation = await PresentationFile.importPptx(await FileBlob.load(pptxPath));
const snapshot = await presentation.inspect({
  kind: "slide,textbox,shape,image,table,chart,notes,layout",
  maxChars: 200000,
});
await fs.writeFile(path.join(output, "template-inspect.ndjson"), snapshot.ndjson ?? "", "utf8");

for (const [index, slide] of presentation.slides.items.entries()) {
  const number = String(index + 1).padStart(2, "0");
  await saveBlob(
    await presentation.export({ slide, format: "png", scale: 1 }),
    path.join(output, "slides", `slide-${number}.png`),
  );
  await saveBlob(
    await presentation.export({ slide, format: "layout" }),
    path.join(output, "layouts", `slide-${number}.layout.json`),
  );
}

await fs.writeFile(
  path.join(output, "manifest.json"),
  JSON.stringify({ slideCount: presentation.slides.items.length }, null, 2),
  "utf8",
);
