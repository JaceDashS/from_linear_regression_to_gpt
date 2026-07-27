import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const source = "C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt_with_ai_landscape.pptx";
const outDir = "C:/docs/from_linear_regression_to_gpt/tmp/linkedin_intro_image/verified-final";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
await fs.mkdir(`${outDir}/layouts`, { recursive: true });
for (const [index, slide] of presentation.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(`${outDir}/${stem}.png`, await slide.export({ format: "png", scale: 1 }));
  await fs.writeFile(`${outDir}/layouts/${stem}.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");
}
const snapshot = await presentation.inspect({ kind: "slide,textbox,image,notes,layout", maxChars: 100000 });
await fs.writeFile(`${outDir}/inspect.ndjson`, snapshot.ndjson, "utf8");
console.log(JSON.stringify({ slides: presentation.slides.items.length, outDir }));
