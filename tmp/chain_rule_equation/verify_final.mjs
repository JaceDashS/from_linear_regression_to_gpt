import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation.pptx";
const OUT = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\chain_rule_equation\\verified";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

await fs.mkdir(OUT, { recursive: true });
const presentation = await PresentationFile.importPptx(await FileBlob.load(FINAL));
const slide = presentation.slides.getItem(14);
await writeBlob(`${OUT}\\slide-15.png`, await slide.export({ format: "png", scale: 2 }));
await fs.writeFile(`${OUT}\\slide-15.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");

const textboxes = await presentation.inspect({ kind: "textbox", search: "dL/dw", maxChars: 5000 });
console.log(`slides=${presentation.slides.items.length}`);
console.log(`oldFormulaTextboxes=${textboxes.ndjson.trim() ? textboxes.ndjson.trim().split(/\r?\n/).length : 0}`);
console.log(`verified=${OUT}\\slide-15.png`);
