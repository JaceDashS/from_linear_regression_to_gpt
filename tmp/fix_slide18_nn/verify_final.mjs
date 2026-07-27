import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed.pptx";
const OUT = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\fix_slide18_nn\\verified";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

await fs.mkdir(OUT, { recursive: true });
const presentation = await PresentationFile.importPptx(await FileBlob.load(FINAL));
const slide = presentation.slides.getItem(17);
await writeBlob(`${OUT}\\slide-18.png`, await slide.export({ format: "png", scale: 2 }));
const layoutText = await (await slide.export({ format: "layout" })).text();
await fs.writeFile(`${OUT}\\slide-18.layout.json`, layoutText, "utf8");
const layout = JSON.parse(layoutText);
const connectors = layout.elements.filter((element) => element.scope === "slide" && element.geometry === "connector");
const legacyLines = layout.elements.filter((element) => element.scope === "slide" && element.geometry === "line");
console.log(`slides=${presentation.slides.items.length}`);
console.log(`connectors=${connectors.length}`);
console.log(`legacyLines=${legacyLines.length}`);
console.log(`verified=${OUT}\\slide-18.png`);
