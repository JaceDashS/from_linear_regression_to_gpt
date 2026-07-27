import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed.pptx";
const OUT = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\fix_slide18_nn\\verified\\final-render";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

await fs.mkdir(`${OUT}\\layouts`, { recursive: true });
const presentation = await PresentationFile.importPptx(await FileBlob.load(FINAL));
for (let index = 0; index < presentation.slides.items.length; index += 1) {
  const slide = presentation.slides.getItem(index);
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(`${OUT}\\${stem}.png`, await slide.export({ format: "png", scale: 1 }));
  await fs.writeFile(`${OUT}\\layouts\\${stem}.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");
}
console.log(`rendered=${presentation.slides.items.length}`);
