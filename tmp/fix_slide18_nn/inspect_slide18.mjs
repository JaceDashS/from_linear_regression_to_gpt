import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation.pptx";
const TMP = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\fix_slide18_nn";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const slide = presentation.slides.getItem(17);
await writeBlob(`${TMP}\\before-slide-18.png`, await slide.export({ format: "png", scale: 2 }));
await fs.writeFile(`${TMP}\\before-slide-18.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");

const snapshot = await presentation.inspect({
  kind: "slide,shape,textbox",
  include: "id,slide,name,bbox,textPreview",
  maxChars: 200000,
});
const records = snapshot.ndjson.split(/\r?\n/).filter(Boolean).map(JSON.parse);
const slide18 = records.filter((record) => record.slide === 18);
await fs.writeFile(`${TMP}\\slide-18-inspect.ndjson`, slide18.map((record) => JSON.stringify(record)).join("\n") + "\n", "utf8");

const starter = await PresentationFile.exportPptx(presentation);
await starter.save(`${TMP}\\template-starter.pptx`);
console.log(`slide18Records=${slide18.length}`);
console.log(`slides=${presentation.slides.items.length}`);
