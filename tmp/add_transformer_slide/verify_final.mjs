import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed_cnn_added_rnn_gru_added_seq2seq_added_transformer_added.pptx";
const OUT = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_transformer_slide\\verified\\final-render";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

await fs.mkdir(`${OUT}\\layouts`, { recursive: true });
const presentation = await PresentationFile.importPptx(await FileBlob.load(FINAL));
for (let index = 0; index < presentation.slides.items.length; index += 1) {
  const slide = presentation.slides.getItem(index);
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  const scale = [12, 18, 26, 27, 36].includes(index) ? 2 : 1;
  await writeBlob(`${OUT}\\${stem}.png`, await slide.export({ format: "png", scale }));
  await fs.writeFile(`${OUT}\\layouts\\${stem}.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");
}
console.log(`slides=${presentation.slides.items.length}`);
console.log(`rendered=${presentation.slides.items.length}`);
