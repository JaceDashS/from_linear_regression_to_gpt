import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const source = "C:/docs/from_linear_regression_to_gpt/tmp/attention_dot_product_update/candidate.pptx";
const out = "C:/docs/from_linear_regression_to_gpt/tmp/attention_dot_product_update/verified";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
await fs.mkdir(out, { recursive: true });

for (const slideNumber of [32, 33]) {
  const slide = presentation.slides.getItem(slideNumber - 1);
  const png = await slide.export({ format: "png", scale: 1 });
  await fs.writeFile(`${out}/slide-${slideNumber}.png`, new Uint8Array(await png.arrayBuffer()));
  await fs.writeFile(`${out}/slide-${slideNumber}.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");
}

const snapshot = await presentation.inspect({
  kind: "slide,textbox,notes",
  search: "dot product|Softmax|Attention(Q",
  maxChars: 12000,
});
await fs.writeFile(`${out}/inspect.ndjson`, snapshot.ndjson, "utf8");
console.log(JSON.stringify({ slides: presentation.slides.items.length }));
