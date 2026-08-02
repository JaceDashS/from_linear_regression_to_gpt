import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const deckPath = "C:/workspace/from_linear_regression_to_gpt/from_linear_regression_to_gpt - Future of LLM.pptx";
const outputDir = "C:/workspace/from_linear_regression_to_gpt/tmp/future_llm_edit";

await fs.mkdir(outputDir, { recursive: true });

const presentation = await PresentationFile.importPptx(await FileBlob.load(deckPath));
const snapshot = await presentation.inspect({
  kind: "slide,textbox,shape,layout",
  search: "sglang",
  include: "id,slide,name,title,text,textPreview,bbox,bboxUnit,isPlaceholder",
  maxChars: 12000,
});

console.log(snapshot.ndjson);

for (const slideNumber of [5, 6]) {
  const slide = presentation.slides.getItem(slideNumber - 1);
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(`${outputDir}/slide-${slideNumber}-before.layout.json`, await layout.text());
}
