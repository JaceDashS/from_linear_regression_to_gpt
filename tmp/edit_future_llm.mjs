import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const sourcePath = "C:/workspace/from_linear_regression_to_gpt/from_linear_regression_to_gpt - Future of LLM.pptx";
const outputPath = "C:/workspace/from_linear_regression_to_gpt/tmp/future_llm_edit/edited.pptx";
const outputDir = "C:/workspace/from_linear_regression_to_gpt/tmp/future_llm_edit";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

await fs.mkdir(outputDir, { recursive: true });

const presentation = await PresentationFile.importPptx(await FileBlob.load(sourcePath));
const slide5 = presentation.resolve("sl/i107q5of");

const slideSnapshot = await presentation.inspect({
  kind: "slide,textbox,shape",
  search: "sglang",
  include: "id,slide,name,text,textPreview,bbox,bboxUnit",
  maxChars: 8000,
});

const records = slideSnapshot.ndjson
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line));

const slide5TitleRecord = records.find((record) => record.slide === 5 && record.text === "sglang");
const slide6TitleRecord = records.find((record) => record.slide === 6 && record.text === "sglang");

if (!slide5TitleRecord || !slide6TitleRecord) {
  throw new Error("Could not resolve the existing SGLang title shapes on slides 5 and 6.");
}

const slide5Title = presentation.resolve(slide5TitleRecord.id);
const slide6Title = presentation.resolve(slide6TitleRecord.id);
slide5Title.text = "SGLang optimizes the serving stack";
slide6Title.text = "SGLang cuts runtime by nearly half";

const summary = slide5.shapes.add({
  geometry: "textbox",
  name: "SGLang optimization summary",
  position: { left: 600, top: 225, width: 500, height: 180 },
  fill: "none",
  line: { style: "solid", fill: "none", width: 0 },
});
summary.text = "Prefix reuse\nContinuous batching\nSpeculative decoding";
summary.text.style = {
  fontSize: 27,
  typeface: "Times New Roman",
  color: "#0B0F19",
  alignment: "left",
  verticalAlignment: "middle",
};

slide5.speakerNotes.append(
  "\n[Sources]\n- https://github.com/sgl-project/sglang (feature overview; accessed 2026-08-03)\n[/Sources]",
);

for (const slideNumber of [5, 6]) {
  const slide = presentation.slides.getItem(slideNumber - 1);
  await writeBlob(
    `${outputDir}/slide-${slideNumber}-after.png`,
    await presentation.export({ slide, format: "png", scale: 1.5 }),
  );
  await fs.writeFile(
    `${outputDir}/slide-${slideNumber}-after.layout.json`,
    await (await slide.export({ format: "layout" })).text(),
  );
}

const verification = await presentation.inspect({
  kind: "slide,textbox,shape,notes",
  search: "SGLang",
  include: "id,slide,name,text,textPreview,bbox,bboxUnit",
  maxChars: 12000,
});
await fs.writeFile(`${outputDir}/verification.ndjson`, verification.ndjson);

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(outputPath);
console.log(outputPath);
