import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const presentation = await PresentationFile.importPptx(
  await FileBlob.load("C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx"),
);
for (const query of ["slides reorder move", "slide collection remove insert", "reorder slide"]) {
  const result = presentation.help(query, { include: ["index", "examples", "notes"], maxChars: 8000 });
  console.log(`QUERY ${query}\n${result.ndjson ?? result}`);
}
