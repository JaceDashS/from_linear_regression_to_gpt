import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";

async function main() {
  const p = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const snap = await p.inspect({ kind: "textbox", search: "Do not train on the exam", maxChars: 20000 });
  let removed = 0;
  for (const row of snap.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    if (item.slide === 8 && item.text && item.text.includes("Do not train on the exam")) {
      const shape = p.resolve(item.id);
      shape.delete();
      removed++;
    }
  }
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log("removed", removed);
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
