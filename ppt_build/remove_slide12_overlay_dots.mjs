import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";

async function main() {
  const p = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const snapshot = await p.inspect({ kind: "shape", maxChars: 360000 });
  let removed = 0;
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    if (item.slide !== 12) continue;
    const b = item.bbox || [];
    if (b.length !== 4) continue;
    const smallDot = b[2] <= 8 && b[3] <= 8;
    if (!smallDot) continue;
    const shape = p.resolve(item.id);
    const color = shape.toProto?.()?.shape?.fill?.color?.value;
    if (color === "EF4444" || color === "F97316" || color === "10B981") {
      shape.delete();
      removed++;
    }
  }
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log("removed", removed, "overlay dots");
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
