import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";

const lineColors = {
  "sh/je9wfyxs": "10B981",
  "sh/c7ilwbax": "F97316",
  "sh/ytcnmh4b": "EF4444",
};

async function main() {
  const p = await PresentationFile.importPptx(await FileBlob.load(PPTX));

  for (const [id, color] of Object.entries(lineColors)) {
    const shape = p.resolve(id);
    shape.data.shape.line.fill.color.value = color;
  }

  const snapshot = await p.inspect({ kind: "shape", maxChars: 320000 });
  let removed = 0;
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    if (item.slide !== 11) continue;
    const b = item.bbox || [];
    if (b.length !== 4) continue;
    const shape = p.resolve(item.id);
    const proto = shape.toProto?.();
    const color = proto?.shape?.fill?.color?.value;
    const smallDot = b[2] <= 8 && b[3] <= 8;
    if (smallDot && (color === "10B981" || color === "F97316")) {
      shape.delete();
      removed++;
    }
  }

  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log("updated tangent colors, removed", removed, "small guide dots");
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
