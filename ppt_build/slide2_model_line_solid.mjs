import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const ORANGE = "F97316";

function addSolidLine(slide) {
  slide.shapes.add({
    geometry: "line",
    position: { left: 249, top: 348, width: 705, height: 56 },
    fill: "none",
    line: { style: "solid", fill: `#${ORANGE}`, width: 4 },
  });
}

async function main() {
  const p = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const slide = p.slides.getItem(1);
  const snapshot = await p.inspect({ kind: "shape", maxChars: 160000 });
  let removed = 0;
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    if (item.slide !== 2) continue;
    const b = item.bbox || [];
    if (b.length !== 4) continue;
    const smallDot = b[2] <= 10 && b[3] <= 10;
    if (!smallDot) continue;
    const shape = p.resolve(item.id);
    const color = shape.toProto?.()?.shape?.fill?.color?.value;
    if (color === ORANGE) {
      shape.delete();
      removed++;
    }
  }
  addSolidLine(slide);
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log("removed orange dots", removed, "and added solid line");
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
