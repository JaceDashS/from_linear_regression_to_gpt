import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const CURVE = "C:\\workspace\\Attention is all you need\\ppt_build\\assets\\slide6_nonlinear_curve.png";

async function main() {
  const p = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const slide = p.slides.getItem(5);
  const snapshot = await p.inspect({ kind: "shape", maxChars: 360000 });
  let removed = 0;
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    if (item.slide !== 6) continue;
    const b = item.bbox || [];
    if (b.length !== 4) continue;
    const inChart = b[0] >= 190 && b[0] <= 980 && b[1] >= 245 && b[1] <= 575;
    if (!inChart) continue;
    const shape = p.resolve(item.id);
    const proto = shape.toProto?.();
    const lineColor = proto?.shape?.line?.fill?.color?.value;
    const fillColor = proto?.shape?.fill?.color?.value;
    const isLine = proto?.shape?.geometry === 1;
    if (isLine && (lineColor === "10B981" || fillColor === "10B981")) {
      shape.delete();
      removed++;
    }
  }

  const bytes = await fs.readFile(CURVE);
  slide.images.add({
    blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    contentType: "image/png",
    alt: "Smooth nonlinear regression curve",
    fit: "contain",
    position: { left: 200, top: 252, width: 760, height: 318 },
  });

  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log("removed", removed, "green line fragments and added smooth curve image");
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
