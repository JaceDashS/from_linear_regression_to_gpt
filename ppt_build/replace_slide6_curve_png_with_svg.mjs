import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const SVG = "C:\\workspace\\Attention is all you need\\ppt_build\\assets\\slide6_nonlinear_curve.svg";

async function main() {
  const p = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const slide = p.slides.getItem(5);
  const snap = await p.inspect({ kind: "image", maxChars: 80000 });
  let removed = 0;
  for (const row of snap.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    if (item.slide !== 6) continue;
    const b = item.bbox || [];
    if (b.length === 4 && Math.abs(b[0] - 200) < 2 && Math.abs(b[1] - 252) < 2 && Math.abs(b[2] - 760) < 2 && Math.abs(b[3] - 318) < 2) {
      p.resolve(item.id).delete();
      removed++;
    }
  }

  const bytes = await fs.readFile(SVG);
  slide.images.add({
    blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    contentType: "image/svg+xml",
    alt: "Smooth nonlinear regression curve as SVG",
    fit: "contain",
    position: { left: 200, top: 252, width: 760, height: 318 },
  });

  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log("removed png images", removed, "and added svg curve");
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
