import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const SVG = "C:\\workspace\\Attention is all you need\\ppt_build\\assets\\slide43_top_p_dotted_guide.svg";

async function main() {
  const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720">
  <rect width="1280" height="720" fill="none"/>
  <path d="M 722 300 L 900 300" fill="none" stroke="#10B981" stroke-width="2" stroke-linecap="round" stroke-dasharray="8 9"/>
</svg>`;
  await fs.writeFile(SVG, svg, "utf8");

  const presentation = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const snap = await presentation.inspect({ kind: "shape,image", maxChars: 500000 });
  let removed = 0;
  for (const row of snap.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    if (item.slide !== 43) continue;
    const b = item.bbox || [];
    if (b.length !== 4) continue;
    const resolved = presentation.resolve(item.id);
    const proto = resolved.toProto?.();
    const isTopPDashedLine =
      proto?.shape?.geometry === 1 &&
      proto?.shape?.line?.style === 2 &&
      proto?.shape?.line?.fill?.color?.value === "10B981" &&
      Math.abs(b[0] - 722) < 2 &&
      Math.abs(b[1] - 300) < 2 &&
      Math.abs(b[2] - 178) < 2;
    const isOldOverlay =
      item.kind === "image" &&
      b[0] === 0 &&
      b[1] === 0 &&
      b[2] === 1280 &&
      b[3] === 720;
    if (isTopPDashedLine || isOldOverlay) {
      resolved.delete();
      removed++;
    }
  }

  const bytes = await fs.readFile(SVG);
  presentation.slides.getItem(42).images.add({
    blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    contentType: "image/svg+xml",
    alt: "Top-p dotted threshold guide as SVG",
    fit: "contain",
    position: { left: 0, top: 0, width: 1280, height: 720 },
  });

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(PPTX);
  console.log(`removed ${removed} remaining dotted guide objects and added SVG`);
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
