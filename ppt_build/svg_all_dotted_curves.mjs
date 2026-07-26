import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const SVG = "C:\\workspace\\Attention is all you need\\ppt_build\\assets\\slide13_loss_path_dotted.svg";

const dottedPath = [
  [445, 290],
  [482, 304],
  [497, 338],
  [479, 377],
  [447, 414],
  [366, 418],
  [352, 454],
  [365, 493],
  [426, 531],
];

function catmullRomPath(points) {
  let d = `M ${points[0][0].toFixed(2)} ${points[0][1].toFixed(2)}`;
  for (let i = 0; i < points.length - 1; i++) {
    const p0 = points[Math.max(0, i - 1)];
    const p1 = points[i];
    const p2 = points[i + 1];
    const p3 = points[Math.min(points.length - 1, i + 2)];
    const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6];
    const c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6];
    d += ` C ${c1[0].toFixed(2)} ${c1[1].toFixed(2)}, ${c2[0].toFixed(2)} ${c2[1].toFixed(2)}, ${p2[0].toFixed(2)} ${p2[1].toFixed(2)}`;
  }
  return d;
}

async function main() {
  const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720">
  <rect width="1280" height="720" fill="none"/>
  <path d="${catmullRomPath(dottedPath)}" fill="none" stroke="#93D64F" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round" stroke-dasharray="9 8"/>
</svg>`;
  await fs.writeFile(SVG, svg, "utf8");

  const presentation = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const snap = await presentation.inspect({ kind: "shape,image", maxChars: 800000 });
  let removed = 0;
  for (const row of snap.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    const b = item.bbox || [];
    if (b.length !== 4) continue;
    const resolved = presentation.resolve(item.id);
    const proto = resolved.toProto?.();

    const redOnePxCurveRemnant =
      item.slide === 7 &&
      proto?.shape?.geometry === 1 &&
      proto?.shape?.line?.fill?.color?.value === "EF4444" &&
      b[0] >= 260 &&
      b[0] <= 920 &&
      b[1] >= 280 &&
      b[1] <= 560 &&
      b[2] <= 1.2 &&
      b[3] <= 1.2;

    const oldDottedOverlay =
      item.slide === 13 &&
      item.kind === "image" &&
      item.contentType === "image/svg+xml" &&
      b[0] === 0 &&
      b[1] === 0 &&
      b[2] === 1280 &&
      b[3] === 720 &&
      item.id !== undefined;

    if (redOnePxCurveRemnant) {
      resolved.delete();
      removed++;
    }
  }

  const bytes = await fs.readFile(SVG);
  presentation.slides.getItem(12).images.add({
    blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    contentType: "image/svg+xml",
    alt: "Dotted gradient descent path as SVG",
    fit: "contain",
    position: { left: 0, top: 0, width: 1280, height: 720 },
  });

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(PPTX);
  console.log(`removed ${removed} tiny curve remnants and added slide 13 dotted path SVG`);
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
