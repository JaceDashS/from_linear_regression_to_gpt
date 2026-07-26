import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const SVG = "C:\\workspace\\Attention is all you need\\ppt_build\\assets\\slide13_annealing_curve.svg";

function svgWrap(body) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720">
  <rect width="1280" height="720" fill="none"/>
${body}
</svg>`;
}

function pointsForAnnealing() {
  const pts = [];
  const x0 = 912;
  const x1 = 1092;
  const y0 = 350;
  const y1 = 442;
  for (let i = 0; i <= 28; i++) {
    const t = i / 28;
    const x = x0 + (x1 - x0) * t;
    const y = y0 + (y1 - y0) * (1 - Math.exp(-3.2 * t)) / (1 - Math.exp(-3.2));
    pts.push([x, y]);
  }
  return pts;
}

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
  await fs.writeFile(
    SVG,
    svgWrap(`  <path d="${catmullRomPath(pointsForAnnealing())}" fill="none" stroke="#2563EB" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>`),
    "utf8",
  );

  const presentation = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const snap = await presentation.inspect({ kind: "shape,image", maxChars: 250000 });
  let removed = 0;
  for (const row of snap.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    if (item.slide !== 13) continue;
    const b = item.bbox || [];
    if (b.length !== 4) continue;
    const resolved = presentation.resolve(item.id);
    const proto = resolved.toProto?.();
    const color = proto?.shape?.fill?.color?.value;
    const line = proto?.shape?.line?.fill?.color?.value;
    const isExistingFullSlideSvg =
      item.kind === "image" &&
      item.contentType === "image/svg+xml" &&
      b[0] === 0 &&
      b[1] === 0 &&
      b[2] === 1280 &&
      b[3] === 720;
    const isSmallBlueDot =
      proto?.shape?.geometry === 35 &&
      color === "2563EB" &&
      line === "2563EB" &&
      b[0] >= 900 &&
      b[0] <= 1110 &&
      b[1] >= 335 &&
      b[1] <= 455 &&
      b[2] <= 8 &&
      b[3] <= 8;

    if (isSmallBlueDot || isExistingFullSlideSvg) {
      resolved.delete();
      removed++;
    }
  }

  const bytes = await fs.readFile(SVG);
  presentation.slides.getItem(12).images.add({
    blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    contentType: "image/svg+xml",
    alt: "Annealing decay curve as SVG",
    fit: "contain",
    position: { left: 0, top: 0, width: 1280, height: 720 },
  });

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(PPTX);
  console.log(`removed ${removed} blue dots and added SVG annealing curve`);
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
