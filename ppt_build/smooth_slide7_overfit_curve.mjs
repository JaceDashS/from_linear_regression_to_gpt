import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const SVG = "C:\\workspace\\Attention is all you need\\ppt_build\\assets\\slide7_curves_smooth.svg";
const C = { green: "#10B981", red: "#EF4444" };

function svgWrap(body) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720">
  <rect width="1280" height="720" fill="none"/>
${body}
</svg>`;
}

function catmullRomPath(points) {
  let d = `M ${points[0][0].toFixed(2)} ${points[0][1].toFixed(2)}`;
  for (let i = 0; i < points.length - 1; i++) {
    const p0 = points[Math.max(0, i - 1)];
    const p1 = points[i];
    const p2 = points[i + 1];
    const p3 = points[Math.min(points.length - 1, i + 2)];
    const c1 = [p1[0] + (p2[0] - p0[0]) / 7.5, p1[1] + (p2[1] - p0[1]) / 7.5];
    const c2 = [p2[0] - (p3[0] - p1[0]) / 7.5, p2[1] - (p3[1] - p1[1]) / 7.5];
    d += ` C ${c1[0].toFixed(2)} ${c1[1].toFixed(2)}, ${c2[0].toFixed(2)} ${c2[1].toFixed(2)}, ${p2[0].toFixed(2)} ${p2[1].toFixed(2)}`;
  }
  return d;
}

function stroke(points, color, width) {
  return `  <path d="${catmullRomPath(points)}" fill="none" stroke="${color}" stroke-width="${width}" stroke-linecap="round" stroke-linejoin="round"/>`;
}

function attendanceMapper() {
  const chart = { left: 200, top: 252, width: 760, height: 318 };
  const sx = x => chart.left + ((x - 50) / 50) * chart.width;
  const sy = y => chart.top + chart.height - ((y - 45) / 55) * chart.height;
  return { sx, sy };
}

function sampledPoints(sx, sy, fn, start, end, n) {
  const pts = [];
  for (let i = 0; i <= n; i++) {
    const x = start + ((end - start) * i) / n;
    pts.push([sx(x), sy(fn(x))]);
  }
  return pts;
}

async function main() {
  const { sx, sy } = attendanceMapper();
  const trendFn = x => 48 + 47 / (1 + Math.exp(-(x - 72) / 7));
  const trend = sampledPoints(sx, sy, trendFn, 52, 99, 22);

  const data = [
    [52, 61], [55, 48], [57, 67], [61, 58], [63, 73], [66, 54],
    [69, 76], [72, 70], [76, 82], [78, 65], [81, 88], [84, 77],
    [87, 94], [90, 83], [93, 90], [96, 72], [98, 97],
  ];
  const overfit = [];
  for (let i = 0; i < data.length - 1; i++) {
    const [x1, y1] = data[i];
    const [x2, y2] = data[i + 1];
    const midX = (x1 + x2) / 2;
    const midY = (y1 + y2) / 2 + (i % 2 === 0 ? 3.8 : -3.8);
    if (i === 0) overfit.push([sx(x1), sy(y1)]);
    overfit.push([sx(midX), sy(midY)]);
    overfit.push([sx(x2), sy(y2)]);
  }

  const body = [stroke(overfit, C.red, 2.3), stroke(trend, C.green, 3.4)];
  const svgText = svgWrap(body.join("\n"));
  await fs.writeFile(SVG, svgText, "utf8");

  const presentation = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const snap = await presentation.inspect({ kind: "image", maxChars: 200000 });
  let removed = 0;
  for (const row of snap.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    const b = item.bbox || [];
    if (item.slide === 7 && b.length === 4 && b[0] === 0 && b[1] === 0 && b[2] === 1280 && b[3] === 720) {
      presentation.resolve(item.id).delete();
      removed++;
    }
  }

  const slide = presentation.slides.getItem(6);
  const bytes = await fs.readFile(SVG);
  slide.images.add({
    blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    contentType: "image/svg+xml",
    alt: "Smooth overfit and trend curves as SVG",
    fit: "contain",
    position: { left: 0, top: 0, width: 1280, height: 720 },
  });

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(PPTX);
  console.log(`replaced ${removed} slide 7 curve overlay with smoother SVG`);
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
