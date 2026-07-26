import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const ASSET_DIR = "C:\\workspace\\Attention is all you need\\ppt_build\\assets";
const C = { blue: "#2563EB", green: "#10B981", orange: "#F97316", red: "#EF4444" };

function svgWrap(body) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720">
  <rect width="1280" height="720" fill="none"/>
${body}
</svg>`;
}

function catmullRomPath(points) {
  if (points.length < 2) return "";
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

function strokePath(points, color, width) {
  return `  <path d="${catmullRomPath(points)}" fill="none" stroke="${color}" stroke-width="${width}" stroke-linecap="round" stroke-linejoin="round"/>`;
}

function attendanceMapper() {
  const chart = { left: 200, top: 252, width: 760, height: 318 };
  const sx = x => chart.left + ((x - 50) / 50) * chart.width;
  const sy = y => chart.top + chart.height - ((y - 45) / 55) * chart.height;
  return { sx, sy };
}

function lossMapper() {
  const chart = { left: 185, top: 265, width: 760, height: 300 };
  const f = x => 0.055 * (x - 5.1) * (x - 5.1) + 1.0;
  const sx = x => chart.left + (x / 10) * chart.width;
  const sy = y => chart.top + chart.height - ((y - 0.7) / 2.25) * chart.height;
  return { f, sx, sy };
}

function sampledPoints(sx, sy, fn, start, end, n) {
  const pts = [];
  for (let i = 0; i <= n; i++) {
    const x = start + ((end - start) * i) / n;
    pts.push([sx(x), sy(fn(x))]);
  }
  return pts;
}

async function writeSvg(name, paths) {
  const path = `${ASSET_DIR}\\${name}`;
  await fs.writeFile(path, svgWrap(paths.join("\n")), "utf8");
  return path;
}

function addSvg(slide, svgPath, alt) {
  return fs.readFile(svgPath).then(bytes => {
    slide.images.add({
      blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
      contentType: "image/svg+xml",
      alt,
      fit: "contain",
      position: { left: 0, top: 0, width: 1280, height: 720 },
    });
  });
}

async function removeCurveLines(presentation) {
  const snapshot = await presentation.inspect({ kind: "shape,image", maxChars: 1000000 });
  let removed = 0;
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    const b = item.bbox || [];
    if (b.length !== 4) continue;

    const shape = presentation.resolve(item.id);
    const proto = shape.toProto?.();
    const geometry = proto?.shape?.geometry;
    const color = proto?.shape?.line?.fill?.color?.value;
    if (geometry !== 1 || !color) continue;

    let shouldDelete = false;
    const [x, y, w, h] = b;
    const insideAttendance = x >= 180 && x <= 1000 && y >= 240 && y <= 590;
    const insideError = x >= 120 && x <= 1140 && y >= 310 && y <= 600;
    const insideLoss = x >= 170 && x <= 970 && y >= 250 && y <= 575;

    if (item.slide === 7 && insideAttendance && ["EF4444", "10B981"].includes(color) && (w > 3 || h > 3)) shouldDelete = true;
    if (item.slide === 9 && insideError && ["2563EB", "F97316"].includes(color) && (w > 3 || h > 3)) shouldDelete = true;
    if ([11, 12].includes(item.slide) && insideLoss && color === "2563EB" && (w > 3 || h > 3)) shouldDelete = true;

    if (shouldDelete) {
      shape.delete();
      removed++;
    }
  }
  return removed;
}

async function main() {
  await fs.mkdir(ASSET_DIR, { recursive: true });

  const { sx, sy } = attendanceMapper();
  const nonlinear = x => 48 + 47 / (1 + Math.exp(-(x - 72) / 7));
  const trend = sampledPoints(sx, sy, nonlinear, 52, 99, 22);
  const data = [
    [52, 61], [55, 48], [57, 67], [61, 58], [63, 73], [66, 54],
    [69, 76], [72, 70], [76, 82], [78, 65], [81, 88], [84, 77],
    [87, 94], [90, 83], [93, 90], [96, 72], [98, 97],
  ];
  const overfit = [];
  for (let i = 0; i < data.length - 1; i++) {
    const [x1, y1] = data[i], [x2, y2] = data[i + 1];
    for (let j = 0; j <= 2; j++) {
      const t = j / 2;
      const x = x1 + (x2 - x1) * t;
      const wiggle = Math.sin(t * Math.PI) * (i % 2 === 0 ? 8 : -8);
      const y = y1 * (1 - t) + y2 * t + wiggle;
      overfit.push([sx(x), sy(y)]);
    }
  }

  const errorBlue = [[150, 394], [300, 430], [455, 470], [610, 510], [780, 542], [950, 564], [1115, 578]];
  const errorOrange = [[150, 372], [300, 400], [455, 436], [610, 458], [780, 438], [950, 386], [1115, 330]];

  const { f, sx: lsx, sy: lsy } = lossMapper();
  const loss = sampledPoints(lsx, lsy, f, 0, 10, 24);

  const slide7Svg = await writeSvg("slide7_curves.svg", [strokePath(overfit, C.red, 2.4), strokePath(trend, C.green, 3.4)]);
  const slide9Svg = await writeSvg("slide9_error_curves.svg", [strokePath(errorBlue, C.blue, 3.2), strokePath(errorOrange, C.orange, 3.2)]);
  const slideLossSvg = await writeSvg("slide11_12_loss_curve.svg", [strokePath(loss, C.blue, 3.2)]);

  const presentation = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const removed = await removeCurveLines(presentation);

  await addSvg(presentation.slides.getItem(6), slide7Svg, "Overfit and trend curves as SVG");
  await addSvg(presentation.slides.getItem(8), slide9Svg, "Training and validation error curves as SVG");
  await addSvg(presentation.slides.getItem(10), slideLossSvg, "MSE loss curve as SVG");
  await addSvg(presentation.slides.getItem(11), slideLossSvg, "MSE loss curve as SVG");

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(PPTX);
  console.log(`removed ${removed} segmented curve lines; added SVG curves on slides 7, 9, 11, and 12`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
