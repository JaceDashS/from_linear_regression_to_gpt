import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const C = { blue: "#2563EB", green: "#10B981", orange: "#F97316", red: "#EF4444" };
const GUIDE_COLORS = new Set(["2563EB", "10B981", "F97316", "EF4444"]);

function addNamedLine(slide, x1, y1, x2, y2, color, width = 3) {
  const flip = y2 < y1 ? "flipV" : "normal";
  const s = slide.shapes.add({
    geometry: "line",
    position: {
      left: Math.min(x1, x2),
      top: Math.min(y1, y2),
      width: Math.abs(x2 - x1) || 1,
      height: Math.abs(y2 - y1) || 1,
    },
    fill: "none",
    line: { style: "solid", fill: color, width },
  });
  s.data.name = `solidguide_${flip}`;
  return s;
}

function addPolyline(slide, points, color, width = 3) {
  for (let i = 0; i < points.length - 1; i++) {
    addNamedLine(slide, points[i][0], points[i][1], points[i + 1][0], points[i + 1][1], color, width);
  }
}

function regressionMapper() {
  const chart = { left: 210, top: 305, width: 780, height: 260 };
  const sx = x => chart.left + ((x - 0.5) / 6) * chart.width;
  const sy = y => chart.top + ((6 - y) / 6) * chart.height;
  return { sx, sy };
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

async function removeWrongGuideLines(presentation) {
  const snapshot = await presentation.inspect({ kind: "shape", maxChars: 700000 });
  let removed = 0;
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    const slide = item.slide;
    if (![3, 4, 5, 6, 7, 9, 11, 12].includes(slide)) continue;
    const b = item.bbox || [];
    if (b.length !== 4) continue;
    const shape = presentation.resolve(item.id);
    const proto = shape.toProto?.();
    const geometry = proto?.shape?.geometry;
    const color = proto?.shape?.line?.fill?.color?.value;
    if (geometry !== 1 || !GUIDE_COLORS.has(color)) continue;

    let shouldDelete = false;
    if ([3, 4, 5, 6].includes(slide) && ["F97316", "10B981"].includes(color) && (b[2] > 30 || b[3] > 30)) shouldDelete = true;
    if (slide === 7 && ["EF4444", "10B981"].includes(color) && (b[2] > 8 || b[3] > 8)) shouldDelete = true;
    if (slide === 9 && ["2563EB", "F97316", "EF4444", "10B981"].includes(color) && (b[2] > 8 || b[3] > 8)) shouldDelete = true;
    if ([11, 12].includes(slide) && color === "2563EB" && (b[2] > 8 || b[3] > 8)) shouldDelete = true;
    if (shouldDelete) {
      shape.delete();
      removed++;
    }
  }
  return removed;
}

async function main() {
  const p = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const removed = await removeWrongGuideLines(p);

  const { sx: rsx, sy: rsy } = regressionMapper();
  addNamedLine(p.slides.getItem(2), rsx(0.8), rsy(0.55 * 0.8 + 1.25), rsx(6.2), rsy(0.55 * 6.2 + 1.25), C.orange, 3.2);
  addNamedLine(p.slides.getItem(3), rsx(0.8), rsy(0.80 * 0.8 + 0.35), rsx(6.2), rsy(0.80 * 6.2 + 0.35), C.orange, 3.2);

  const { sx, sy } = attendanceMapper();
  const linear = x => 0.67 * x + 25;
  const nonlinear = x => 48 + 47 / (1 + Math.exp(-(x - 72) / 7));
  addNamedLine(p.slides.getItem(4), sx(53), sy(linear(53)), sx(97), sy(linear(97)), C.orange, 3.2);
  addNamedLine(p.slides.getItem(5), sx(52), sy(linear(52)), sx(99), sy(linear(99)), C.orange, 2.5);
  addPolyline(p.slides.getItem(5), sampledPoints(sx, sy, nonlinear, 52, 99, 32), C.green, 3.2);

  const data = [
    [52, 61], [55, 48], [57, 67], [61, 58], [63, 73], [66, 54],
    [69, 76], [72, 70], [76, 82], [78, 65], [81, 88], [84, 77],
    [87, 94], [90, 83], [93, 90], [96, 72], [98, 97],
  ].sort((a, b) => a[0] - b[0]);
  const overfitPts = [];
  for (let i = 0; i < data.length - 1; i++) {
    const [x1, y1] = data[i], [x2, y2] = data[i + 1];
    for (let j = 0; j <= 4; j++) {
      const t = j / 4;
      const x = x1 + (x2 - x1) * t;
      const wiggle = Math.sin(t * Math.PI) * (i % 2 === 0 ? 7 : -7);
      const y = y1 * (1 - t) + y2 * t + wiggle;
      overfitPts.push([sx(x), sy(y)]);
    }
  }
  addPolyline(p.slides.getItem(6), overfitPts, C.red, 2.3);
  addPolyline(p.slides.getItem(6), sampledPoints(sx, sy, nonlinear, 52, 99, 32), C.green, 3.2);

  const slide9 = p.slides.getItem(8);
  addPolyline(slide9, [[150, 394], [300, 430], [455, 470], [610, 510], [780, 542], [950, 564], [1115, 578]], C.blue, 3);
  addPolyline(slide9, [[150, 372], [300, 400], [455, 436], [610, 458], [780, 438], [950, 386], [1115, 330]], C.orange, 3);
  for (const [x, color] of [[195, C.red], [565, C.green], [912, C.orange]]) {
    addNamedLine(slide9, x, 336, x, 593, color, 1.5);
  }

  const { f, sx: lsx, sy: lsy } = lossMapper();
  const lossPts = sampledPoints(lsx, lsy, f, 0, 10, 40);
  addPolyline(p.slides.getItem(10), lossPts, C.blue, 3);
  addPolyline(p.slides.getItem(11), lossPts, C.blue, 3);

  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log("removed wrong guide lines", removed);
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
