import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const C = { blue: "#2563EB", green: "#10B981", orange: "#F97316", red: "#EF4444" };
const COLOR_VALUES = new Set(["2563EB", "10B981", "F97316", "EF4444"]);

function addLine(slide, x1, y1, x2, y2, color, width = 3) {
  return slide.shapes.add({
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
}

function addPolyline(slide, points, color, width = 3) {
  for (let i = 0; i < points.length - 1; i++) {
    addLine(slide, points[i][0], points[i][1], points[i + 1][0], points[i + 1][1], color, width);
  }
}

function regressionMapper() {
  const chart = { left: 210, top: 305, width: 780, height: 260 };
  const xMin = 0.5, xMax = 6.5, yMin = 0, yMax = 6;
  const sx = x => chart.left + ((x - xMin) / (xMax - xMin)) * chart.width;
  const sy = y => chart.top + ((yMax - y) / (yMax - yMin)) * chart.height;
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

async function removeSmallGuideDots(presentation) {
  const snapshot = await presentation.inspect({ kind: "shape", maxChars: 520000 });
  let removed = 0;
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    if (item.slide < 3 || item.slide > 12) continue;
    const b = item.bbox || [];
    if (b.length !== 4) continue;
    if (b[2] > 10 || b[3] > 10) continue;
    const shape = presentation.resolve(item.id);
    const color = shape.toProto?.()?.shape?.fill?.color?.value;
    if (COLOR_VALUES.has(color)) {
      shape.delete();
      removed++;
    }
  }
  return removed;
}

async function main() {
  const p = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const removed = await removeSmallGuideDots(p);

  // Slides 3-4: model lines from linear regression sequence.
  {
    const { sx, sy } = regressionMapper();
    addLine(p.slides.getItem(2), sx(0.8), sy(0.55 * 0.8 + 1.25), sx(6.2), sy(0.55 * 6.2 + 1.25), C.orange, 3.2);
    addLine(p.slides.getItem(3), sx(0.8), sy(0.80 * 0.8 + 0.35), sx(6.2), sy(0.80 * 6.2 + 0.35), C.orange, 3.2);
  }

  // Slides 5-7: attendance regression, nonlinear trend, and overfit curve.
  {
    const { sx, sy } = attendanceMapper();
    const linear = x => 0.67 * x + 25;
    const nonlinear = x => 48 + 47 / (1 + Math.exp(-(x - 72) / 7));
    addLine(p.slides.getItem(4), sx(53), sy(linear(53)), sx(97), sy(linear(97)), C.orange, 3.2);
    addLine(p.slides.getItem(5), sx(52), sy(linear(52)), sx(99), sy(linear(99)), C.orange, 2.5);
    addPolyline(p.slides.getItem(5), sampledPoints(sx, sy, nonlinear, 52, 99, 36), C.green, 3.2);

    const data = [
      [52, 61], [55, 48], [57, 67], [61, 58], [63, 73], [66, 54],
      [69, 76], [72, 70], [76, 82], [78, 65], [81, 88], [84, 77],
      [87, 94], [90, 83], [93, 90], [96, 72], [98, 97],
    ].sort((a, b) => a[0] - b[0]);
    const overfitPts = [];
    for (let i = 0; i < data.length - 1; i++) {
      const [x1, y1] = data[i], [x2, y2] = data[i + 1];
      for (let j = 0; j <= 5; j++) {
        const t = j / 5;
        const x = x1 + (x2 - x1) * t;
        const wiggle = Math.sin(t * Math.PI) * (i % 2 === 0 ? 7 : -7);
        const y = y1 * (1 - t) + y2 * t + wiggle;
        overfitPts.push([sx(x), sy(y)]);
      }
    }
    addPolyline(p.slides.getItem(6), overfitPts, C.red, 2.5);
    addPolyline(p.slides.getItem(6), sampledPoints(sx, sy, nonlinear, 52, 99, 36), C.green, 3.2);
  }

  // Slide 9: train / validation error curves.
  {
    const slide = p.slides.getItem(8);
    const train = [[150, 394], [300, 430], [455, 470], [610, 510], [780, 542], [950, 564], [1115, 578]];
    const valid = [[150, 372], [300, 400], [455, 436], [610, 458], [780, 438], [950, 386], [1115, 330]];
    addPolyline(slide, train, C.blue, 3);
    addPolyline(slide, valid, C.orange, 3);
  }

  // Slides 11-12: MSE loss curve.
  {
    const { f, sx, sy } = lossMapper();
    const pts = sampledPoints(sx, sy, f, 0, 10, 44);
    addPolyline(p.slides.getItem(10), pts, C.blue, 3);
    addPolyline(p.slides.getItem(11), pts, C.blue, 3);
  }

  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log("removed guide dots", removed);
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
