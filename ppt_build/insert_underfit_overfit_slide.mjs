import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const C = { ink:"#0B0F19", muted:"#667085", line:"#CBD5E1", panel:"#F8FAFC", blue:"#2563EB", green:"#10B981", orange:"#F97316", red:"#EF4444", white:"#FFFFFF" };

function addText(slide, text, pos, style = {}) {
  const s = slide.shapes.add({ geometry:"textbox", position:pos, fill:"none", line:{style:"solid", fill:"none", width:0} });
  s.text = text;
  s.text.style = { fontSize:style.fontSize ?? 24, bold:style.bold ?? false, color:style.color ?? C.ink, alignment:style.alignment ?? "left" };
  return s;
}
function box(slide, text, pos, opts = {}) {
  const s = slide.shapes.add({ geometry:opts.geometry ?? "roundRect", position:pos, fill:opts.fill ?? C.white, line:{style:"solid", fill:opts.line ?? C.line, width:opts.lineWidth ?? 1.5}, borderRadius:"rounded-xl" });
  if (text) { s.text = text; s.text.style = { fontSize:opts.fontSize ?? 22, bold:opts.bold ?? true, color:opts.color ?? C.ink, alignment:opts.alignment ?? "center" }; }
  return s;
}
function circle(slide, x, y, r, color) {
  return slide.shapes.add({ geometry:"ellipse", position:{left:x-r, top:y-r, width:r*2, height:r*2}, fill:color, line:{style:"solid", fill:color, width:1} });
}
function line(slide, x1, y1, x2, y2, opts = {}) {
  return slide.shapes.add({ geometry:"line", position:{left:Math.min(x1,x2), top:Math.min(y1,y2), width:Math.abs(x2-x1)||1, height:Math.abs(y2-y1)||1}, fill:"none", line:{style:opts.style ?? "solid", fill:opts.color ?? C.line, width:opts.width ?? 2} });
}
function title(slide) {
  addText(slide, "Regression", {left:64, top:34, width:420, height:30}, {fontSize:16, bold:true, color:C.blue});
  addText(slide, "Underfitting and overfitting are two failure modes", {left:64, top:62, width:1080, height:62}, {fontSize:40, bold:true});
}
function frame(slide) { box(slide, "", {left:64, top:142, width:1152, height:500}, {fill:C.panel, line:"#E2E8F0"}); }
function footer(slide, n) { addText(slide, String(n).padStart(2,"0"), {left:1166, top:660, width:50, height:24}, {fontSize:13, color:C.muted, alignment:"right"}); }
function notes(slide) {
  slide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard machine learning concepts.");
  slide.speakerNotes.setVisible(true);
}

const panels = [
  { title:"underfit", color:C.red, x:105, line:"too simple", desc:"misses the pattern", kind:"under" },
  { title:"good fit", color:C.green, x:460, line:"right complexity", desc:"captures the trend", kind:"good" },
  { title:"overfit", color:C.orange, x:815, line:"too complex", desc:"memorizes noise", kind:"over" },
];

function drawMini(slide, cfg) {
  const left = cfg.x, top = 230, w = 300, h = 245;
  box(slide, "", {left, top, width:w, height:h}, {fill:C.white, line:"#E2E8F0"});
  addText(slide, cfg.title, {left, top:180, width:w, height:34}, {fontSize:28, bold:true, color:cfg.color, alignment:"center"});
  const sx = x => left + 42 + x * 34;
  const sy = y => top + 195 - y * 28;
  const pts = [[0.4,1.1],[1.1,1.8],[1.7,1.4],[2.2,2.7],[2.9,2.4],[3.4,3.6],[4.1,3.1],[4.7,4.2],[5.4,4.0],[6.0,5.0]];
  line(slide, left+34, top+205, left+w-30, top+205, {color:"#94A3B8", width:1.4});
  line(slide, left+34, top+205, left+34, top+35, {color:"#94A3B8", width:1.4});
  for (const [x,y] of pts) circle(slide, sx(x), sy(y), 5.8, C.blue);
  if (cfg.kind === "under") {
    for (let i=0;i<=34;i++) { const x=i/34*6.2; circle(slide, sx(x), sy(2.5), 2.5, cfg.color); }
  } else if (cfg.kind === "good") {
    for (let i=0;i<=42;i++) { const x=i/42*6.2; circle(slide, sx(x), sy(0.95+0.67*x), 2.5, cfg.color); }
  } else {
    for (let i=0;i<pts.length-1;i++) {
      const [x1,y1]=pts[i], [x2,y2]=pts[i+1];
      for (let j=0;j<=7;j++) {
        const t=j/7; const x=x1+(x2-x1)*t; const y=y1*(1-t)+y2*t+Math.sin(t*Math.PI)*(i%2?-.55:.55);
        circle(slide, sx(x), sy(y), 2.5, cfg.color);
      }
    }
  }
  addText(slide, cfg.line, {left:left+25, top:500, width:w-50, height:28}, {fontSize:22, bold:true, color:cfg.color, alignment:"center"});
  addText(slide, cfg.desc, {left:left+25, top:532, width:w-50, height:28}, {fontSize:18, color:C.muted, alignment:"center"});
}

async function renumberFooters(presentation) {
  const snapshot = await presentation.inspect({ kind:"textbox", maxChars:220000 });
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    const b = item.bbox || [];
    if (b.length === 4 && Math.abs(b[0]-1166)<4 && Math.abs(b[1]-660)<8 && item.slide) {
      const shape = presentation.resolve(item.id);
      shape.text = String(item.slide).padStart(2,"0");
      shape.text.style = { fontSize:13, color:C.muted, alignment:"right" };
    }
  }
}

async function updateRoadmapXml() {
  // Roadmap update is handled by direct XML replacement after export in a separate Python-free flow would be awkward here.
}

async function main() {
  const p = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const after = p.slides.getItem(6);
  const slide = p.slides.insert({ after }).slide;
  slide.background.fill = C.white;
  title(slide); frame(slide);
  panels.forEach(panel => drawMini(slide, panel));
  footer(slide, 8); notes(slide);
  await renumberFooters(p);
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log(PPTX);
}

main().catch(e => { console.error(e); process.exitCode = 1; });
