import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

const ROOT = path.resolve("../..");
const OUT_DIR = path.resolve(".");
const TMP_DIR = path.resolve("../../tmp/slides/porod-prezentace");
const PPTX_PATH = path.join(OUT_DIR, "porod-neni-jen-medicina.pptx");

const W = 1280;
const H = 720;
const EMU = 9525;

const C = {
  bg: "FFFFFF",
  bgSoft: "FFFBFB",
  bgWarm: "FFF5F0",
  rose: "DEC0B2",
  taupe: "B79C8B",
  accent: "654334",
  cta: "911C0B",
  text: "361B0F",
  white: "FFFFFF",
};

const FONT = {
  head: "Merriweather",
  body: "Source Sans Pro",
};

const imageDims = {
  "1.jpg": [5464, 8192],
  "2.jpg": [1920, 1080],
  "3.jpg": [1600, 1066],
  "4.jpg": [1066, 1600],
  "5.jpg": [6016, 4016],
  "6.jpg": [1536, 2048],
  "7.jpg": [1066, 1600],
  "8.jpg": [1366, 2048],
  "9.jpg": [1200, 1600],
  "10a.jpg": [1920, 1080],
  "10b.jpg": [1920, 1080],
  "11 (1).jpg": [1920, 1080],
  "11 (2).jpg": [1920, 1080],
  "11 (3).jpg": [1920, 1080],
  "12.jpg": [1200, 1600],
  "13.jpg": [1066, 1600],
  "14.jpg": [5464, 8192],
  "x.jpg": [3832, 3832],
};

const slides = [
  {
    kind: "title",
    image: "1.jpg",
    kicker: "Bc. Štěpánka Trappová",
    title: "Porod není jen medicína",
    subtitle: "aneb proč na porodu záleží",
    footer: "komunitní porodní asistentka",
  },
  {
    image: "2.jpg",
    title: "Kdo jsem a co dělám",
    lead: "Jsem porodní asistentka.",
    body: [
      "Doprovázím ženy těhotenstvím, porodem i po porodu.",
      "Nabízím předporodní přípravu, doprovod k porodu a poporodní péči.",
    ],
    quote: "Moje práce není jen o zdravotní péči. Je hlavně o podpoře, důvěře a bezpečí.",
  },
  {
    kind: "statement",
    image: "3.jpg",
    title: "Proč právě porod?",
    lead: "Protože porod není jen „lékařská událost“.",
    body: [
      "Je to silný životní moment.",
      "Zážitek, který si ženy pamatují celý život.",
      "Může být posilující, nebo naopak zraňující.",
    ],
  },
  {
    image: "4.jpg",
    title: "Jak si ženy porod představují",
    lead: "Podle filmů a médií:",
    body: [
      "žena leží na zádech",
      "hodně zásahů a léků",
      "rychlý, dramatický průběh",
      "často pasivní role ženy",
    ],
  },
  {
    image: "5.jpg",
    title: "Realita v systému",
    body: [
      "méně prostoru pro individuální přístup",
      "časový tlak",
      "zavedené postupy",
      "omezený prostor pro vytvoření důvěry",
    ],
    quote: "Ne proto, že by někdo chtěl ublížit. Ale protože systém má svoje limity.",
  },
  {
    image: "6.jpg",
    title: "Porod může vypadat i jinak",
    lead: "Různé polohy, více svobody pohybu a větší zapojení ženy.",
    body: ["gauč", "voda", "porodní stolička", "postel"],
    quote: "Každé ženě vyhovuje něco jiného.",
  },
  {
    kind: "statement",
    image: "7.jpg",
    title: "Neexistuje jeden „správný“ porod",
    body: ["každá žena je jiná", "každé tělo je jiné", "každá zkušenost je jiná"],
    quote: "Důležité je, aby se žena cítila bezpečně, respektovaně a vyslyšeně.",
  },
  {
    kind: "statement",
    image: "8.jpg",
    title: "Proč na porodu záleží",
    body: ["porod si žena pamatuje celý život", "většinou ho zažije jen 1-2x"],
    quote: "Není to „jen jeden den“.",
  },
  {
    image: "9.jpg",
    title: "Jak porod ovlivňuje ženu",
    body: ["sebevědomí", "vztah k sobě", "vztah k dítěti", "začátek kojení"],
    quote: "Může ji posílit... nebo naopak oslabit.",
  },
  {
    image: "10a.jpg",
    secondaryImage: "10b.jpg",
    title: "Proč dělám to, co dělám",
    lead: "Protože věřím, že porod může být posilující zážitek a ženy mají v sobě obrovskou sílu.",
    body: ["podporu", "respekt", "individuální přístup", "pochopení", "nesouzení"],
  },
  {
    image: "11 (1).jpg",
    secondaryImage: "11 (2).jpg",
    title: "Co dělá rozdíl",
    body: [
      "když žena není „jen další pacientka“",
      "když má kolem sebe známé lidi",
      "když má kontinuální péči",
    ],
    quote: "Důvěra = klíč k dobrému zážitku.",
  },
  {
    image: "12.jpg",
    title: "Co mě ovlivnilo",
    body: ["zkušenost z Německa", "vlastní porod", "ukázalo mi to, že věci mohou fungovat i jinak"],
  },
  {
    kind: "statement",
    image: "13.jpg",
    title: "Na porod nemusíte být sama",
    body: ["můžete mít podporu", "můžete mít informace", "můžete mít někoho na své straně"],
    quote: "A přesně to nabízím.",
  },
  {
    kind: "contact",
    image: "14.jpg",
    title: "Pokud vás to oslovilo...",
    lead: "Ráda vás podpořím. Můžete se na mě obrátit.",
    body: ["Bc. Štěpánka Trappová", "+420 605 074 332", "stepanka@trappea.cz", "www.trappea.cz"],
  },
];

const esc = (s) =>
  String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");

const emu = (v) => Math.round(v * EMU);
const hex = (h) => h.replace("#", "").toUpperCase();
const geom = (prst) =>
  prst === "roundRect"
    ? `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 3500"/></a:avLst></a:prstGeom>`
    : `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>`;

function cropFor(imageName, frameW, frameH) {
  const [iw, ih] = imageDims[imageName] || [frameW, frameH];
  const imgRatio = iw / ih;
  const frameRatio = frameW / frameH;
  if (imgRatio > frameRatio) {
    const visible = frameRatio / imgRatio;
    const crop = Math.round(((1 - visible) / 2) * 100000);
    return { l: crop, r: crop, t: 0, b: 0 };
  }
  const visible = imgRatio / frameRatio;
  const crop = Math.round(((1 - visible) / 2) * 100000);
  return { l: 0, r: 0, t: crop, b: crop };
}

function fill(color, alpha) {
  const alphaVal = alpha == null ? null : alpha > 1000 ? alpha : Math.round(alpha * 1000);
  const trans = alphaVal == null ? "" : `<a:alpha val="${alphaVal}"/>`;
  return `<a:solidFill><a:srgbClr val="${hex(color)}">${trans}</a:srgbClr></a:solidFill>`;
}

function noLine() {
  return `<a:ln><a:noFill/></a:ln>`;
}

function shape(id, x, y, w, h, color, opts = {}) {
  const prst = opts.prst || "rect";
  const alpha = opts.alpha;
  const line = opts.line ? `<a:ln w="${Math.round(opts.line.width * 12700)}">${fill(opts.line.color)}</a:ln>` : noLine();
  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>${geom(prst)}${fill(color, alpha)}${line}</p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>`;
}

function run(text, size, typeface, color, bold = false, italic = false) {
  return `<a:r><a:rPr lang="cs-CZ" sz="${Math.round(size * 100)}" b="${bold ? 1 : 0}" i="${italic ? 1 : 0}" dirty="0">${fill(color)}<a:latin typeface="${esc(typeface)}"/><a:cs typeface="${esc(typeface)}"/></a:rPr><a:t>${esc(text)}</a:t></a:r>`;
}

function paragraph(text, opts = {}) {
  const size = opts.size || 24;
  const font = opts.font || FONT.body;
  const color = opts.color || C.text;
  const bold = !!opts.bold;
  const italic = !!opts.italic;
  const algn = opts.align ? ` algn="${opts.align}"` : "";
  const marL = opts.bullet ? ` marL="285750" indent="-171450"` : "";
  const spcAft = opts.after != null ? `<a:spcAft><a:spcPts val="${opts.after}"/></a:spcAft>` : "";
  const lnSpc = opts.line != null ? `<a:lnSpc><a:spcPct val="${opts.line}"/></a:lnSpc>` : "";
  const pPr = `<a:pPr${algn}${marL}>${lnSpc}${spcAft}${opts.bullet ? '<a:buChar char="•"/>' : ""}</a:pPr>`;
  return `<a:p>${pPr}${run(text, size, font, color, bold, italic)}</a:p>`;
}

function textBox(id, x, y, w, h, paragraphs, opts = {}) {
  const inset = opts.inset || { l: 0, t: 0, r: 0, b: 0 };
  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>${noLine()}</p:spPr>
  <p:txBody><a:bodyPr wrap="square" anchor="${opts.anchor || "t"}" lIns="${emu(inset.l)}" tIns="${emu(inset.t)}" rIns="${emu(inset.r)}" bIns="${emu(inset.b)}"/><a:lstStyle/>${paragraphs.join("")}</p:txBody>
</p:sp>`;
}

function imagePic(id, relId, imageName, x, y, w, h, opts = {}) {
  const crop = opts.crop || cropFor(imageName, w, h);
  const srcRect = `<a:srcRect l="${crop.l}" r="${crop.r}" t="${crop.t}" b="${crop.b}"/>`;
  const picGeom = geom(opts.round ? "roundRect" : "rect");
  return `<p:pic>
  <p:nvPicPr><p:cNvPr id="${id}" name="${esc(imageName)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="${relId}"/>${srcRect}<a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>${picGeom}</p:spPr>
</p:pic>`;
}

function slideXml(content, bg = C.bgSoft) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgPr>${fill(bg)}</p:bgPr></p:bg><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    ${content.join("\n")}
  </p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}

function slideRels(imageRels) {
  const rels = [
    `<Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`,
    ...imageRels.map((r) => `<Relationship Id="${r.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${esc(r.target)}"/>`),
  ];
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;
}

function contentSlide(s, idx, mediaMap) {
  let id = 10;
  const imgRel = mediaMap.get(s.image).rId;
  const flip = idx % 2 === 0;
  const ix = flip ? 58 : 735;
  const tx = flip ? 655 : 92;
  const imgW = 485;
  const imgH = 585;
  const out = [];
  out.push(shape(id++, 0, 0, W, H, idx % 3 === 0 ? C.bgWarm : C.bgSoft));
  out.push(shape(id++, tx - 32, 84, 520, 548, C.white, { prst: "roundRect", alpha: 90000 }));
  out.push(imagePic(id++, imgRel, s.image, ix, 68, imgW, imgH, { round: true }));
  if (s.secondaryImage) {
    out.push(imagePic(id++, mediaMap.get(s.secondaryImage).rId, s.secondaryImage, ix + imgW - 170, 480, 150, 130, { round: true }));
  }
  out.push(textBox(id++, tx, 104, 472, 118, [paragraph(s.title, { size: 39, font: FONT.head, color: C.text, line: 112000 })]));
  out.push(shape(id++, tx, 232, 168, 2, C.cta, { prst: "roundRect" }));
  let y = 262;
  if (s.lead) {
    out.push(textBox(id++, tx, y, 470, 78, [paragraph(s.lead, { size: 23, font: FONT.body, color: C.accent, bold: true, line: 118000 })]));
    y += 92;
  }
  if (s.body?.length) {
    const bodyParas = s.body.map((b) => paragraph(b, { size: 22, bullet: true, after: 700, line: 120000 }));
    out.push(textBox(id++, tx, y, 472, Math.min(220, 42 * s.body.length + 12), bodyParas));
    y += Math.min(220, 42 * s.body.length + 12) + 18;
  }
  if (s.quote) {
    out.push(shape(id++, tx, Math.min(y, 540), 452, 78, C.bgWarm, { prst: "roundRect" }));
    out.push(textBox(id++, tx + 18, Math.min(y, 540) + 3, 34, 44, [paragraph("“", { size: 52, font: FONT.head, color: C.rose, line: 100000 })]));
    out.push(textBox(id++, tx + 58, Math.min(y, 540) + 14, 368, 58, [paragraph(s.quote, { size: 21, font: FONT.head, color: C.cta, line: 112000 })]));
  }
  out.push(textBox(id++, 1116, 666, 80, 24, [paragraph(String(idx).padStart(2, "0"), { size: 13, font: FONT.head, color: C.taupe, align: "r" })]));
  return out;
}

function titleSlide(s, mediaMap) {
  let id = 10;
  const out = [];
  out.push(shape(id++, 0, 0, W, H, C.bgWarm));
  out.push(shape(id++, 0, 0, 640, H, C.bgSoft));
  out.push(imagePic(id++, mediaMap.get(s.image).rId, s.image, 702, 44, 496, 632, { round: true }));
  out.push(textBox(id++, 76, 104, 560, 34, [paragraph(s.kicker, { size: 20, font: FONT.body, color: C.accent, bold: true })]));
  out.push(textBox(id++, 76, 148, 560, 34, [paragraph(s.footer, { size: 17, font: FONT.body, color: C.accent })]));
  out.push(textBox(id++, 76, 224, 565, 166, [paragraph(s.title, { size: 58, font: FONT.head, color: C.text, line: 106000 })]));
  out.push(shape(id++, 76, 410, 168, 2, C.cta, { prst: "roundRect" }));
  out.push(textBox(id++, 76, 442, 565, 80, [paragraph(s.subtitle, { size: 32, font: FONT.head, color: C.accent, line: 112000 })]));
  return out;
}

function statementSlide(s, idx, mediaMap) {
  let id = 10;
  const out = [];
  out.push(shape(id++, 0, 0, W, H, C.bg));
  out.push(imagePic(id++, mediaMap.get(s.image).rId, s.image, 0, 0, W, H));
  out.push(shape(id++, 0, 0, W, H, C.bgWarm, { alpha: 76000 }));
  out.push(shape(id++, 91, 92, 700, 500, C.white, { prst: "roundRect", alpha: 92000 }));
  out.push(textBox(id++, 135, 134, 610, 128, [paragraph(s.title, { size: 54, font: FONT.head, color: C.text, line: 108000 })]));
  out.push(shape(id++, 145, 274, 168, 2, C.cta, { prst: "roundRect" }));
  let y = 310;
  if (s.lead) {
    out.push(textBox(id++, 145, y, 540, 62, [paragraph(s.lead, { size: 24, font: FONT.body, color: C.accent, bold: true, line: 118000 })]));
    y += 78;
  }
  out.push(textBox(id++, 145, y, 510, 112, s.body.map((b) => paragraph(b, { size: 25, bullet: true, after: 900 }))));
  if (s.quote) {
    out.push(shape(id++, 145, 468, 500, 92, C.bgWarm, { prst: "roundRect" }));
    out.push(textBox(id++, 163, 470, 42, 60, [paragraph("“", { size: 58, font: FONT.head, color: C.rose, line: 100000 })]));
    out.push(textBox(id++, 212, 490, 400, 62, [paragraph(s.quote, { size: 31, font: FONT.head, color: C.cta, line: 110000 })]));
  }
  out.push(textBox(id++, 1116, 666, 80, 24, [paragraph(String(idx).padStart(2, "0"), { size: 13, font: FONT.head, color: C.white, align: "r" })]));
  return out;
}

function dividerSlide(s, mediaMap) {
  let id = 10;
  const out = [];
  out.push(shape(id++, 0, 0, W, H, C.bgWarm));
  out.push(imagePic(id++, mediaMap.get(s.image).rId, s.image, 410, 42, 460, 636, { round: true, crop: { l: 0, r: 0, t: 0, b: 0 } }));
  return out;
}

function contactSlide(s, idx, mediaMap, logoRel) {
  let id = 10;
  const out = [];
  out.push(shape(id++, 0, 0, W, H, C.bgWarm));
  out.push(imagePic(id++, mediaMap.get(s.image).rId, s.image, 746, 66, 384, 384, { round: true }));
  out.push(shape(id++, 92, 90, 590, 540, C.white, { prst: "roundRect", alpha: 94000 }));
  out.push(textBox(id++, 136, 132, 500, 110, [paragraph(s.title, { size: 48, font: FONT.head, color: C.text, line: 108000 })]));
  out.push(shape(id++, 140, 256, 168, 2, C.cta, { prst: "roundRect" }));
  out.push(textBox(id++, 140, 285, 455, 70, [paragraph(s.lead, { size: 24, font: FONT.body, color: C.accent, line: 120000 })]));
  out.push(shape(id++, 140, 382, 80, 4, C.cta, { prst: "roundRect" }));
  out.push(textBox(id++, 140, 402, 455, 148, s.body.map((b, i) => paragraph(b, { size: i === 0 ? 24 : 22, font: i === 0 ? FONT.head : FONT.body, color: i === 0 ? C.text : C.accent, after: 600 }))));
  if (logoRel) out.push(imagePic(id++, logoRel, "logo.png", 872, 528, 232, 68, { crop: { l: 0, r: 0, t: 0, b: 0 } }));
  out.push(textBox(id++, 1116, 666, 80, 24, [paragraph(String(idx).padStart(2, "0"), { size: 13, font: FONT.head, color: C.taupe, align: "r" })]));
  return out;
}

async function listFiles(dir) {
  const out = [];
  for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
    const p = path.join(dir, entry.name);
    if (entry.isDirectory()) out.push(...(await listFiles(p)));
    else out.push(p);
  }
  return out;
}

async function writeFile(p, content) {
  await fs.mkdir(path.dirname(p), { recursive: true });
  await fs.writeFile(p, content);
}

async function main() {
  await fs.rm(TMP_DIR, { recursive: true, force: true });
  await fs.mkdir(TMP_DIR, { recursive: true });
  const ppt = path.join(TMP_DIR, "ppt");
  const relsDir = path.join(TMP_DIR, "_rels");
  const docProps = path.join(TMP_DIR, "docProps");
  const mediaDir = path.join(ppt, "media");
  await fs.mkdir(path.join(ppt, "slides", "_rels"), { recursive: true });
  await fs.mkdir(path.join(ppt, "slideLayouts", "_rels"), { recursive: true });
  await fs.mkdir(path.join(ppt, "slideMasters", "_rels"), { recursive: true });
  await fs.mkdir(path.join(ppt, "theme"), { recursive: true });
  await fs.mkdir(mediaDir, { recursive: true });
  await fs.mkdir(relsDir, { recursive: true });
  await fs.mkdir(docProps, { recursive: true });

  const usedImages = new Set();
  for (const s of slides) {
    if (s.image) usedImages.add(s.image);
    if (s.secondaryImage) usedImages.add(s.secondaryImage);
  }

  const mediaMap = new Map();
  let mediaId = 1;
  for (const name of usedImages) {
    const src = path.join(ROOT, "foto prezentace", name);
    const target = `image${mediaId}${path.extname(name).toLowerCase()}`;
    await fs.copyFile(src, path.join(mediaDir, target));
    mediaMap.set(name, { rId: `rIdImg${mediaId}`, target });
    mediaId += 1;
  }

  let logoRel = null;
  try {
    const logoRoot = path.join(ROOT, "ŠT LOGO ROZŠÍŘENÝ O NÁPIS - FINÁLNÍ BALÍK");
    const files = await listFiles(logoRoot);
    const logo = files.find((f) => f.endsWith(".png") && f.includes("VÍNOVÉ_TEXT")) || files.find((f) => f.endsWith(".png") && f.includes("TEXT")) || files.find((f) => f.endsWith(".png"));
    if (logo) {
      const target = `image${mediaId}.png`;
      await fs.copyFile(logo, path.join(mediaDir, target));
      logoRel = `rIdImg${mediaId}`;
      mediaMap.set("logo.png", { rId: logoRel, target });
    }
  } catch {}

  const slideIdList = [];
  const presRels = [
    `<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>`,
    `<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`,
  ];

  for (let i = 0; i < slides.length; i++) {
    const s = slides[i];
    let content;
    if (s.kind === "title") content = titleSlide(s, mediaMap);
    else if (s.kind === "statement") content = statementSlide(s, i + 1, mediaMap);
    else if (s.kind === "divider") content = dividerSlide(s, mediaMap);
    else if (s.kind === "contact") content = contactSlide(s, i + 1, mediaMap, logoRel);
    else content = contentSlide(s, i + 1, mediaMap);

    await writeFile(path.join(ppt, "slides", `slide${i + 1}.xml`), slideXml(content));
    const imgRels = [];
    if (s.image) imgRels.push(mediaMap.get(s.image));
    if (s.secondaryImage) imgRels.push(mediaMap.get(s.secondaryImage));
    if (s.kind === "contact" && logoRel) imgRels.push(mediaMap.get("logo.png"));
    await writeFile(path.join(ppt, "slides", "_rels", `slide${i + 1}.xml.rels`), slideRels(imgRels));
    slideIdList.push(`<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`);
    presRels.push(`<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`);
  }

  await writeFile(path.join(TMP_DIR, "[Content_Types].xml"), contentTypes(slides.length, mediaMap));
  await writeFile(path.join(relsDir, ".rels"), rootRels());
  await writeFile(path.join(docProps, "core.xml"), coreProps());
  await writeFile(path.join(docProps, "app.xml"), appProps(slides.length));
  await writeFile(path.join(ppt, "presentation.xml"), presentationXml(slideIdList));
  await writeFile(path.join(ppt, "_rels", "presentation.xml.rels"), relsXml(presRels));
  await writeFile(path.join(ppt, "viewProps.xml"), viewProps());
  await writeFile(path.join(ppt, "presProps.xml"), presProps());
  await writeFile(path.join(ppt, "tableStyles.xml"), tableStyles());
  await writeFile(path.join(ppt, "theme", "theme1.xml"), themeXml());
  await writeFile(path.join(ppt, "slideMasters", "slideMaster1.xml"), slideMasterXml());
  await writeFile(path.join(ppt, "slideMasters", "_rels", "slideMaster1.xml.rels"), slideMasterRels());
  await writeFile(path.join(ppt, "slideLayouts", "slideLayout1.xml"), slideLayoutXml());
  await writeFile(path.join(ppt, "slideLayouts", "_rels", "slideLayout1.xml.rels"), slideLayoutRels());

  await fs.rm(PPTX_PATH, { force: true });
  await execFileAsync("zip", ["-qr", PPTX_PATH, "."], { cwd: TMP_DIR });
  console.log(PPTX_PATH);
}

function relsXml(rels) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;
}

function rootRels() {
  return relsXml([
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>`,
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`,
    `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>`,
  ]);
}

function contentTypes(slideCount, mediaMap) {
  const imageDefaults = new Set([...mediaMap.values()].map((m) => path.extname(m.target).slice(1)));
  const defaults = [
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
    `<Default Extension="xml" ContentType="application/xml"/>`,
    ...[...imageDefaults].map((ext) => `<Default Extension="${ext}" ContentType="image/${ext === "jpg" ? "jpeg" : ext}"/>`),
  ];
  const overrides = [
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>`,
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`,
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`,
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`,
    `<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>`,
    `<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>`,
    `<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>`,
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`,
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`,
  ];
  for (let i = 1; i <= slideCount; i++) {
    overrides.push(`<Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults.join("")}${overrides.join("")}</Types>`;
}

function presentationXml(slideIdList) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>
  <p:sldIdLst>${slideIdList.join("")}</p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle><a:defPPr><a:defRPr lang="cs-CZ"/></a:defPPr></p:defaultTextStyle>
</p:presentation>`;
}

function themeXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Trappea">
<a:themeElements><a:clrScheme name="Trappea"><a:dk1><a:srgbClr val="${C.text}"/></a:dk1><a:lt1><a:srgbClr val="${C.bg}"/></a:lt1><a:dk2><a:srgbClr val="${C.accent}"/></a:dk2><a:lt2><a:srgbClr val="${C.bgWarm}"/></a:lt2><a:accent1><a:srgbClr val="${C.cta}"/></a:accent1><a:accent2><a:srgbClr val="${C.rose}"/></a:accent2><a:accent3><a:srgbClr val="${C.taupe}"/></a:accent3><a:accent4><a:srgbClr val="${C.accent}"/></a:accent4><a:accent5><a:srgbClr val="${C.bgSoft}"/></a:accent5><a:accent6><a:srgbClr val="${C.bgWarm}"/></a:accent6><a:hlink><a:srgbClr val="${C.cta}"/></a:hlink><a:folHlink><a:srgbClr val="${C.accent}"/></a:folHlink></a:clrScheme><a:fontScheme name="Trappea"><a:majorFont><a:latin typeface="${FONT.head}"/><a:ea typeface=""/><a:cs typeface="${FONT.head}"/></a:majorFont><a:minorFont><a:latin typeface="${FONT.body}"/><a:ea typeface=""/><a:cs typeface="${FONT.body}"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

function slideMasterXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`;
}

function slideMasterRels() {
  return relsXml([
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`,
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>`,
  ]);
}

function slideLayoutXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function slideLayoutRels() {
  return relsXml([`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>`]);
}

function coreProps() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Porod není jen medicína</dc:title><dc:creator>OpenAI</dc:creator><cp:lastModifiedBy>OpenAI</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2026-05-13T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-05-13T00:00:00Z</dcterms:modified></cp:coreProperties>`;
}

function appProps(count) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft PowerPoint</Application><PresentationFormat>On-screen Show (16:9)</PresentationFormat><Slides>${count}</Slides><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Trappea</vt:lpstr></vt:vector></TitlesOfParts></Properties>`;
}

function viewProps() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr></p:viewPr>`;
}

function presProps() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`;
}

function tableStyles() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
