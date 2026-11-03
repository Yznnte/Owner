const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ImageRun,
  ShadingType, BorderStyle, AlignmentType, HeadingLevel, LevelFormat, Header, Footer,
  PageNumber, PageBreak, VerticalAlign,
} = require("docx");

const OUT = "/home/user/Owner/Schulung_Word";
const IMG = path.join(__dirname, "bilder");
const NAVY = "1F3A5F", ACCENT = "F28C28", GREY = "6B778C", LIGHT = "EEF3FB", LINE = "D5DDE8";
const FONT = "Calibri";
const M = 850;                       // 1.5 cm margins
const W = 11906 - 2 * M;             // text width in DXA
const WPX = Math.floor(W / 15);      // text width in px (1px = 15 DXA)

const r = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
const P = (c, o = {}) => new Paragraph({ children: typeof c === "string" ? [r(c)] : c, ...o });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [r(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [r(t)] });
const pb = () => new Paragraph({ children: [new PageBreak()] });
const sp = (after = 120) => new Paragraph({ children: [], spacing: { after } });
const bullet = (c) => new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 60 }, children: typeof c === "string" ? [r(c)] : c });
const check = (c) => new Paragraph({ numbering: { reference: "chk", level: 0 }, spacing: { after: 90 }, children: typeof c === "string" ? [r(c, { size: 24 })] : c });
let numCount = 0;
const numList = (items) => { const ref = "n" + (numCount++ % 8); return items.map((t) => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80 }, children: typeof t === "string" ? [r(t, { size: 24 })] : t })); };

function pngSize(file) { const b = fs.readFileSync(file); return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) }; }
function img(name, widthPx = WPX, o = {}) {
  const f = path.join(IMG, name + ".png"); const s = pngSize(f);
  return new Paragraph({ alignment: AlignmentType.CENTER, keepNext: o.keepNext, spacing: { before: o.before ?? 60, after: o.after ?? 160 },
    children: [new ImageRun({ type: "png", data: fs.readFileSync(f), transformation: { width: widthPx, height: Math.round(widthPx * s.h / s.w) } })] });
}

const thin = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const tb = { top: thin, bottom: thin, left: thin, right: thin };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const nb = { top: none, bottom: none, left: none, right: none };
function cell(c, w, o = {}) {
  return new TableCell({ width: { size: w, type: WidthType.DXA }, borders: o.borders || tb, verticalAlign: VerticalAlign.CENTER,
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined, margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ alignment: o.align, spacing: { after: 0 }, children: typeof c === "string" ? [r(c, o.run || {})] : c })] });
}
function table(widths, head, rows, o = {}) {
  const tr = [];
  if (head) tr.push(new TableRow({ tableHeader: true, children: head.map((h, i) => cell(h, widths[i], { fill: NAVY, run: { bold: true, color: "FFFFFF" }, align: o.headAlign?.[i] })) }));
  rows.forEach((row, ri) => tr.push(new TableRow({ cantSplit: true, children: row.map((c, i) => cell(c, widths[i], { fill: ri % 2 ? "F6F8FB" : undefined, run: i === 0 && o.boldFirst ? { bold: true, color: NAVY } : {}, align: o.align?.[i] })) })));
  return new Table({ width: { size: widths.reduce((a, b) => a + b), type: WidthType.DXA }, columnWidths: widths, rows: tr });
}
function box(paras, fill = "FFF4E8", edge = ACCENT) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [new TableCell({
    width: { size: W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill, color: "auto" },
    borders: { ...nb, left: { style: BorderStyle.SINGLE, size: 36, color: edge } }, margins: { top: 140, bottom: 140, left: 220, right: 220 },
    children: paras.map((p) => p instanceof Paragraph ? p : P(p, { spacing: { after: 40 } })) })] })] });
}

const styles = {
  default: { document: { run: { font: FONT, size: 23 }, paragraph: { spacing: { after: 120, line: 276, lineRule: "auto" } } } },
  paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 34, bold: true, color: NAVY },
      paragraph: { spacing: { before: 240, after: 120 }, keepNext: true, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 26, bold: true, color: ACCENT },
      paragraph: { spacing: { before: 180, after: 80 }, keepNext: true, outlineLevel: 1 } },
  ],
};
const numbering = { config: [
  { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "●", alignment: AlignmentType.LEFT, style: { run: { color: ACCENT, size: 18 }, paragraph: { indent: { left: 460, hanging: 280 } } } }] },
  { reference: "chk", levels: [{ level: 0, format: LevelFormat.BULLET, text: "☐", alignment: AlignmentType.LEFT, style: { run: { size: 30, color: NAVY }, paragraph: { indent: { left: 520, hanging: 400 } } } }] },
  ...Array.from({ length: 8 }, (_, i) => ({ reference: "n" + i, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
    style: { run: { bold: true, color: ACCENT, size: 24 }, paragraph: { indent: { left: 520, hanging: 400 } } } }] })),
] };
const header = (label) => new Header({ children: [new Paragraph({ tabStops: [{ type: "right", position: W }],
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } },
  children: [r("Die Allrounder", { bold: true, color: NAVY, size: 18 }), r("  |  Word-Schulung", { color: GREY, size: 18 }), r("\t" + label, { color: GREY, size: 18 })] })] });
const footer = () => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
  r("Seite ", { size: 18, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], size: 18, color: GREY, font: FONT }),
  r(" von ", { size: 18, color: GREY }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: GREY, font: FONT }),
  r("  ·  Nico & Yannik", { size: 18, color: GREY })] })] });
const page = { size: { width: 11906, height: 16838 }, margin: { top: 850, bottom: 850, left: M, right: M, header: 425, footer: 425 } };
const sec = (label, children, o = {}) => ({ properties: { page: o.page || page }, headers: label ? { default: header(label) } : undefined, footers: { default: o.footer || footer() }, children });
async function save(name, sections) {
  const doc = new Document({ creator: "Nico & Yannik – Die Allrounder", title: name, styles, numbering, sections });
  fs.writeFileSync(path.join(OUT, name + ".docx"), await Packer.toBuffer(doc)); console.log("wrote", name);
}
const simpleFooter = () => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [r("Team „Die Allrounder“  ·  Nico & Yannik", { size: 18, color: GREY })] })] });

(async () => {
// ---------------- 01 EINLADUNG ----------------
await save("01_Einladung", [sec(null, [
  img("hero", WPX, { before: 0, after: 240 }),
  P([r("Liebe Schülerinnen und Schüler,", { bold: true, size: 26 })]),
  P([r("wir zeigen euch, wie ihr in Word in wenigen Minuten richtig schöne Dokumente macht. Ihr probiert alles direkt selbst am PC aus. Ihr braucht keine Vorkenntnisse.", { size: 24 })], { spacing: { after: 200 } }),
  img("tiles", WPX, { after: 120 }),
  H2("Das lernt ihr"),
  img("lernen", WPX, { before: 0, after: 200 }),
  box([[r("✋  Bitte meldet euch bis "), r("Freitag, 13.11.2026", { bold: true }), r(" bei Nico oder Yannik an.")]]),
  P([r("Wir freuen uns auf euch!  ", { size: 24 }), r("Nico & Yannik", { bold: true, color: NAVY, size: 26 })], { spacing: { before: 280 } }),
], { footer: simpleFooter() })]);

// ---------------- 02 SCHULUNGSBLATT ----------------
await save("02_Schulungsblatt", [sec("Schulungsblatt", [
  img("ban_schulung", WPX, { before: 0, after: 200 }),
  H1("Unser Ablauf"),
  img("timeline"),
  H1("Das könnt ihr danach"),
  img("lernen"),
  pb(),
  H1("1  Formate"),
  img("rib_format"),
  img("zeichen_absatz"),
  img("pinsel"),
  pb(),
  img("shortcuts", 560),
  H1("2  Formatvorlagen"),
  img("ohne_mit", 560),
  pb(),
  img("rib_vorlagen"),
  img("aendern"),
  pb(),
  img("toc"),
  img("navi"),
  pb(),
  H1("3  Kopf- und Fußzeile"),
  img("kopffuss", 560),
  img("rib_einfuegen"),
  pb(),
  img("erste_seite"),
  H1("💻  Jetzt ihr!"),
  P([r("Öffnet ", { size: 24 }), r("04_Uebungsdatei_Steckbrief.docx", { bold: true, size: 24 }), r(" und hakt ab:", { size: 24 })]),
  check("Titel „Mein Steckbrief“ → Vorlage „Titel“"),
  check("Über mich, Hobbys, Traumberuf → „Überschrift 1“"),
  check("„Überschrift 1“ ändern: neue Farbe"),
  check("Nach dem Titel Strg + Enter → Inhaltsverzeichnis"),
  check("Kopfzeile: Name links, Klasse rechts"),
  check("Fußzeile: Seitenzahl"),
  check("☑ Erste Seite anders"),
  pb(),
  img("ziel"),
  H1("❓  Quiz"),
  ...numList([
    "Zeichenformat oder Absatzformat: Was ist „zentriert“?",
    "Warum braucht ein Inhaltsverzeichnis Formatvorlagen?",
    "Wie öffnet ihr die Kopfzeile am schnellsten?",
    "Warum tippt man Seitenzahlen nie von Hand?",
  ]),
])]);

// ---------------- 03 HANDOUT ----------------
const hp = (title, imgName, line, w = WPX) => [H1(title), img(imgName, w, { before: 0, after: 60 }), P([r(line, { size: 22, color: GREY })], { alignment: AlignmentType.CENTER, spacing: { after: 140 } })];
await save("03_Handout", [sec("Handout", [
  img("ban_handout", WPX, { before: 0, after: 120 }),
  ...hp("1  Formate", "rib_format", "Text markieren → Start → Schrift oder Absatz wählen"),
  ...hp("2  Formatvorlagen", "rib_vorlagen", "Ändern: Rechtsklick auf die Vorlage → „Ändern …“  ·  Inhaltsverzeichnis: Referenzen"),
  pb(),
  ...hp("3  Kopf- und Fußzeile", "rib_einfuegen", "Schneller: Doppelklick in den oberen oder unteren Rand"),
  img("shortcuts", 520, { after: 100 }),
  img("regeln", WPX, { after: 0 }),
])]);

// ---------------- 04 ÜBUNGSDATEI (Steckbrief) ----------------
const plain = (t) => new Paragraph({ children: [new TextRun(t)] });
await save("04_Uebungsdatei_Steckbrief", [{ properties: { page }, children: [
  "Mein Steckbrief",
  "Über mich",
  "Ich heiße … und bin … Jahre alt. Ich gehe in die Klasse … .",
  "Hobbys",
  "In meiner Freizeit mache ich gerne … .",
  "Traumberuf",
  "Später möchte ich gerne … werden, weil … .",
].map(plain) }]);

// ---------------- 05 FEEDBACK ----------------
const fw = [4906, 1060, 1060, 1060, 1060, 1060];
const smile = ["😃", "🙂", "😐", "🙁", "😞"];
const fq = ["Die Schulung war gut organisiert.", "Alles wurde verständlich erklärt.", "Das Tempo war passend.", "Die Übungen am PC haben geholfen.", "Ich kann jetzt Formatvorlagen nutzen.", "Ich kann jetzt Kopf- und Fußzeile einfügen."];
const lines = (n) => Array.from({ length: n }, () => new Paragraph({ spacing: { before: 300, after: 0 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA", space: 1 } }, children: [] }));
await save("05_Feedbackbogen", [sec("Feedback", [
  img("ban_feedback", WPX, { before: 0, after: 240 }),
  table(fw, ["Kreuzt an:", ...smile], fq.map((q) => [q, "☐", "☐", "☐", "☐", "☐"]), { headAlign: [null, ...Array(5).fill(AlignmentType.CENTER)], align: [null, ...Array(5).fill(AlignmentType.CENTER)] }),
  H2("👍  Das war gut:"), ...lines(2),
  H2("💡  Das geht besser:"), ...lines(2),
  P([r("Danke! 🙏", { bold: true, color: NAVY, size: 26 })], { alignment: AlignmentType.RIGHT, spacing: { before: 300 } }),
])]);

// ---------------- 06 BESCHEINIGUNG ----------------
const sig = (n) => new TableCell({ width: { size: 4300, type: WidthType.DXA }, borders: nb, children: [
  new Paragraph({ spacing: { before: 700, after: 0 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "333333", space: 1 } }, children: [] }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [r(n, { size: 20, color: GREY })] })] });
const dbl = { style: BorderStyle.DOUBLE, size: 12, color: NAVY, space: 24 };
await save("06_Teilnahmebescheinigung", [sec(null, [
  img("siegel", 190, { before: 200, after: 200 }),
  P([r("TEILNAHMEBESCHEINIGUNG", { bold: true, color: ACCENT, size: 30 })], { alignment: AlignmentType.CENTER, spacing: { after: 400 } }),
  P([r("Hiermit bestätigen wir, dass", { size: 26 })], { alignment: AlignmentType.CENTER }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 600, after: 60 }, indent: { left: 1500, right: 1500 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 1 } }, children: [] }),
  P([r("Name, Klasse", { size: 18, color: GREY })], { alignment: AlignmentType.CENTER, spacing: { after: 400 } }),
  P([r("am 17. November 2026 an der", { size: 26 })], { alignment: AlignmentType.CENTER }),
  P([r("Word-Schulung", { bold: true, color: NAVY, size: 52 })], { alignment: AlignmentType.CENTER, spacing: { before: 120, after: 40 } }),
  P([r("Formate · Formatvorlagen · Kopf- und Fußzeile", { color: GREY, size: 26 })], { alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
  P([r("erfolgreich teilgenommen hat.", { size: 26 })], { alignment: AlignmentType.CENTER, spacing: { after: 400 } }),
  new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [4300, W - 8600, 4300], rows: [new TableRow({ children: [
    sig("Nico"), new TableCell({ width: { size: W - 8600, type: WidthType.DXA }, borders: nb, children: [P("")] }), sig("Yannik")] })] }),
], { footer: simpleFooter(), page: { ...page, borders: { pageBorderTop: dbl, pageBorderBottom: dbl, pageBorderLeft: dbl, pageBorderRight: dbl } } })]);

// ---------------- 00 ABLAUFPLAN (intern) ----------------
await save("00_Ablaufplan_intern", [sec("Ablaufplan (intern)", [
  img("ban_ablauf", WPX, { before: 0, after: 100 }),
  img("timeline", 440, { after: 0 }),
  H1("Wer macht was?"),
  table([1300, 2300, 4706, 1900], ["Zeit", "Teil", "Was passiert", "Wer"], [
    ["08:30", "Start", "Begrüßen, Ablauf zeigen, Fragen sammeln", "Yannik"],
    ["08:40", "Formate", "Live zeigen: Schrift, Absatz, Pinsel", "Nico"],
    ["08:55", "Formatvorlagen", "Vorher/Nachher, Vorlage ändern, Inhaltsverzeichnis", "Yannik"],
    ["09:10", "Kopf-/Fußzeile", "Doppelklick, Seitenzahl, Erste Seite anders", "Nico"],
    ["09:20", "Üben", "Steckbrief in Gruppen (3–5), wir helfen", "Beide"],
    ["09:45", "Quiz", "Fragen gemeinsam lösen", "Beide"],
    ["09:52", "Ende", "Handout, Feedback, Bescheinigung", "Beide"],
  ], { boldFirst: true }),
  H1("Checkliste"),
  check("Einladung verteilt, Anmeldungen gezählt"),
  check("Raum 0.68: PCs, Beamer, Word getestet, gelüftet"),
  check("Übungsdatei auf allen PCs"),
  check("Gedruckt: Schulungsblatt, Handout, Feedback, Bescheinigungen"),
  check("Einmal komplett geprobt (mit Uhr)"),
  box([[r("💡 Tipps: ", { bold: true }), r("Frei sprechen · langsam klicken und ansagen · Fragen stellen · Handout erst am Ende austeilen")]], LIGHT, NAVY),
  H1("Quiz-Lösungen"),
  ...numList([
    "Absatzformat – gilt für den ganzen Absatz.",
    "Word erkennt Überschriften nur an ihrer Formatvorlage.",
    "Doppelklick in den oberen Seitenrand.",
    "Getippte Zahlen passen sich nicht an, wenn Seiten dazukommen.",
  ]),
])]);
})();
