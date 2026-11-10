const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ImageRun,
  ShadingType, BorderStyle, AlignmentType, HeadingLevel, LevelFormat, Header, Footer,
  PageNumber, PageBreak, VerticalAlign,
} = require("docx");

const OUT = process.env.OUT || "/home/user/Owner/Schulung_Word";
const IMG = process.env.IMG || path.join(__dirname, "bilder");
const NAVY = "1F3A5F", ACCENT = "E07A1F", GREY = "5F6B7D", BLUEBG = "EEF3FB", ORGBG = "FFF4E8", LINE = "D5DDE8";
const FONT = "Arial";
const M = 850;
const W = 11906 - 2 * M;
const WPX = Math.floor(W / 15);

// ---------- basic helpers ----------
const r = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
const P = (c, o = {}) => new Paragraph({ children: typeof c === "string" ? [r(c)] : c, ...o });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [r(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [r(t)] });
const pb = () => new Paragraph({ children: [new PageBreak()] });
const runs = (c, o = {}) => (typeof c === "string" ? [r(c, o)] : c);
const check = (c) => new Paragraph({ numbering: { reference: "chk", level: 0 }, spacing: { after: 100 }, children: runs(c, { size: 24 }) });
let nc = 0;
const steps = (items, size = 24) => { const ref = "n" + (nc++ % 30); return items.map((t) => new Paragraph({ numbering: { reference: ref, level: 0 }, keepNext: true, keepLines: true, spacing: { after: 70 }, children: runs(t, { size }) })); };
// Menu path like:  Start › Formatvorlagen
const path_ = (...items) => { const out = []; items.forEach((it, i) => { if (i) out.push(r("  ›  ", { bold: true, color: ACCENT })); out.push(r(it, { bold: true, color: NAVY })); }); return out; };

function pngSize(f) { const b = fs.readFileSync(f); return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) }; }
function imgRun(name, widthPx) { const f = path.join(IMG, name + ".png"); const s = pngSize(f);
  return new ImageRun({ type: "png", data: fs.readFileSync(f), transformation: { width: widthPx, height: Math.round(widthPx * s.h / s.w) } }); }
const img = (name, widthPx = WPX, o = {}) => new Paragraph({ alignment: AlignmentType.CENTER, keepNext: o.keepNext,
  spacing: { before: o.before ?? 80, after: o.after ?? 160, line: 240, lineRule: "auto" }, children: [imgRun(name, widthPx)] });

const thin = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const tb = { top: thin, bottom: thin, left: thin, right: thin };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const nb = { top: none, bottom: none, left: none, right: none };
function cell(c, w, o = {}) {
  const kids = Array.isArray(c) && c[0] instanceof Paragraph ? c : [new Paragraph({ alignment: o.align, spacing: { after: 0 }, children: runs(c, o.run || {}) })];
  return new TableCell({ width: { size: w, type: WidthType.DXA }, borders: o.borders || tb, verticalAlign: o.valign || VerticalAlign.CENTER,
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined, margins: o.margins || { top: 80, bottom: 80, left: 120, right: 120 }, children: kids });
}
function table(widths, head, rows, o = {}) {
  const tr = [];
  if (head) tr.push(new TableRow({ tableHeader: true, children: head.map((h, i) => cell(h, widths[i], { fill: NAVY, run: { bold: true, color: "FFFFFF" }, align: o.align?.[i] })) }));
  rows.forEach((row, ri) => tr.push(new TableRow({ cantSplit: true, children: row.map((c, i) => cell(c, widths[i], { fill: ri % 2 ? "F6F8FB" : undefined, run: i === 0 && o.boldFirst ? { bold: true, color: NAVY } : {}, align: o.align?.[i] })) })));
  return new Table({ width: { size: widths.reduce((a, b) => a + b), type: WidthType.DXA }, columnWidths: widths, rows: tr });
}
function box(paras, fill = ORGBG, edge = ACCENT) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ cantSplit: true, children: [new TableCell({
    width: { size: W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill, color: "auto" },
    borders: { ...nb, left: { style: BorderStyle.SINGLE, size: 36, color: edge } }, margins: { top: 120, bottom: 120, left: 220, right: 220 },
    children: paras.map((p) => (p instanceof Paragraph ? p : P(p, { spacing: { after: 40 } }))) })] })] });
}
// "Was ist das?" box – one sentence
const was = (text) => box([P([r("Was ist das?  ", { bold: true, color: NAVY, size: 24 }), r(text, { size: 24 })], { spacing: { after: 0 } })], BLUEBG, "3D73C4");
const soGehts = () => P([r("So geht’s:", { bold: true, color: ACCENT, size: 24 })], { spacing: { before: 160, after: 60 }, keepNext: true });

const styles = {
  default: { document: { run: { font: FONT, size: 23 }, paragraph: { spacing: { after: 120, line: 276, lineRule: "auto" } } } },
  paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 36, bold: true, color: NAVY },
      paragraph: { spacing: { before: 240, after: 120, line: 240, lineRule: "auto" }, keepNext: true, outlineLevel: 0,
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 4 } } } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 27, bold: true, color: ACCENT },
      paragraph: { spacing: { before: 200, after: 80, line: 240, lineRule: "auto" }, keepNext: true, outlineLevel: 1 } },
  ],
};
const numbering = { config: [
  { reference: "chk", levels: [{ level: 0, format: LevelFormat.BULLET, text: "☐", alignment: AlignmentType.LEFT, style: { run: { font: "Segoe UI Symbol", size: 28, color: NAVY }, paragraph: { indent: { left: 520, hanging: 400 } } } }] },
  ...Array.from({ length: 30 }, (_, i) => ({ reference: "n" + i, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
    style: { run: { font: FONT, bold: true, color: ACCENT, size: 24 }, paragraph: { indent: { left: 460, hanging: 360 } } } }] })),
] };
const header = (label) => new Header({ children: [new Paragraph({ tabStops: [{ type: "right", position: W }],
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } },
  children: [r("Team „Die Allrounder“", { bold: true, color: NAVY, size: 18 }), r("  |  Word-Schulung", { color: GREY, size: 18 }), r("\t" + label, { color: GREY, size: 18 })] })] });
const footer = () => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
  r("Seite ", { size: 18, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], size: 18, color: GREY, font: FONT }),
  r(" von ", { size: 18, color: GREY }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: GREY, font: FONT }),
  r("  ·  Team „Die Allrounder“ (Nico und Yannik)", { size: 18, color: GREY })] })] });
const simpleFooter = () => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [r("Team „Die Allrounder“  ·  Nico und Yannik", { size: 18, color: GREY })] })] });
const page = { size: { width: 11906, height: 16838 }, margin: { top: 850, bottom: 850, left: M, right: M, header: 425, footer: 425 } };
const sec = (label, children, o = {}) => ({ properties: { page: o.page || page }, headers: label ? { default: header(label) } : undefined, footers: { default: o.footer || footer() }, children });
async function save(name, sections) {
  const doc = new Document({ creator: "Team „Die Allrounder“ – Nico und Yannik", title: name, styles, numbering, sections });
  fs.writeFileSync(path.join(OUT, name + ".docx"), await Packer.toBuffer(doc)); console.log("wrote", name);
}

(async () => {
// =========================== 01 EINLADUNG ===========================
await save("01_Einladung", [sec(null, [
  img("hero", WPX, { before: 0, after: 260 }),
  P([r("Liebe Schülerinnen und Schüler,", { bold: true, size: 26 })]),
  P([r("wir, das Team „Die Allrounder“, zeigen euch, wie ihr in Word schnell schöne und ordentliche Dokumente macht – zum Beispiel für Referate und Projekte. Ihr probiert es direkt am PC aus. Vorkenntnisse braucht ihr nicht.", { size: 24 })], { spacing: { after: 220 } }),
  img("tiles", WPX, { after: 100 }),
  H2("Das lernt ihr"),
  img("lernen", WPX, { before: 0, after: 220 }),
  P([r("Wir freuen uns auf euch!", { size: 24 })], { spacing: { before: 300, after: 40 } }),
  P([r("Nico und Yannik", { bold: true, color: NAVY, size: 28 })], { spacing: { after: 0 } }),
  P([r("Team „Die Allrounder“", { color: GREY, size: 24 })]),
], { footer: simpleFooter() })]);

// =========================== 02 SCHULUNGSBLATT ===========================
await save("02_Schulungsblatt", [sec("Schulungsblatt", [
  img("ban_schulung", WPX, { before: 0, after: 200 }),
  H1("Unser Ablauf"),
  img("timeline"),
  H1("Das könnt ihr danach"),
  img("lernen"),
  pb(),

  H1("1  Formate"),
  was("Formate bestimmen, wie euer Text aussieht – zum Beispiel Schrift, Farbe oder Ausrichtung."),
  soGehts(),
  ...steps(["Text markieren.", [r("Oben auf ", { size: 24 }), ...path_("Start"), r(" klicken.", { size: 24 })], "Gewünschten Knopf anklicken."]),
  img("rib_format"),
  img("zeichen_absatz", 600),
  H2("Tipp: Format kopieren mit dem Pinsel"),
  img("pinsel"),
  H2("Schneller mit der Tastatur"),
  img("shortcuts", 600),
  pb(),

  H1("2  Formatvorlagen"),
  was("Eine Formatvorlage ist ein gespeichertes Aussehen mit Namen, zum Beispiel „Überschrift 1“."),
  img("ohne_mit", 600),
  soGehts(),
  ...steps(["In die Überschrift klicken.", [...path_("Start", "Formatvorlagen"), r(" öffnen.", { size: 24 })], "Vorlage anklicken – fertig."]),
  img("rib_vorlagen"),
  H2("Formatvorlage ändern"),
  ...steps(["Rechtsklick auf die Vorlage.", "„Ändern …“ wählen.", "Schrift, Größe oder Farbe einstellen und OK klicken."]),
  img("aendern"),

  H1("3  Gliederungsansicht"),
  was("Die Gliederung zeigt nur eure Überschriften – wie ein Plan von eurem Dokument."),
  soGehts(),
  ...steps([[...path_("Ansicht", "Gliederung"), r(" klicken.", { size: 24 })], "Mit ▲ ▼ ganze Kapitel samt Text verschieben.", "Mit ✕ „Gliederungsansicht schließen“."]),
  img("gliederung"),
  H1("4  Inhaltsverzeichnis"),
  was("Eine Liste aller Überschriften mit Seitenzahlen. Word erstellt sie automatisch."),
  soGehts(),
  ...steps(["An den Anfang des Dokuments klicken.", [...path_("Referenzen", "Inhaltsverzeichnis", "Automatische Tabelle 1"), r(".", { size: 24 })], [r("Etwas geändert? ", { size: 24 }), ...path_("Referenzen", "Tabelle aktualisieren"), r(".", { size: 24 })]]),
  img("toc"),
  pb(),

  H1("5  Kopf- und Fußzeile"),
  was("Kopf- und Fußzeile stehen oben und unten auf jeder Seite – zum Beispiel Name und Seitenzahl."),
  img("kopffuss", 560),
  soGehts(),
  ...steps(["Doppelklick in den oberen oder unteren Rand.", "Text eintippen, zum Beispiel Name und Klasse.", [r("Seitenzahl: ", { size: 24 }), ...path_("Einfügen", "Seitenzahl", "Seitenende"), r(".", { size: 24 })]]),
  img("rib_einfuegen"),
  H2("Extra-Knöpfe, sobald ihr im Rand seid"),
  img("erste_seite"),

  H1("6  Verschiedene Seitenformate"),
  was("Mit einem Abschnittsumbruch bekommt ein Teil des Dokuments ein eigenes Format, zum Beispiel eine Seite im Querformat."),
  soGehts(),
  ...steps([[...path_("Layout", "Umbrüche", "Nächste Seite"), r(".", { size: 24 })], [...path_("Layout", "Ausrichtung", "Querformat"), r(".", { size: 24 })], "Andere Kopfzeile nötig? „Mit vorheriger verknüpfen“ ausschalten."]),
  img("abschnitte"),
  pb(),

  H1("7  Eigene Dokumentvorlage"),
  was("Eine Dokumentvorlage (.dotx) ist eine Musterdatei. Jedes neue Dokument daraus ist schon fertig eingerichtet."),
  img("dokvorlage"),
  soGehts(),
  ...steps(["Dokument einmal fertig einrichten (Formatvorlagen, Kopfzeile …).", [...path_("Datei", "Speichern unter"), r(" und als Dateityp „Word-Vorlage (*.dotx)“ wählen.", { size: 24 })], "Doppelklick auf die Vorlage öffnet ein neues Dokument."]),
  img("dotx", 600),
  pb(),

  H1("Jetzt seid ihr dran"),
  P([r("Öffnet die Datei ", { size: 24 }), r("04_Uebungsdatei_Fussball", { bold: true, size: 24 }), r(". Ihr habt 25 Minuten. Arbeitet in Gruppen (3–5 Personen) und hakt ab:", { size: 24 })]),
  check("Den Titel „Fußball – mehr als nur ein Spiel“ mit der Vorlage „Titel“ formatieren."),
  check("Die 8 Hauptkapitel (Einleitung bis Fazit) bekommen „Überschrift 1“."),
  check("Die Unterkapitel (z. B. „Die Anfänge“, „Abseits“) bekommen „Überschrift 2“."),
  check("Die Einleitung in Blocksatz setzen und „beliebteste Sportart der Welt“ fett machen."),
  check("Die 4 Rekorde unter „Zahlen und Rekorde“ als Aufzählung mit Punkten formatieren."),
  check("„Überschrift 1“ ändern: eine neue Farbe."),
  check("In der Gliederung „Fußball in Deutschland“ vor „Die großen Turniere“ schieben."),
  check("Nach dem Titel: Strg + Enter, Inhaltsverzeichnis, wieder Strg + Enter."),
  check("Kopfzeile: euren Namen eintragen. Fußzeile: Seitenzahl."),
  check("„Erste Seite anders“ anhaken."),
  check("Als Word-Vorlage (.dotx) speichern."),
  P([r("Schnell fertig? Bonus: „Zahlen und Rekorde“ auf eine eigene Seite im Querformat setzen.", { italics: true, color: GREY, size: 22 })], { spacing: { before: 120 } }),
  H2("So soll es am Ende aussehen"),
  img("ziel", 600),
  pb(),
  H1("Quiz"),
  ...steps([
    "Ist „zentriert“ ein Zeichenformat oder ein Absatzformat?",
    "Warum braucht ein Inhaltsverzeichnis Formatvorlagen?",
    "Wie öffnet ihr die Kopfzeile am schnellsten?",
    "Wozu ist eine Dokumentvorlage (.dotx) gut?",
  ], 25),
  img("regeln", WPX, { before: 300 }),
])]);

// =========================== 03 HANDOUT ===========================
// Each row: left = question + 1-sentence answer + menu path, right = picture
const IW = 300;                       // picture width in px
const LW = W - IW * 15 - 200;
const hrow = (nr, frage, antwort, wege, pic) => new TableRow({ cantSplit: true, children: [
  cell([
    P([r(nr + "  ", { bold: true, color: ACCENT, size: 26 }), r(frage, { bold: true, color: NAVY, size: 25 })], { spacing: { after: 60 } }),
    P([r(antwort, { size: 22 })], { spacing: { after: 60 } }),
    ...wege.map((w) => P(w, { spacing: { after: 30 } })),
  ], LW, { borders: { ...nb, bottom: thin }, valign: VerticalAlign.CENTER, margins: { top: 140, bottom: 140, left: 60, right: 160 } }),
  new TableCell({ width: { size: W - LW, type: WidthType.DXA }, borders: { ...nb, bottom: thin }, verticalAlign: VerticalAlign.CENTER, margins: { top: 100, bottom: 100, left: 60, right: 0 },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0, line: 240, lineRule: "auto" }, children: [imgRun(pic, IW)] })] }),
] });
const w = (...p) => [r("→ ", { color: ACCENT, bold: true, size: 22 }), ...path_(...p).map((x) => x)];
const handoutTable = (rows) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [LW, W - LW], rows });
await save("03_Handout", [sec("Handout", [
  img("ban_handout", WPX, { before: 0, after: 120 }),
  handoutTable([
    hrow("1", "Wie gestalte ich Text?", "Text markieren und Schrift oder Absatz einstellen.", [w("Start", "Schriftart / Absatz")], "rib_format"),
    hrow("2", "Was ist eine Dokumentvorlage?", "Eine Musterdatei (.dotx): Jedes neue Dokument daraus ist schon fertig eingerichtet.", [], "dokvorlage"),
    hrow("3", "Wie erstelle ich eine eigene Dokumentvorlage?", "Dokument einrichten und als Word-Vorlage speichern.", [w("Datei", "Speichern unter", "Word-Vorlage (*.dotx)")], "dotx"),
    hrow("4", "Welche Formatvorlagen nutze ich?", "Titel, Überschrift 1 und 2, Standard. Ändern: Rechtsklick › Ändern.", [w("Start", "Formatvorlagen")], "rib_vorlagen"),
    hrow("5", "Wie füge ich Kopf-, Fußzeile und Seitenzahl ein?", "Am schnellsten: Doppelklick in den oberen oder unteren Rand.", [w("Einfügen", "Kopfzeile / Fußzeile / Seitenzahl")], "rib_einfuegen"),
  ]),
  pb(),
  handoutTable([
    hrow("6", "Wie mache ich unterschiedliche Seiten?", "Abschnittsumbruch einfügen, dann z. B. Querformat wählen.", [w("Layout", "Umbrüche", "Nächste Seite"), w("Layout", "Ausrichtung")], "abschnitte"),
    hrow("7", "Wie nutze ich die Gliederungsansicht?", "Zeigt nur die Überschriften. Kapitel mit ▲ ▼ verschieben.", [w("Ansicht", "Gliederung")], "gliederung"),
    hrow("8", "Wie erstelle ich das Inhaltsverzeichnis?", "Word baut es automatisch aus den Überschriften.", [w("Referenzen", "Inhaltsverzeichnis")], "toc"),
  ]),
  P([r("Tastenkürzel", { bold: true, color: NAVY, size: 28 })], { spacing: { before: 200, after: 40 } }),
  img("shortcuts", 420, { before: 0, after: 60 }),
  img("regeln", 560, { before: 0, after: 0 }),
])]);

// =========================== 04 ÜBUNGSDATEI ===========================
const plain = (t) => new Paragraph({ children: [new TextRun(t)] });
await save("04_Uebungsdatei_Fussball", [{ properties: { page }, headers: { default: new Header({ children: [new Paragraph({ children: [new TextRun("Übungsdatei – Team „Die Allrounder“")] })] }) },
  children: require("./fussball_text.js").map(plain) }]);

// =========================== 05 FEEDBACK ===========================
const fw = [4406, 1160, 1160, 1160, 1160, 1160];
const fq = ["Die Schulung war gut organisiert.", "Alles wurde verständlich erklärt.", "Das Tempo war passend.", "Die Übungen am PC haben geholfen.", "Ich kann jetzt Formatvorlagen nutzen.", "Ich kann jetzt Kopf- und Fußzeile einfügen."];
const lines = (n) => Array.from({ length: n }, () => new Paragraph({ spacing: { before: 320, after: 0 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA", space: 1 } }, children: [] }));
const C = AlignmentType.CENTER;
await save("05_Feedbackbogen", [sec("Feedback", [
  img("ban_feedback", WPX, { before: 0, after: 240 }),
  P([r("Kreuzt bitte an:", { bold: true, size: 24 })]),
  table(fw, ["", "super", "gut", "okay", "eher nicht", "gar nicht"], fq.map((q) => [q, "☐", "☐", "☐", "☐", "☐"]), { align: [null, C, C, C, C, C] }),
  H2("Das war gut:"), ...lines(2),
  H2("Das geht besser:"), ...lines(2),
  P([r("Danke für euer Feedback!", { bold: true, color: NAVY, size: 26 })], { alignment: AlignmentType.RIGHT, spacing: { before: 320 } }),
])]);

// =========================== 06 BESCHEINIGUNG ===========================
const sig = (n) => new TableCell({ width: { size: 4300, type: WidthType.DXA }, borders: nb, children: [
  new Paragraph({ spacing: { before: 700, after: 0 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "333333", space: 1 } }, children: [] }),
  new Paragraph({ alignment: C, children: [r(n, { size: 20, color: GREY })] })] });
const dbl = { style: BorderStyle.DOUBLE, size: 12, color: NAVY, space: 24 };
await save("06_Teilnahmebescheinigung", [sec(null, [
  img("siegel", 180, { before: 300, after: 200 }),
  P([r("TEILNAHMEBESCHEINIGUNG", { bold: true, color: ACCENT, size: 32 })], { alignment: C, spacing: { after: 400 } }),
  P([r("Hiermit bestätigen wir, dass", { size: 26 })], { alignment: C }),
  new Paragraph({ alignment: C, spacing: { before: 600, after: 60 }, indent: { left: 1500, right: 1500 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 1 } }, children: [] }),
  P([r("Name, Klasse", { size: 18, color: GREY })], { alignment: C, spacing: { after: 400 } }),
  P([r("am 17. November 2026 (09:15 – 10:45 Uhr) an der", { size: 26 })], { alignment: C }),
  P([r("Word-Schulung", { bold: true, color: NAVY, size: 52 })], { alignment: C, spacing: { before: 160, after: 60 } }),
  P([r("Formate · Inhaltsverzeichnis · Kopf- und Fußzeile · Dokumentvorlagen", { color: GREY, size: 24 })], { alignment: C, spacing: { after: 200 } }),
  P([r("erfolgreich teilgenommen hat.", { size: 26 })], { alignment: C, spacing: { after: 400 } }),
  new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [4300, W - 8600, 4300], rows: [new TableRow({ children: [
    sig("Nico – Team „Die Allrounder“"), new TableCell({ width: { size: W - 8600, type: WidthType.DXA }, borders: nb, children: [P("")] }), sig("Yannik – Team „Die Allrounder“")] })] }),
], { footer: simpleFooter(), page: { ...page, borders: { pageBorderTop: dbl, pageBorderBottom: dbl, pageBorderLeft: dbl, pageBorderRight: dbl } } })]);

// =========================== 00 ABLAUFPLAN ===========================
await save("00_Ablaufplan_intern", [sec("Ablaufplan (intern)", [
  img("ban_ablauf", 520, { before: 0, after: 40 }),
  H1("Wer macht was?"),
  table([1900, 2100, 4706, 1500], ["Zeit", "Teil", "Was passiert", "Wer"], [
    ["09:15 (10 min)", "Start", "Begrüßen, Ablauf zeigen, Fragen sammeln", "Yannik"],
    ["09:25 (15 min)", "Formate und Vorlagen", "Schrift, Absatz, Pinsel, Formatvorlagen anwenden und ändern", "Nico"],
    ["09:40 (15 min)", "Inhaltsverzeichnis", "Gliederung zeigen, Inhaltsverzeichnis einfügen und aktualisieren", "Yannik"],
    ["09:55 (10 min)", "Seiten, Kopf und Fuß", "Kopf-/Fußzeile, Seitenzahl, Abschnitt mit Querformat", "Nico"],
    ["10:05 (5 min)", "Eigene Vorlage", "Als .dotx speichern und neu öffnen", "Yannik"],
    ["10:10 (25 min)", "Üben", "Fußball-Text in Gruppen (3–5), wir helfen", "Beide"],
    ["10:35 (10 min)", "Quiz und Ende", "Quiz, Handout, Feedback, Bescheinigung", "Beide"],
  ], { boldFirst: true }),
  H1("Checkliste"),
  check("Einladung aufgehängt, Anmeldungen gezählt"),
  check("Raum 0.68: PCs, Beamer und Word getestet, gelüftet"),
  check("Übungsdatei auf allen PCs"),
  check("Gedruckt: Schulungsblatt, Handout, Feedback, Bescheinigungen"),
  check("Einmal komplett geprobt (mit Uhr)"),
  box([[r("Tipps: ", { bold: true }), r("frei sprechen · langsam klicken und ansagen · Fragen stellen · Handout erst am Ende austeilen")]], BLUEBG, NAVY),
  H1("Quiz-Lösungen"),
  ...steps(["Absatzformat – gilt für den ganzen Absatz.", "Word erkennt Überschriften nur an der Formatvorlage.", "Doppelklick in den oberen Seitenrand.", "Jedes neue Dokument ist sofort fertig eingerichtet."], 23),
])]);
})();
