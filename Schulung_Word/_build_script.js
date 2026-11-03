const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, BorderStyle, AlignmentType, HeadingLevel, LevelFormat, Header, Footer,
  PageNumber, PageBreak, TableOfContents, VerticalAlign,
} = require("docx");

const OUT = "/home/user/Owner/Schulung_Word";
const NAVY = "1F3A5F", ACCENT = "E07A1F", LIGHT = "EAF0F7", LIGHT2 = "FDF1E6", GREY = "666666", LINE = "C9D3E0";
const FONT = "Calibri";
const W = 9638; // A4 text width with 2 cm margins (11906 - 2*1134)

// ---------- helpers ----------
const r = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
const P = (children, o = {}) =>
  new Paragraph({ children: typeof children === "string" ? [r(children)] : children, ...o });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [r(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [r(t)] });
const bullet = (children, level = 0) =>
  new Paragraph({ numbering: { reference: "bul", level }, spacing: { after: 60 },
    children: typeof children === "string" ? [r(children)] : children });
const num = (children, ref = "num") =>
  new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80 },
    children: typeof children === "string" ? [r(children)] : children });
const check = (children) =>
  new Paragraph({ numbering: { reference: "chk", level: 0 }, spacing: { after: 80 },
    children: typeof children === "string" ? [r(children)] : children });
// "Button" look for menu names
const key = (t) => r(` ${t} `, { bold: true, color: NAVY, shading: { type: ShadingType.CLEAR, fill: LIGHT, color: "auto" } });
const sep = () => r("  ›  ", { color: ACCENT, bold: true });
const clickPath = (...items) => {
  const runs = [r("Klickpfad:  ", { color: GREY, size: 20 })];
  items.forEach((it, i) => { if (i) runs.push(sep()); runs.push(key(it)); });
  return new Paragraph({ children: runs, spacing: { before: 60, after: 120 } });
};
const kbd = (t) => r(` ${t} `, { bold: true, font: "Consolas", size: 20, shading: { type: ShadingType.CLEAR, fill: "EFEFEF", color: "auto" } });
const rule = () => new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 4 } }, spacing: { after: 200 } });
const spacer = (after = 120) => new Paragraph({ children: [], spacing: { after } });
const pb = () => new Paragraph({ children: [new PageBreak()] });

const noBorders = { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } };
const thin = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const thinBorders = { top: thin, bottom: thin, left: thin, right: thin };

function cell(content, width, o = {}) {
  if (typeof content === "string" && content.includes("\n")) content = content.split("\n");
  const paras = (Array.isArray(content) && !(content[0] instanceof TextRun) ? content : [content]).map((c) =>
    c instanceof Paragraph ? c : P(typeof c === "string" ? [r(c, o.run || {})] : c, { spacing: { after: 40 } }));
  return new TableCell({
    children: paras, width: { size: width, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined,
    borders: o.borders || thinBorders, verticalAlign: o.valign || VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 140, right: 140 }, columnSpan: o.span,
  });
}
function table(widths, header, rows, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const trs = [];
  if (header) trs.push(new TableRow({ tableHeader: true,
    children: header.map((h, i) => cell(h, widths[i], { fill: NAVY, run: { bold: true, color: "FFFFFF" } })) }));
  rows.forEach((row, ri) => trs.push(new TableRow({ cantSplit: true,
    children: row.map((c, i) => cell(c, widths[i], { fill: o.zebra && ri % 2 ? "F6F8FB" : (o.firstColFill && i === 0 ? LIGHT : undefined), run: o.firstColBold && i === 0 ? { bold: true, color: NAVY } : {} })) })));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: trs });
}
// Coloured info box (1-cell table with thick left border)
function box(title, paras, fill = LIGHT, edge = NAVY) {
  const kids = [];
  if (title) kids.push(P([r(title, { bold: true, color: edge, size: 23 })], { spacing: { after: 80 } }));
  paras.forEach((p) => kids.push(p instanceof Paragraph ? p : P(p, { spacing: { after: 60 } })));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [
    new TableCell({ children: kids, width: { size: W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill, color: "auto" },
      borders: { ...noBorders, left: { style: BorderStyle.SINGLE, size: 36, color: edge } },
      margins: { top: 140, bottom: 140, left: 220, right: 220 } })] })] });
}

const styles = {
  default: { document: { run: { font: FONT, size: 22 }, paragraph: { spacing: { after: 120, line: 276 } } } },
  paragraphStyles: [
    { id: "Title", name: "Title", basedOn: "Normal", next: "Normal", run: { font: FONT, size: 52, bold: true, color: NAVY }, paragraph: { spacing: { before: 0, after: 80 }, keepNext: true } },
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 32, bold: true, color: NAVY },
      paragraph: { spacing: { before: 320, after: 140 }, keepNext: true, outlineLevel: 0, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 3 } } } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 25, bold: true, color: ACCENT },
      paragraph: { spacing: { before: 220, after: 100 }, keepNext: true, outlineLevel: 1 } },
  ],
};
const numbering = { config: [
  { reference: "bul", levels: [
    { level: 0, format: LevelFormat.BULLET, text: "■", alignment: AlignmentType.LEFT, style: { run: { color: ACCENT, size: 16 }, paragraph: { indent: { left: 500, hanging: 300 } } } },
    { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 900, hanging: 300 } } } } ] },
  { reference: "chk", levels: [{ level: 0, format: LevelFormat.BULLET, text: "☐", alignment: AlignmentType.LEFT, style: { run: { size: 26, color: NAVY }, paragraph: { indent: { left: 500, hanging: 360 } } } }] },
  ...["num", "num2", "num3", "num4", "num5", "num6", "num7", "num8", "num9"].map((ref) => ({ reference: ref, levels: [
    { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { run: { bold: true, color: NAVY }, paragraph: { indent: { left: 500, hanging: 360 } } } }] })),
] };

const header = (label) => new Header({ children: [new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } },
  tabStops: [{ type: "right", position: W }],
  children: [r("Die Allrounder", { bold: true, color: NAVY, size: 18 }), r("  |  Word-Schulung", { color: GREY, size: 18 }), r(`\t${label}`, { color: GREY, size: 18 })] })] });
const footer = () => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
  r("Seite ", { size: 18, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], size: 18, color: GREY, font: FONT }),
  r(" von ", { size: 18, color: GREY }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: GREY, font: FONT }),
  r("   ·   Nico & Yannik", { size: 18, color: GREY })] })] });
const page = { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134, header: 567, footer: 567 } };

async function save(name, sections, extra = {}) {
  const doc = new Document({ creator: "Nico & Yannik – Die Allrounder", title: name, styles, numbering, features: extra.features, sections });
  fs.writeFileSync(path.join(OUT, name + ".docx"), await Packer.toBuffer(doc));
  console.log("wrote", name);
}

// Training data (please verify!)
const DATUM = "Dienstag, 17. November 2026";
const ZEIT = "08:30 – 10:00 Uhr";
const RAUM = "Raum 0.68";

const titleBlock = (kicker, title, sub) => [
  P([r(kicker.toUpperCase(), { bold: true, color: ACCENT, size: 20 })], { spacing: { before: 0, after: 100, line: 240 } }),
  P([r(title, { bold: true, size: 48, color: NAVY })], { spacing: { before: 160, after: 60, line: 240 } }),
  sub ? P([r(sub, { size: 26, color: GREY })], { spacing: { after: 60 } }) : null,
  rule(),
].filter(Boolean);

(async () => {
// =====================================================================
// 01 EINLADUNG
// =====================================================================
await save("01_Einladung", [{ properties: { page }, footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
    children: [r("Team „Die Allrounder“  ·  Nico & Yannik", { size: 18, color: GREY })] })] }) },
  children: [
    P([r("E I N L A D U N G", { bold: true, color: ACCENT, size: 22 })], { alignment: AlignmentType.CENTER, spacing: { before: 400, after: 160, line: 240 } }),
    P([r("Word-Schulung", { bold: true, color: NAVY, size: 64 })], { alignment: AlignmentType.CENTER, spacing: { before: 120, after: 40, line: 900, lineRule: "atLeast" } }),
    P([r("Formate  ·  Formatvorlagen  ·  Kopf- und Fußzeile", { size: 28, color: GREY })], { alignment: AlignmentType.CENTER, spacing: { after: 120 } }),
    rule(),
    P([r("Liebe Schülerinnen und Schüler,", { bold: true })], { spacing: { before: 200 } }),
    P("ob Referat, Bewerbung oder Facharbeit – wer seine Dokumente in Word sauber gestaltet, spart Zeit und macht sofort einen professionellen Eindruck. Deshalb laden wir euch herzlich zu unserer Schulung ein."),
    P("In kurzer Zeit zeigen wir euch die wichtigsten Werkzeuge und ihr probiert alles direkt selbst am Computer aus. Vorkenntnisse braucht ihr keine.", { spacing: { after: 240 } }),
    table([2600, W - 2600], null, [
      ["Datum", DATUM],
      ["Uhrzeit", `${ZEIT} (ca. 90 Minuten)`],
      ["Ort", RAUM],
      ["Thema", "Formate, Formatvorlagen sowie Kopf- und Fußzeile in Microsoft Word"],
      ["Schulungsleitung", "Nico und Yannik (Team „Die Allrounder“)"],
      ["Mitbringen", "Nur euch selbst und gute Laune – Computer und Material stellen wir."],
    ], { firstColFill: true, firstColBold: true }),
    spacer(200),
    H2("Das lernt ihr bei uns"),
    bullet([r("Formate: ", { bold: true }), r("Schrift, Farben und Absätze gezielt gestalten")]),
    bullet([r("Formatvorlagen: ", { bold: true }), r("Überschriften mit einem Klick einheitlich formatieren – inklusive automatischem Inhaltsverzeichnis")]),
    bullet([r("Kopf- und Fußzeile: ", { bold: true }), r("Name, Datum und Seitenzahlen auf jeder Seite einfügen")]),
    spacer(120),
    box("Bitte gebt uns kurz Bescheid", [[r("Damit wir genug Plätze und Material vorbereiten können, meldet euch bitte bis "), r("Freitag, 13. November 2026", { bold: true }), r(" bei Nico oder Yannik an.")]], LIGHT2, ACCENT),
    spacer(240),
    P("Wir freuen uns auf euch!"),
    P([r("Nico & Yannik", { bold: true, color: NAVY, size: 26 })], { spacing: { before: 200, after: 0 } }),
    P([r("Team „Die Allrounder“", { color: GREY })]),
  ] }]);

// =====================================================================
// 02 SCHULUNGSBLATT (visual guide, projected + used in exercise phase)
// =====================================================================
const sb = [
  ...titleBlock("Schulungsunterlage", "Word wie die Profis", "Formate · Formatvorlagen · Kopf- und Fußzeile"),
  box("Eure Lernziele – nach der Schulung könnt ihr …", [
    bullet("Text mit Zeichen- und Absatzformaten gezielt gestalten,"),
    bullet("Formatvorlagen anwenden, ändern und daraus ein Inhaltsverzeichnis erstellen,"),
    bullet("Kopf- und Fußzeilen mit Name, Datum und Seitenzahl einfügen."),
  ]),
  spacer(),
  H2("Ablauf"),
  table([1500, 5000, 3138], ["Zeit", "Programmpunkt", "Form"], [
    ["08:30", "Begrüßung, Ablauf & eure Fragen", "Gespräch"],
    ["08:40", "Teil 1: Formate", "Vorführung"],
    ["08:55", "Teil 2: Formatvorlagen", "Vorführung"],
    ["09:10", "Teil 3: Kopf- und Fußzeile", "Vorführung"],
    ["09:20", "Praxis: Übungsaufgaben am PC", "Gruppen (3–5 Pers.)"],
    ["09:45", "Quiz & offene Fragen", "Gemeinsam"],
    ["09:52", "Handout, Feedback, Bescheinigung", "Abschluss"],
  ], { zebra: true }),

  pb(),
  H1("Teil 1 – Formate"),
  P("Ein Format bestimmt, wie Text aussieht. Word unterscheidet zwei Arten:"),
  table([2200, 3719, 3719], ["", "Zeichenformat", "Absatzformat"], [
    ["Gilt für", "markierte Zeichen / Wörter", "den ganzen Absatz (bis zur Enter-Taste)"],
    ["Beispiele", "Schriftart, Schriftgröße, Fett, Kursiv, Farbe", "Ausrichtung, Zeilenabstand, Einzug, Abstand vor/nach"],
    ["Wo?", "Start › Gruppe „Schriftart“", "Start › Gruppe „Absatz“"],
  ], { firstColFill: true, firstColBold: true }),
  H2("So geht’s"),
  num("Text markieren (mit der Maus ziehen oder Doppelklick auf ein Wort).", "num"),
  num("Im Register Start die gewünschte Formatierung anklicken.", "num"),
  clickPath("Start", "Schriftart / Absatz"),
  H2("Wichtige Tastenkürzel"),
  table([3200, 1619, 3200, 1619], null, [
    [[kbd("Strg"), r(" + "), kbd("Umschalt"), r(" + "), kbd("F")], "Fett", [kbd("Strg"), r(" + "), kbd("E")], "Zentriert"],
    [[kbd("Strg"), r(" + "), kbd("Umschalt"), r(" + "), kbd("K")], "Kursiv", [kbd("Strg"), r(" + "), kbd("L")], "Linksbündig"],
    [[kbd("Strg"), r(" + "), kbd("Umschalt"), r(" + "), kbd("U")], "Unterstrichen", [kbd("Strg"), r(" + "), kbd("B")], "Blocksatz"],
    [[kbd("Strg"), r(" + "), kbd("Leertaste")], "Format entfernen", [kbd("Strg"), r(" + "), kbd("Z")], "Rückgängig"],
  ]),
  spacer(),
  box("Profi-Trick: Format übertragen (Pinsel)", [
    [r("Text mit dem gewünschten Format anklicken › "), key("Start"), sep(), key("Format übertragen"), r(" › anderen Text überstreichen. "), r("Doppelklick", { bold: true }), r(" auf den Pinsel = mehrmals übertragen, mit "), kbd("Esc"), r(" beenden.")],
  ], LIGHT2, ACCENT),

  H1("Teil 2 – Formatvorlagen"),
  P("Eine Formatvorlage ist ein gespeichertes Paket aus vielen Formaten (z. B. Schrift, Größe, Farbe, Abstand) mit einem Namen wie „Überschrift 1“."),
  table([4819, 4819], ["Ohne Formatvorlage", "Mit Formatvorlage"], [
    ["Jede Überschrift einzeln formatieren", "Ein Klick – fertig"],
    ["Änderung = alle Überschriften neu bearbeiten", "Vorlage einmal ändern – alle passen sich an"],
    ["Uneinheitliches Aussehen", "Immer einheitlich und professionell"],
    ["Kein automatisches Inhaltsverzeichnis", "Inhaltsverzeichnis per Knopfdruck"],
  ]),
  H2("A) Formatvorlage anwenden"),
  num("Cursor in die Überschrift setzen.", "num2"),
  num("Im Katalog die passende Vorlage anklicken, z. B. „Überschrift 1“.", "num2"),
  clickPath("Start", "Formatvorlagen", "Überschrift 1"),
  H2("B) Formatvorlage ändern"),
  num("Rechtsklick auf die Vorlage im Katalog › „Ändern …“.", "num3"),
  num("Schrift, Größe, Farbe usw. einstellen › OK. Alle Überschriften ändern sich automatisch.", "num3"),
  clickPath("Start", "Formatvorlagen", "Rechtsklick", "Ändern …"),
  H2("C) Inhaltsverzeichnis erstellen"),
  num("Cursor an den Anfang des Dokuments setzen.", "num4"),
  num("Automatisches Inhaltsverzeichnis wählen. Word nimmt alle Überschriften-Vorlagen auf.", "num4"),
  num("Nach Änderungen: Rechtsklick ins Verzeichnis › „Felder aktualisieren“.", "num4"),
  clickPath("Referenzen", "Inhaltsverzeichnis", "Automatische Tabelle 1"),
  box("Tipp: Navigationsbereich", [[r("Unter "), key("Ansicht"), sep(), key("Navigationsbereich"), r(" seht ihr links alle Überschriften und könnt per Klick durchs Dokument springen – funktioniert nur mit Formatvorlagen!")]]),

  H1("Teil 3 – Kopf- und Fußzeile"),
  P("Kopf- und Fußzeile stehen im oberen bzw. unteren Seitenrand und erscheinen automatisch auf jeder Seite – ideal für Name, Klasse, Thema, Datum und Seitenzahl."),
  H2("A) Kopf- oder Fußzeile einfügen"),
  num([r("Doppelklick in den oberen (bzw. unteren) Seitenrand "), r("oder", { bold: true }), r(" über das Menü:")], "num5"),
  clickPath("Einfügen", "Kopfzeile / Fußzeile", "Leer"),
  num("Text eingeben, z. B. Name, Klasse und Thema.", "num5"),
  num("Zurück zum Text: Doppelklick in den Textbereich oder „Kopf- und Fußzeile schließen“.", "num5"),
  H2("B) Seitenzahl einfügen"),
  clickPath("Einfügen", "Seitenzahl", "Seitenende", "Einfache Zahl 2"),
  H2("C) Datum einfügen"),
  clickPath("Kopf- und Fußzeile", "Datum und Uhrzeit"),
  H2("D) Erste Seite anders (z. B. Deckblatt ohne Seitenzahl)"),
  clickPath("Kopf- und Fußzeile", "Optionen", "☑ Erste Seite anders"),
  box("Gut zu wissen", [
    bullet("Kopf- und Fußzeile werden grau angezeigt, solange ihr im normalen Text arbeitet – das ist richtig so."),
    bullet("Seitenzahlen niemals von Hand tippen! Nur die automatische Seitenzahl passt sich an."),
  ], LIGHT2, ACCENT),

  pb(),
  H1("Praxis – Übungsaufgaben"),
  P([r("Öffnet die Datei "), r("„04_Uebungsdatei.docx“", { bold: true }), r(" und bearbeitet die Aufgaben in eurer Gruppe. Hakt ab, was ihr geschafft habt.")]),
  H2("Formate"),
  check("Formatiert den Titel „Bienen – kleine Tiere, große Wirkung“: zentriert, Schriftgröße 20, fett."),
  check("Setzt den Absatz unter „Einleitung“ in Blocksatz und den Zeilenabstand auf 1,5."),
  check("Hebt im Abschnitt „Bedrohungen“ das Wort „Pestizide“ kursiv und rot hervor. Übertragt dieses Format mit dem Pinsel auf „Monokulturen“."),
  H2("Formatvorlagen"),
  check("Weist dem Titel die Formatvorlage „Titel“ zu."),
  check("Weist allen Hauptüberschriften (Einleitung, Das Leben im Bienenvolk, Warum Bienen so wichtig sind, Bedrohungen, Was wir tun können, Fazit) die Vorlage „Überschrift 1“ zu."),
  check("Weist den Unterüberschriften (Die Königin, Die Arbeiterinnen, Die Drohnen) die Vorlage „Überschrift 2“ zu."),
  check("Ändert die Vorlage „Überschrift 1“: Schriftfarbe dunkelgrün, Größe 16. Beobachtet, was passiert!"),
  check("Öffnet den Navigationsbereich und springt zu „Fazit“."),
  check([r("Fügt nach dem Titel einen Seitenumbruch ein ("), kbd("Strg"), r(" + "), kbd("Enter"), r(") und erstellt auf Seite 2 ein automatisches Inhaltsverzeichnis.")]),
  H2("Kopf- und Fußzeile"),
  check("Kopfzeile: eure Namen links, eure Klasse rechts."),
  check("Fußzeile: Seitenzahl in der Mitte und das heutige Datum."),
  check("Aktiviert „Erste Seite anders“, damit auf dem Deckblatt keine Kopf- und Fußzeile erscheint."),
  spacer(),
  box("Schnell fertig? Bonusaufgabe", [
    P("Erstellt eine eigene Formatvorlage „Merksatz“ (z. B. kursiv, blau, mit Rahmen): Text formatieren › markieren › im Formatvorlagen-Katalog auf den Pfeil › „Formatvorlage erstellen“. Wendet sie auf zwei Sätze an."),
  ], LIGHT2, ACCENT),

  H1("Quiz – Habt ihr’s verstanden?"),
  num("Was ist der Unterschied zwischen einem Zeichen- und einem Absatzformat?", "num6"),
  num("Nennt zwei Vorteile von Formatvorlagen.", "num6"),
  num("Warum funktioniert ein automatisches Inhaltsverzeichnis nur mit Formatvorlagen?", "num6"),
  num("Wie kommt ihr in die Kopfzeile – und wie wieder heraus?", "num6"),
  num("Warum sollte man Seitenzahlen nie von Hand eintippen?", "num6"),
];
await save("02_Schulungsblatt", [{ properties: { page }, headers: { default: header("Schulungsunterlage") }, footers: { default: footer() }, children: sb }]);

// =====================================================================
// 03 HANDOUT (only key points, max 2 pages)
// =====================================================================
const ho = [
  ...titleBlock("Handout", "Word – das Wichtigste auf einen Blick", null),
  new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 160, after: 80 }, children: [r("1  Formate")] }),
  table([2400, 7238], null, [
    ["Zeichenformat", "gilt für markierte Zeichen: Schriftart, Größe, Fett, Kursiv, Farbe › Start › Schriftart"],
    ["Absatzformat", "gilt für den ganzen Absatz: Ausrichtung, Zeilenabstand, Einzug › Start › Absatz"],
    ["Format übertragen", "Pinsel-Symbol: Format von A auf B kopieren (Doppelklick = mehrfach, Esc = Ende)"],
  ], { firstColFill: true, firstColBold: true }),
  new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 160, after: 80 }, children: [r("2  Formatvorlagen")] }),
  P("Gespeichertes Formatpaket mit Namen (z. B. „Überschrift 1“) – für einheitliche, schnell änderbare Dokumente."),
  table([2400, 7238], null, [
    ["Anwenden", [key("Start"), sep(), key("Formatvorlagen"), r(" › Vorlage anklicken")]],
    ["Ändern", [r("Rechtsklick auf Vorlage › "), key("Ändern …"), r(" › alle Stellen passen sich an")]],
    ["Inhaltsverzeichnis", [key("Referenzen"), sep(), key("Inhaltsverzeichnis"), r(" (aktualisieren: Rechtsklick › Felder aktualisieren)")]],
    ["Navigation", [key("Ansicht"), sep(), key("Navigationsbereich")]],
  ], { firstColFill: true, firstColBold: true }),
  new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 160, after: 80 }, children: [r("3  Kopf- und Fußzeile")] }),
  table([2400, 7238], null, [
    ["Öffnen", [r("Doppelklick in den oberen/unteren Rand oder "), key("Einfügen"), sep(), key("Kopfzeile / Fußzeile")]],
    ["Seitenzahl", [key("Einfügen"), sep(), key("Seitenzahl"), sep(), key("Seitenende")]],
    ["Datum", [key("Kopf- und Fußzeile"), sep(), key("Datum und Uhrzeit")]],
    ["Deckblatt ohne", [key("Kopf- und Fußzeile"), r(" › ☑ "), key("Erste Seite anders")]],
    ["Schließen", "Doppelklick in den Text oder „Kopf- und Fußzeile schließen“"],
  ], { firstColFill: true, firstColBold: true }),
  new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 160, after: 80 }, children: [r("4  Die wichtigsten Tastenkürzel")] }),
  table([2900, 1919, 2900, 1919], null, [
    [[kbd("Strg"), r("+"), kbd("Umschalt"), r("+"), kbd("F")], "Fett", [kbd("Strg"), r("+"), kbd("E")], "Zentriert"],
    [[kbd("Strg"), r("+"), kbd("Umschalt"), r("+"), kbd("K")], "Kursiv", [kbd("Strg"), r("+"), kbd("B")], "Blocksatz"],
    [[kbd("Strg"), r("+"), kbd("Leertaste")], "Format entfernen", [kbd("Strg"), r("+"), kbd("Enter")], "Seitenumbruch"],
    [[kbd("Strg"), r("+"), kbd("Alt"), r("+"), kbd("1")], "Überschrift 1", [kbd("Strg"), r("+"), kbd("Z")], "Rückgängig"],
  ]),
  box("Die 4 goldenen Regeln", [
    num("Überschriften immer mit Formatvorlagen – nie von Hand formatieren.", "num7"),
    num("Seitenzahlen immer automatisch einfügen – nie eintippen.", "num7"),
    num("Neue Seite mit Strg + Enter – nicht mit vielen Enter-Tasten.", "num7"),
    num("Lieber wenige Schriften und Farben – weniger ist mehr.", "num7"),
  ], LIGHT2, ACCENT),
];
await save("03_Handout", [{ properties: { page: { ...page, margin: { ...page.margin, top: 800, bottom: 700 } } }, headers: { default: header("Handout") }, footers: { default: footer() }, children: ho }]);

// =====================================================================
// 04 UEBUNGSDATEI (deliberately unformatted)
// =====================================================================
const plain = (t) => new Paragraph({ children: [new TextRun(t)] });
const ue = [
  "Bienen – kleine Tiere, große Wirkung",
  "Einleitung",
  "Bienen gehören zu den wichtigsten Tieren unseres Planeten. Ohne sie gäbe es viele Obst- und Gemüsesorten nicht, die wir jeden Tag essen. In diesem Text erfährst du, wie ein Bienenvolk aufgebaut ist, warum Bienen so wichtig sind und was jeder von uns tun kann, um sie zu schützen.",
  "Das Leben im Bienenvolk",
  "Ein Bienenvolk besteht im Sommer aus bis zu 50.000 Tieren. Jede Biene hat eine feste Aufgabe. Man unterscheidet drei Gruppen: die Königin, die Arbeiterinnen und die Drohnen.",
  "Die Königin",
  "Die Königin ist die einzige Biene im Volk, die Eier legt. An einem Tag kann sie bis zu 2.000 Eier legen. Sie wird bis zu fünf Jahre alt und ist deutlich größer als die anderen Bienen.",
  "Die Arbeiterinnen",
  "Die Arbeiterinnen erledigen fast alle Aufgaben im Stock. Sie putzen die Waben, füttern die Larven, bauen neue Waben aus Wachs, bewachen den Eingang und sammeln Nektar und Pollen. Im Sommer leben sie nur etwa sechs Wochen.",
  "Die Drohnen",
  "Drohnen sind die männlichen Bienen. Sie haben keinen Stachel und sammeln keinen Nektar. Ihre einzige Aufgabe ist es, eine junge Königin zu begatten. Im Herbst werden sie aus dem Stock vertrieben.",
  "Warum Bienen so wichtig sind",
  "Bienen bestäuben beim Sammeln von Nektar die Blüten von Pflanzen. Dabei tragen sie Pollen von einer Blüte zur nächsten. Erst dadurch können viele Pflanzen Früchte und Samen bilden. Etwa ein Drittel unserer Nahrungsmittel hängt direkt oder indirekt von der Bestäubung durch Insekten ab. Dazu gehören zum Beispiel Äpfel, Kirschen, Erdbeeren, Gurken und Kürbisse.",
  "Neben der Bestäubung liefern Honigbienen auch Honig und Wachs. Ein Bienenvolk kann in einem guten Jahr über 30 Kilogramm Honig produzieren. Für ein einziges Glas Honig fliegen die Bienen insgesamt eine Strecke, die etwa dreimal um die Erde reicht.",
  "Bedrohungen",
  "Leider geht es vielen Bienenarten schlecht. In Deutschland gibt es über 550 Wildbienenarten, und mehr als die Hälfte davon gilt als gefährdet. Gründe dafür sind unter anderem Pestizide, die in der Landwirtschaft eingesetzt werden, und Monokulturen, also riesige Felder, auf denen nur eine einzige Pflanzenart wächst. Dort finden Bienen nach der kurzen Blütezeit keine Nahrung mehr.",
  "Auch Krankheiten und Parasiten wie die Varroamilbe schwächen die Bienenvölker. Zusätzlich verschwinden durch Bebauung und aufgeräumte Gärten immer mehr Nistplätze.",
  "Was wir tun können",
  "Jeder kann etwas für Bienen tun. Wer einen Garten oder Balkon hat, kann heimische Blumen und Kräuter pflanzen, zum Beispiel Lavendel, Thymian oder Sonnenblumen. Ein Insektenhotel bietet Wildbienen einen Platz zum Nisten. Außerdem hilft es, auf Pflanzenschutzmittel zu verzichten und beim Einkauf regionale Produkte zu wählen.",
  "Auch an Schulen gibt es viele Möglichkeiten: Eine Blühwiese auf dem Schulhof oder eine Schul-AG zum Thema Imkerei sind tolle Projekte, bei denen alle etwas lernen können.",
  "Fazit",
  "Bienen sind klein, aber für Natur und Menschen unverzichtbar. Wenn wir sie schützen, schützen wir gleichzeitig unsere Ernährung und die Vielfalt der Natur. Schon kleine Schritte können dabei einen großen Unterschied machen.",
].map(plain);
await save("04_Uebungsdatei", [{ properties: { page: { size: page.size, margin: { top: 1418, bottom: 1134, left: 1418, right: 1134 } } }, children: ue }],);

// =====================================================================
// 05 FEEDBACKBOGEN + TEILNAHMEBESCHEINIGUNG
// =====================================================================
const scale = (q) => [q, "☐", "☐", "☐", "☐", "☐"];
const fbW = [5138, 900, 900, 900, 900, 900];
const centered = (t) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [r(t, { size: 26, color: NAVY })] });
const fbRow = (q, ri) => new TableRow({ cantSplit: true, children: scale(q).map((c, i) => i === 0 ? cell(c, fbW[0], { fill: ri % 2 ? "F6F8FB" : undefined })
  : new TableCell({ children: [centered(c)], width: { size: fbW[i], type: WidthType.DXA }, borders: thinBorders, verticalAlign: VerticalAlign.CENTER,
      shading: ri % 2 ? { type: ShadingType.CLEAR, fill: "F6F8FB", color: "auto" } : undefined, margins: { top: 90, bottom: 90, left: 60, right: 60 } })) });
const fbHead = new TableRow({ tableHeader: true, children: ["Aussage", "trifft voll zu", "trifft eher zu", "teils / teils", "trifft eher nicht zu", "trifft nicht zu"].map((h, i) =>
  new TableCell({ children: [new Paragraph({ alignment: i ? AlignmentType.CENTER : AlignmentType.LEFT, spacing: { after: 0 }, children: [r(h, { bold: true, color: "FFFFFF", size: i ? 16 : 22 })] })],
    width: { size: fbW[i], type: WidthType.DXA }, borders: thinBorders, verticalAlign: VerticalAlign.CENTER, shading: { type: ShadingType.CLEAR, fill: NAVY, color: "auto" },
    margins: { top: 80, bottom: 80, left: 60, right: 60 } })) });
const questions = [
  "Die Schulung war gut organisiert (Einladung, Raum, Technik).",
  "Die Inhalte wurden verständlich erklärt.",
  "Das Tempo war für mich passend.",
  "Die Übungen am PC haben mir beim Verstehen geholfen.",
  "Ich kann jetzt Formatvorlagen anwenden und ändern.",
  "Ich kann jetzt Kopf- und Fußzeile mit Seitenzahl einfügen.",
  "Das Handout ist hilfreich zum Nachschlagen.",
  "Ich würde die Schulung weiterempfehlen.",
];
const lines = (n) => Array.from({ length: n }, () => new Paragraph({ spacing: { before: 240, after: 0 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "999999", space: 1 } }, children: [] }));
const feedback = [
  ...titleBlock("Feedbackbogen", "Wie fandet ihr die Schulung?", "Euer Feedback ist anonym und hilft uns, besser zu werden."),
  new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: fbW, rows: [fbHead, ...questions.map(fbRow)] }),
  H2("Das hat mir besonders gut gefallen:"), ...lines(2),
  H2("Das könnt ihr verbessern:"), ...lines(2),
  H2("Diese Frage ist für mich noch offen:"), ...lines(2),
  P([r("Gesamtnote für die Schulung:   ", { bold: true }), r("1 ☐    2 ☐    3 ☐    4 ☐    5 ☐    6 ☐", { size: 24 })], { spacing: { before: 300 } }),
  P([r("Vielen Dank für euer Feedback!", { italics: true, color: GREY })], { alignment: AlignmentType.RIGHT, spacing: { before: 200 } }),
];
const sigCell = (name) => new TableCell({ width: { size: 4319, type: WidthType.DXA }, borders: noBorders, children: [
  new Paragraph({ spacing: { before: 700, after: 0 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "333333", space: 1 } }, children: [] }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [r(name, { size: 20, color: GREY })] })] });
const bescheinigung = [
  P([r("TEILNAHMEBESCHEINIGUNG", { bold: true, color: ACCENT, size: 24, characterSpacing: 120 })], { alignment: AlignmentType.CENTER, spacing: { before: 1200, after: 600 } }),
  P([r("Hiermit bestätigen wir, dass", { size: 26 })], { alignment: AlignmentType.CENTER }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 600, after: 60 }, indent: { left: 1400, right: 1400 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 1 } }, children: [] }),
  P([r("Vor- und Nachname, Klasse", { size: 18, color: GREY })], { alignment: AlignmentType.CENTER, spacing: { after: 500 } }),
  P([r("am " + DATUM.replace(/^\w+, /, "") + " erfolgreich an der Schulung", { size: 26 })], { alignment: AlignmentType.CENTER }),
  P([r("„Word-Schulung: Formate, Formatvorlagen", { bold: true, color: NAVY, size: 36 })], { alignment: AlignmentType.CENTER, spacing: { before: 200, after: 0 } }),
  P([r("sowie Kopf- und Fußzeile“", { bold: true, color: NAVY, size: 36 })], { alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
  P([r("teilgenommen hat (Dauer: 90 Minuten).", { size: 26 })], { alignment: AlignmentType.CENTER, spacing: { after: 500 } }),
  box("Inhalte der Schulung", [
    bullet("Zeichen- und Absatzformate gezielt einsetzen"),
    bullet("Formatvorlagen anwenden, ändern und automatisches Inhaltsverzeichnis erstellen"),
    bullet("Kopf- und Fußzeilen mit Seitenzahl und Datum gestalten"),
  ]),
  P([r(RAUM + ", " + DATUM.replace(/^\w+, /, ""), { color: GREY })], { alignment: AlignmentType.CENTER, spacing: { before: 500 } }),
  new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [4319, 1000, 4319], rows: [new TableRow({ children: [
    sigCell("Nico – Schulungsleitung"), new TableCell({ width: { size: 1000, type: WidthType.DXA }, borders: noBorders, children: [P("")] }), sigCell("Yannik – Schulungsleitung")] })] }),
];
await save("05_Feedbackbogen", [{ properties: { page }, headers: { default: header("Feedbackbogen") }, footers: { default: footer() }, children: feedback }]);
await save("06_Teilnahmebescheinigung", [{ properties: { page: { ...page, borders: {
    pageBorderTop: { style: BorderStyle.DOUBLE, size: 12, color: NAVY, space: 24 }, pageBorderBottom: { style: BorderStyle.DOUBLE, size: 12, color: NAVY, space: 24 },
    pageBorderLeft: { style: BorderStyle.DOUBLE, size: 12, color: NAVY, space: 24 }, pageBorderRight: { style: BorderStyle.DOUBLE, size: 12, color: NAVY, space: 24 } } } },
  footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [r("Team „Die Allrounder“  ·  Nico & Yannik", { size: 18, color: GREY })] })] }) },
  children: bescheinigung }]);

// =====================================================================
// 00 ABLAUFPLAN (internal, for Nico & Yannik)
// =====================================================================
const ap = [
  ...titleBlock("Intern – nur für Nico & Yannik", "Ablauf- und Moderationsplan", "Word-Schulung für die Mittelstufe"),
  table([2600, W - 2600], null, [
    ["Termin", `${DATUM}, ${ZEIT}`],
    ["Ort", RAUM + " (geschlossener, gelüfteter Raum, PCs mit Word, Beamer)"],
    ["Zielgruppe", "Interessierte Schülerinnen und Schüler der Mittelstufe, Sprache: Deutsch"],
    ["Dauer", "90 Minuten (unter der Höchstgrenze von 2 Stunden)"],
    ["Methode", "Kurze Vorführungen am Beamer + Praxis am PC in Kleingruppen (3–5 Personen) + Quiz"],
  ], { firstColFill: true, firstColBold: true }),
  H1("Lernziele"),
  P("Realistisch und überprüfbar – die Teilnehmenden können nach der Schulung …"),
  bullet("Zeichen- und Absatzformate unterscheiden und anwenden,"),
  bullet("Formatvorlagen zuweisen, ändern und ein automatisches Inhaltsverzeichnis erstellen,"),
  bullet("Kopf- und Fußzeilen mit Seitenzahl und Datum einfügen (inkl. „Erste Seite anders“)."),
  H1("Ablauf"),
  table([1250, 2000, 3488, 1300, 1600], ["Zeit", "Phase", "Inhalt / Methode", "Wer", "Material"], [
    ["08:30\n10 min", "Einstieg", "Freundliche Begrüßung, Vorstellung, Ablauf zeigen (Schulungsblatt S. 1). Frage an die Gruppe: „Wer hat schon mal ein Referat in Word geschrieben? Was hat genervt?“", "Yannik", "Beamer, Schulungsblatt"],
    ["08:40\n15 min", "Teil 1: Formate", "Live-Vorführung: Zeichen- vs. Absatzformat, Tastenkürzel, Format übertragen. Zwischenfrage: „Ist Zentrieren ein Zeichen- oder Absatzformat?“", "Nico", "Beamer, Übungsdatei"],
    ["08:55\n15 min", "Teil 2: Formatvorlagen", "Vorher/Nachher zeigen: Überschriften per Hand vs. Vorlage. Vorlage ändern → alles passt sich an („Wow-Effekt“). Navigationsbereich, Inhaltsverzeichnis.", "Yannik", "Beamer, Übungsdatei"],
    ["09:10\n10 min", "Teil 3: Kopf- & Fußzeile", "Doppelklick in den Rand, Name/Klasse, Seitenzahl, Datum, „Erste Seite anders“.", "Nico", "Beamer, Übungsdatei"],
    ["09:20\n25 min", "Praxis", "Kleingruppen (3–5 Pers.) bearbeiten die Übungsaufgaben am PC. Beide gehen herum und helfen. Schnelle Gruppen: Bonusaufgabe.", "Beide", "PCs, Übungsdatei, Schulungsblatt"],
    ["09:45\n7 min", "Lernkontrolle", "Quizfragen gemeinsam besprechen, offene Fragen klären.", "Beide", "Schulungsblatt (Quiz)"],
    ["09:52\n8 min", "Abschluss", "Handout austeilen (erst jetzt!), Feedbackbogen ausfüllen lassen, Teilnahmebescheinigungen übergeben, Dank & Verabschiedung.", "Yannik / Nico", "Handout, Feedbackbogen, Bescheinigungen"],
  ], { zebra: true }),
  H1("Checkliste Vorbereitung"),
  H2("Einige Tage vorher"),
  check("Einladung verteilen (Datum, Uhrzeit, Ort, Thema) und Anmeldungen zählen"),
  check("Raum 0.68 reservieren, genug Tische, Stühle und PCs prüfen"),
  check("Übungsdatei auf alle PCs bzw. in einen gemeinsamen Ordner kopieren"),
  check("Schulungsblatt, Feedbackbögen und Teilnahmebescheinigungen drucken (Anzahl = Anmeldungen + 2)"),
  check("Handouts drucken – werden erst am Ende ausgeteilt"),
  check("Ablauf gemeinsam einmal komplett durchspielen (Probedurchlauf mit Uhr)"),
  H2("Am Schulungstag"),
  check("15 Minuten früher da sein, Raum lüften"),
  check("Beamer, PC und Word testen; Übungsdatei geöffnet bereithalten"),
  check("Ausgeschlafen und fit sein 🙂"),
  H1("Tipps für die Durchführung"),
  bullet([r("Stichpunkte als roten Faden ", { bold: true }), r("nutzen – frei sprechen, nicht ablesen.")]),
  bullet([r("Stimme und Tempo variieren, ", { bold: true }), r("bei Vorführungen langsam klicken und jeden Schritt ansagen.")]),
  bullet([r("Vortragsteile kurz halten ", { bold: true }), r("und die Teilnehmenden immer wieder mit Fragen einbinden.")]),
  bullet([r("Moderatorrolle einnehmen: ", { bold: true }), r("gezielte Fragen stellen, alle zu Wort kommen lassen.")]),
  bullet([r("Rollen klar trennen: ", { bold: true }), r("Wer nicht spricht, bedient den PC oder hilft einzelnen Teilnehmenden.")]),
  H1("Lösungen zum Quiz"),
  table([600, W - 600], ["Nr.", "Musterlösung"], [
    ["1", "Zeichenformat gilt nur für markierte Zeichen (z. B. Fett, Farbe); Absatzformat gilt für den ganzen Absatz (z. B. Ausrichtung, Zeilenabstand)."],
    ["2", "Einheitliches Aussehen; Änderungen nur einmal nötig; Zeitersparnis; Inhaltsverzeichnis und Navigationsbereich möglich."],
    ["3", "Word erkennt Überschriften nur an ihrer Formatvorlage (Überschrift 1, 2, …), nicht an Größe oder Fettdruck."],
    ["4", "Doppelklick in den oberen Seitenrand bzw. Einfügen › Kopfzeile; heraus per Doppelklick in den Text oder „Kopf- und Fußzeile schließen“."],
    ["5", "Getippte Zahlen wären auf jeder Seite gleich bzw. passen sich nicht an, wenn Seiten hinzukommen oder wegfallen."],
  ], { zebra: true }),
];
await save("00_Ablaufplan_intern", [{ properties: { page }, headers: { default: header("Ablaufplan (intern)") }, footers: { default: footer() }, children: ap }]);
})();
