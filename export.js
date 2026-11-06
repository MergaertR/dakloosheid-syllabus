// Word export. Every special block gets a named paragraph style, so render.js can turn
// the .docx back into the same website blocks after Ruben edits it in Word.
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, ExternalHyperlink, BorderStyle, LevelFormat, AlignmentType,
} = docx;

const ACCENT = "F2C500";
const SOFT = "FFF4C2";
const WARN_SOFT = "FDE6DC";
const MUTED = "6B645A";

const clean = (s) => s.replace(/\s+/g, " ");

function inlineRuns(node, fmt = {}) {
  const runs = [];
  for (const child of node.childNodes) {
    if (child.nodeType === Node.TEXT_NODE) {
      const text = clean(child.textContent);
      if (text.trim() || text === " ") runs.push(new TextRun({ text, ...fmt }));
    } else if (child.nodeType === Node.ELEMENT_NODE) {
      const tag = child.tagName;
      if (tag === "STRONG" || tag === "B") runs.push(...inlineRuns(child, { ...fmt, bold: true }));
      else if (tag === "EM" || tag === "I" || tag === "CITE") runs.push(...inlineRuns(child, { ...fmt, italics: true }));
      else if (tag === "A") {
        runs.push(new ExternalHyperlink({
          link: child.href,
          children: [new TextRun({ text: clean(child.textContent), style: "Hyperlink", ...fmt })],
        }));
      } else runs.push(...inlineRuns(child, fmt));
    }
  }
  return runs;
}

const styled = (style, children) => new Paragraph({ style, children });

function table(el) {
  const rows = [...el.querySelectorAll("tr")].map((tr) => {
    const isHead = tr.parentElement.tagName === "THEAD";
    return new TableRow({
      // mammoth reads an explicit tblHeader=false as a header row, so only set it when true.
      ...(isHead ? { tableHeader: true } : {}),
      children: [...tr.children].map((cell, c) => new TableCell({
        shading: isHead ? { fill: ACCENT, type: ShadingType.CLEAR, color: "auto" } : undefined,
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        children: [new Paragraph({ children: inlineRuns(cell, { bold: isHead || c === 0 }) })],
      })),
    });
  });
  return [new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }), new Paragraph({ text: "" })];
}

function splitItem(li, keySel) {
  return [
    new TextRun({ text: li.querySelector(keySel).textContent + ": ", bold: true }),
    new TextRun(clean(li.querySelector(".what").textContent)),
  ];
}

let listInstance = 0;

function convert(el) {
  const tag = el.tagName;
  const cls = el.classList;
  if (cls.contains("end-mark")) return [];
  if (cls.contains("check")) return [styled("Check", inlineRuns(el))];
  if (tag === "H1") return [new Paragraph({ heading: HeadingLevel.TITLE, children: inlineRuns(el) })];
  if (tag === "H2") return [new Paragraph({ heading: HeadingLevel.HEADING_1, children: inlineRuns(el) })];
  if (tag === "H3" || tag === "SUMMARY") return [new Paragraph({ heading: HeadingLevel.HEADING_2, children: inlineRuns(el) })];
  if (tag === "H4") return [new Paragraph({ heading: HeadingLevel.HEADING_3, children: inlineRuns(el) })];
  if (tag === "P") {
    if (cls.contains("kicker")) return [styled("Subtitle", inlineRuns(el))];
    if (cls.contains("lead")) return [styled("Intro", inlineRuns(el))];
    return [new Paragraph({ children: inlineRuns(el) })];
  }
  if (tag === "TABLE") return table(el);
  if (cls.contains("note")) {
    const style = cls.contains("warn") ? "Waarschuwing" : "Kader";
    return [...el.children].map((child) => styled(style, inlineRuns(child, child.tagName === "STRONG" ? { bold: true } : {})));
  }
  if (cls.contains("stats")) {
    return [...el.querySelectorAll(".stat")].map((s) => styled("Kerncijfer", [
      new TextRun({ text: s.querySelector(".num").textContent + ": ", bold: true }),
      new TextRun(clean(s.querySelector(".lbl").textContent)),
    ]));
  }
  if (cls.contains("chain")) return [...el.children].map((li) => styled("Keten", splitItem(li, ".lvl")));
  if (cls.contains("timeline")) return [...el.children].map((li) => styled("Tijdlijn", splitItem(li, ".year")));
  if (tag === "BLOCKQUOTE") {
    return [
      styled("Citaat", inlineRuns(el.querySelector("p"))),
      styled("Citaatbron", [new TextRun(el.querySelector("cite").textContent)]),
    ];
  }
  if (tag === "UL") return [...el.children].map((li) => new Paragraph({ bullet: { level: 0 }, children: inlineRuns(li) }));
  if (tag === "OL") {
    const instance = ++listInstance;
    return [...el.children].map((li) => new Paragraph({ numbering: { reference: "genummerd", level: 0, instance }, children: inlineRuns(li) }));
  }
  return [...el.children].flatMap(convert);
}

const boxStyle = (id, fill, borderColor, extra = {}) => ({
  id, name: id, basedOn: "Normal", next: "Normal", quickFormat: true,
  run: extra.run,
  paragraph: {
    shading: { fill, type: ShadingType.CLEAR, color: "auto" },
    border: { left: { style: extra.dashed ? BorderStyle.DASHED : BorderStyle.SINGLE, size: 24, color: borderColor } },
    indent: { left: 120 },
    spacing: { after: 120 },
  },
});

function buildDocument(sections) {
  listInstance = 0;
  const children = sections.flatMap(convert);
  return new Document({
    creator: "Ruben Mergaert",
    title: document.title,
    numbering: {
      config: [{
        reference: "genummerd",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.START, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      }],
    },
    styles: {
      default: { document: { run: { font: "Calibri", size: 22 }, paragraph: { spacing: { after: 120 } } } },
      paragraphStyles: [
        { id: "Title", name: "Title", basedOn: "Normal", next: "Normal", run: { font: "Georgia", size: 48, bold: true }, paragraph: { spacing: { after: 200 } } },
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", run: { font: "Georgia", size: 32, bold: true, color: "1D1B18" }, paragraph: { spacing: { before: 360, after: 160 } } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", run: { size: 25, bold: true, color: "1D1B18" }, paragraph: { spacing: { before: 240, after: 100 } } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", run: { size: 23, bold: true, color: MUTED }, paragraph: { spacing: { before: 200, after: 80 } } },
        { id: "Subtitle", name: "Subtitle", basedOn: "Normal", next: "Normal", quickFormat: true, run: { allCaps: true, bold: true, color: MUTED, size: 18 } },
        { id: "Intro", name: "Intro", basedOn: "Normal", next: "Normal", quickFormat: true, run: { italics: true, color: MUTED, size: 24 } },
        boxStyle("Kader", SOFT, ACCENT),
        boxStyle("Waarschuwing", WARN_SOFT, "C2491D"),
        boxStyle("Kerncijfer", "F6F2EA", ACCENT),
        boxStyle("Keten", "F6F2EA", ACCENT),
        boxStyle("Tijdlijn", "FFFFFF", ACCENT),
        boxStyle("Check", SOFT, "C29A00", { dashed: true, run: { color: "6B5300" } }),
        { id: "Citaat", name: "Citaat", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: "Georgia", italics: true, size: 26 }, paragraph: { indent: { left: 567 }, spacing: { after: 0 } } },
        { id: "Citaatbron", name: "Citaatbron", basedOn: "Normal", next: "Normal", quickFormat: true, run: { color: MUTED, size: 20 }, paragraph: { indent: { left: 567 }, spacing: { after: 200 } } },
      ],
    },
    sections: [{ children }],
  });
}

async function exportDocx() {
  const btn = document.getElementById("export-docx");
  btn.disabled = true;
  try {
    const blob = await Packer.toBlob(buildDocument([...document.querySelectorAll("#syllabus > section")]));
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = document.getElementById("syllabus").dataset.docxName;
    a.click();
    URL.revokeObjectURL(a.href);
  } catch (err) {
    console.error("[syllabus] docx export failed", err);
    btn.textContent = "Export mislukt";
  } finally {
    btn.disabled = false;
  }
}

document.getElementById("export-docx")?.addEventListener("click", exportDocx);
