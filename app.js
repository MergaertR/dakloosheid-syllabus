const root = document.documentElement;
const themeBtn = document.getElementById("theme-toggle");

function applyTheme(theme) {
  root.dataset.theme = theme;
  themeBtn.textContent = theme === "dark" ? "Licht" : "Donker";
  localStorage.setItem("syllabus-theme", theme);
}
applyTheme(localStorage.getItem("syllabus-theme")
  || (matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light"));
themeBtn.addEventListener("click", () => applyTheme(root.dataset.theme === "dark" ? "light" : "dark"));

const tocLinks = [...document.querySelectorAll(".toc a")];
const observer = new IntersectionObserver((entries) => {
  for (const entry of entries) {
    if (!entry.isIntersecting) continue;
    tocLinks.forEach((a) => a.classList.toggle("active", a.hash === "#" + entry.target.id));
  }
}, { rootMargin: "-30% 0px -60% 0px" });
document.querySelectorAll("main section[id]").forEach((s) => observer.observe(s));

const checkBtn = document.getElementById("check-toggle");
if (checkBtn) {
  const applyChecks = (hidden) => {
    document.body.classList.toggle("hide-checks", hidden);
    checkBtn.textContent = hidden ? "Toon checks" : "Verberg checks";
    localStorage.setItem("syllabus-hide-checks", hidden ? "1" : "");
  };
  applyChecks(localStorage.getItem("syllabus-hide-checks") === "1");
  checkBtn.addEventListener("click", () => applyChecks(!document.body.classList.contains("hide-checks")));
}

// Builds the Word file from the live page, so the website and the .docx never drift apart.
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, ExternalHyperlink, BorderStyle,
} = docx;

const ACCENT = "F2C500";
const SOFT = "FFF4C2";
const WARN_SOFT = "FDE6DC";

const clean = (s) => s.replace(/\s+/g, " ");

function inlineRuns(node, fmt = {}) {
  const runs = [];
  for (const child of node.childNodes) {
    if (child.nodeType === Node.TEXT_NODE) {
      const text = clean(child.textContent);
      if (text.trim() || text === " ") runs.push(new TextRun({ text, ...fmt }));
    } else if (child.nodeType === Node.ELEMENT_NODE) {
      if (child.classList.contains("check")) continue;
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

function para(el, opts = {}) {
  return new Paragraph({ children: inlineRuns(el, opts.fmt), spacing: { after: 120 }, ...opts.para });
}

function table(el) {
  const rows = [...el.querySelectorAll("tr")].map((tr, i) => {
    const isHead = tr.parentElement.tagName === "THEAD";
    return new TableRow({
      tableHeader: isHead,
      children: [...tr.children].map((cell, c) => new TableCell({
        shading: isHead ? { fill: ACCENT, type: ShadingType.CLEAR, color: "auto" } : undefined,
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        children: [new Paragraph({ children: inlineRuns(cell, { bold: isHead || c === 0 }) })],
      })),
    });
  });
  return [new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }), new Paragraph({ text: "" })];
}

function note(el) {
  const fill = el.classList.contains("warn") ? WARN_SOFT : SOFT;
  const border = { left: { style: BorderStyle.SINGLE, size: 24, color: el.classList.contains("warn") ? "C2491D" : ACCENT } };
  const out = [];
  for (const child of el.children) {
    out.push(new Paragraph({
      children: inlineRuns(child, child.tagName === "STRONG" ? { bold: true } : {}),
      shading: { fill, type: ShadingType.CLEAR, color: "auto" },
      border,
      spacing: { after: child.tagName === "STRONG" ? 0 : 160 },
      indent: { left: 120 },
    }));
  }
  return out;
}

function convert(el) {
  const tag = el.tagName;
  const cls = el.classList;
  // Checks are notes for Ruben while editing; they must never reach the submitted document.
  if (cls.contains("end-mark") || cls.contains("check")) return [];
  if (tag === "H1") return [new Paragraph({ heading: HeadingLevel.TITLE, children: inlineRuns(el) })];
  if (tag === "H2") return [new Paragraph({ heading: HeadingLevel.HEADING_1, children: inlineRuns(el), spacing: { before: 360, after: 160 } })];
  if (tag === "H3") return [new Paragraph({ heading: HeadingLevel.HEADING_3, children: inlineRuns(el), spacing: { before: 240, after: 100 } })];
  if (tag === "SUMMARY") return [new Paragraph({ heading: HeadingLevel.HEADING_2, children: inlineRuns(el) })];
  if (tag === "P") {
    if (cls.contains("kicker")) return [para(el, { fmt: { allCaps: true, bold: true, color: "6B645A" } })];
    if (cls.contains("lead")) return [para(el, { fmt: { italics: true } })];
    return [para(el)];
  }
  if (tag === "TABLE") return table(el);
  if (cls.contains("note")) return note(el);
  if (cls.contains("stats")) {
    return [...el.querySelectorAll(".stat")].map((s) => new Paragraph({
      bullet: { level: 0 },
      children: [
        new TextRun({ text: s.querySelector(".num").textContent + ": ", bold: true }),
        new TextRun(s.querySelector(".lbl").textContent),
      ],
    }));
  }
  if (tag === "BLOCKQUOTE") {
    return [
      new Paragraph({ children: inlineRuns(el.querySelector("p"), { italics: true }), indent: { left: 567 }, spacing: { after: 0 } }),
      new Paragraph({ children: [new TextRun({ text: el.querySelector("cite").textContent, color: "6B645A", size: 20 })], indent: { left: 567 }, spacing: { after: 200 } }),
    ];
  }
  if (cls.contains("timeline") || cls.contains("chain")) {
    const key = cls.contains("timeline") ? ".year" : ".lvl";
    return [...el.children].map((li) => new Paragraph({
      bullet: { level: 0 },
      children: [
        new TextRun({ text: li.querySelector(key).textContent + ": ", bold: true }),
        new TextRun(li.querySelector(".what").textContent),
      ],
    }));
  }
  if (tag === "UL" || tag === "OL") {
    return [...el.children].map((li) => new Paragraph({ bullet: { level: 0 }, children: inlineRuns(li) }));
  }
  return [...el.children].flatMap(convert);
}

async function exportDocx() {
  const btn = document.getElementById("export-docx");
  btn.disabled = true;
  try {
    const children = [...document.querySelectorAll("#syllabus > section")].flatMap(convert);
    children.push(new Paragraph({
      children: [new TextRun({ text: document.querySelector(".foot").textContent, italics: true, color: "6B645A", size: 18 })],
      spacing: { before: 400 },
    }));
    const doc = new Document({
      creator: "Student",
      title: document.title,
      styles: {
        default: { document: { run: { font: "Calibri", size: 22 } } },
        paragraphStyles: [
          { id: "Title", name: "Title", basedOn: "Normal", next: "Normal", run: { font: "Georgia", size: 48, bold: true }, paragraph: { spacing: { after: 200 } } },
          { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", run: { font: "Georgia", size: 32, bold: true, color: "1D1B18" } },
          { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", run: { size: 24, bold: true, color: "6B645A" }, paragraph: { spacing: { before: 200, after: 100 } } },
          { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", run: { size: 23, bold: true, color: "1D1B18" } },
        ],
      },
      sections: [{ children }],
    });
    const blob = await Packer.toBlob(doc);
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

document.getElementById("export-docx").addEventListener("click", exportDocx);
