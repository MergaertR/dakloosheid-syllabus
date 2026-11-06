// Turns the Word file into the website. The paragraph styles match the ones app.js writes,
// so a block Ruben styles as "Kader" in Word becomes a yellow box on the site.
const STYLE_MAP = [
  "p[style-name='Title'] => h1:fresh",
  "p[style-name='Subtitle'] => p.kicker:fresh",
  "p[style-name='Intro'] => p.lead:fresh",
  "p[style-name='Heading 1'] => h2:fresh",
  "p[style-name='Heading 2'] => h3:fresh",
  "p[style-name='Heading 3'] => h4:fresh",
  "p[style-name='Kader'] => div.note > p:fresh",
  "p[style-name='Waarschuwing'] => div.note-warn > p:fresh",
  "p[style-name='Kerncijfer'] => div.stats > p.stat:fresh",
  "p[style-name='Keten'] => ol.chain > li:fresh",
  "p[style-name='Tijdlijn'] => ol.timeline > li:fresh",
  "p[style-name='Check'] => p.check:fresh",
  "p[style-name='Citaat'] => blockquote > p:fresh",
  "p[style-name='Citaatbron'] => blockquote > cite:fresh",
  "p[style-name='Uitklapper'] => p.fold-open:fresh",
  "p[style-name='Uitklapper dicht'] => p.fold-closed:fresh",
  "p[style-name='Uitklapper einde'] => p.end-mark:fresh",
];

// A collapsible runs from its "Uitklapper" title to the next "Uitklapper einde", the next
// collapsible or the next chapter, whichever comes first.
function buildFolds(container) {
  for (const title of [...container.querySelectorAll(":scope > p.fold-open, :scope > p.fold-closed")]) {
    const details = document.createElement("details");
    details.open = title.classList.contains("fold-open");
    const summary = document.createElement("summary");
    summary.append(...title.childNodes);
    details.append(summary);
    title.replaceWith(details);
    let next = details.nextSibling;
    while (next && !next.matches?.("h2, p.fold-open, p.fold-closed")) {
      const after = next.nextSibling;
      details.append(next);
      if (next.matches?.("p.end-mark")) break;
      next = after;
    }
  }
}

// mammoth drops direct formatting. Before converting, the formatting we keep (text color,
// background, alignment) gets rewritten into marker character styles that mammoth passes on
// as span classes; restoreWordFormatting() turns those back into inline styles.
// Font family and font size are left out on purpose, so the site keeps its own typography.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const HIGHLIGHTS = {
  yellow: "FFFF00", green: "00FF00", cyan: "00FFFF", magenta: "FF00FF", blue: "0000FF", red: "FF0000",
  darkBlue: "000080", darkCyan: "008080", darkGreen: "008000", darkMagenta: "800080", darkRed: "800000",
  darkYellow: "808000", darkGray: "808080", lightGray: "C0C0C0",
};
const ALIGN = { center: "center", right: "right", end: "right" };

const wAttr = (el, name) => el?.getAttributeNS(W, name) ?? el?.getAttribute("w:" + name);
const wChild = (el, name) => el && [...el.children].find((c) => c.localName === name);
const isHex = (v) => /^[0-9A-Fa-f]{6}$/.test(v || "");
// Near-black text would vanish in dark mode and only repeats the default ink color.
const isInk = (hex) => [0, 2, 4].every((i) => parseInt(hex.slice(i, i + 2), 16) < 0x40);

function wEl(doc, name, val) {
  const el = doc.createElementNS(W, "w:" + name);
  if (val !== undefined) el.setAttributeNS(W, "w:val", val);
  return el;
}

async function prepareWordFormatting(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docFile = zip.file("word/document.xml");
  const stylesFile = zip.file("word/styles.xml");
  if (!docFile || !stylesFile) return { buffer: arrayBuffer, styleMap: [] };
  const parser = new DOMParser();
  const doc = parser.parseFromString(await docFile.async("string"), "application/xml");
  const styles = parser.parseFromString(await stylesFile.async("string"), "application/xml");
  const markers = new Set();

  for (const run of doc.getElementsByTagNameNS(W, "r")) {
    const rPr = wChild(run, "rPr");
    if (!rPr || wChild(rPr, "rStyle")) continue;
    const color = wAttr(wChild(rPr, "color"), "val");
    const highlight = HIGHLIGHTS[wAttr(wChild(rPr, "highlight"), "val")];
    const shading = wAttr(wChild(rPr, "shd"), "fill");
    const bg = highlight || (isHex(shading) && shading.toUpperCase() !== "FFFFFF" ? shading : null);
    const parts = [];
    if (isHex(color) && !isInk(color)) parts.push("c-" + color.toUpperCase());
    if (bg) parts.push("bg-" + bg.toUpperCase());
    if (!parts.length) continue;
    const id = "fx-" + parts.join("-");
    rPr.insertBefore(wEl(doc, "rStyle", id), rPr.firstChild);
    markers.add(id);
  }

  for (const p of doc.getElementsByTagNameNS(W, "p")) {
    const pPr = wChild(p, "pPr");
    const align = ALIGN[wAttr(wChild(pPr, "jc"), "val")];
    if (!align) continue;
    const id = "fx-align-" + align;
    const run = wEl(doc, "r");
    const rPr = wEl(doc, "rPr");
    rPr.append(wEl(doc, "rStyle", id));
    const t = wEl(doc, "t");
    t.textContent = "​";
    run.append(rPr, t);
    p.insertBefore(run, pPr ? pPr.nextSibling : p.firstChild);
    markers.add(id);
  }

  if (!markers.size) return { buffer: arrayBuffer, styleMap: [] };
  const root = styles.documentElement;
  for (const id of markers) {
    const style = wEl(styles, "style");
    style.setAttributeNS(W, "w:type", "character");
    style.setAttributeNS(W, "w:styleId", id);
    style.append(wEl(styles, "name", id));
    root.append(style);
  }
  const xml = new XMLSerializer();
  zip.file("word/document.xml", xml.serializeToString(doc));
  zip.file("word/styles.xml", xml.serializeToString(styles));
  return {
    buffer: await zip.generateAsync({ type: "arraybuffer" }),
    styleMap: [...markers].map((id) => `r[style-name='${id}'] => span.${id}`),
  };
}

function restoreWordFormatting(container) {
  container.querySelectorAll("span[class^='fx-align-']").forEach((marker) => {
    const block = marker.closest("p, li, h1, h2, h3, h4, td, th");
    if (block) block.style.textAlign = marker.className.slice("fx-align-".length);
    marker.remove();
  });
  container.querySelectorAll("span[class^='fx-']").forEach((span) => {
    const color = span.className.match(/c-([0-9A-F]{6})/)?.[1];
    const bg = span.className.match(/bg-([0-9A-F]{6})/)?.[1];
    if (color) span.style.color = "#" + color;
    if (bg) {
      span.style.backgroundColor = "#" + bg;
      if (!color) span.style.color = "#1d1b18";
      span.style.borderRadius = "3px";
      span.style.padding = "0 .15em";
    }
    span.removeAttribute("class");
  });
}

function splitKey(el, keyClass, restClass, keyTag = "span") {
  const text = el.textContent;
  const i = text.indexOf(":");
  if (i < 0) return;
  const key = document.createElement(keyTag);
  key.className = keyClass;
  key.textContent = text.slice(0, i).trim();
  const rest = document.createElement("span");
  rest.className = restClass;
  rest.textContent = text.slice(i + 1).trim();
  el.replaceChildren(key, rest);
}

function slug(text) {
  return text.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 40) || "deel";
}

async function renderDocx(arrayBuffer, main) {
  const prepared = await prepareWordFormatting(arrayBuffer);
  const { value, messages } = await mammoth.convertToHtml(
    { arrayBuffer: prepared.buffer },
    { styleMap: [...STYLE_MAP, "u => u", ...prepared.styleMap] },
  );
  const tmp = document.createElement("div");
  tmp.innerHTML = value;

  tmp.querySelectorAll("td > p:only-child, th > p:only-child").forEach((p) => p.replaceWith(...p.childNodes));
  restoreWordFormatting(tmp);
  tmp.querySelectorAll("p").forEach((p) => { if (!p.textContent.trim()) p.remove(); });
  tmp.querySelectorAll(".note-warn").forEach((d) => { d.className = "note warn"; });
  tmp.querySelectorAll("p.stat").forEach((p) => splitKey(p, "num", "lbl"));
  tmp.querySelectorAll("ol.chain > li").forEach((li) => splitKey(li, "lvl", "what"));
  tmp.querySelectorAll("ol.timeline > li").forEach((li) => splitKey(li, "year", "what"));
  tmp.querySelectorAll("table").forEach((t) => { if (!t.querySelector("thead")) t.classList.add("plain"); });
  buildFolds(tmp);

  const sections = [];
  let current = null;
  const open = (id, cls) => {
    current = document.createElement("section");
    current.id = id;
    if (cls) current.className = cls;
    sections.push(current);
  };
  open("titel", "hero");
  for (const node of [...tmp.childNodes]) {
    if (node.tagName === "H2") open(slug(node.textContent));
    current.append(node);
  }
  const hero = sections[0];
  if (!hero.childNodes.length) sections.shift();
  else hero.querySelector("table")?.classList.add("meta");

  main.replaceChildren(...sections);
  return {
    checks: main.querySelectorAll(".check").length,
    unknownStyles: [...new Set(messages
      .map((m) => m.message.match(/Unrecognised paragraph style: '([^']+)'/)?.[1])
      .filter(Boolean))],
  };
}

function buildToc(main, list) {
  list.replaceChildren(...[...main.querySelectorAll(":scope > section")].map((s) => {
    const li = document.createElement("li");
    const a = document.createElement("a");
    a.href = "#" + s.id;
    a.textContent = s.classList.contains("hero") ? "Titelblad" : s.querySelector("h2").textContent;
    li.append(a);
    return li;
  }));
}

function saveBlob(blob, name) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

// Public pages: load their Word file, render it and wire the buttons.
async function bootPage(docxPath, htmlName) {
  const main = document.getElementById("syllabus");
  const toc = document.getElementById("toc-list");
  try {
    const res = await fetch(docxPath + "?v=" + Date.now(), { cache: "no-store" });
    if (!res.ok) throw new Error("HTTP " + res.status);
    await renderDocx(await res.arrayBuffer(), main);
    buildToc(main, toc);
    initToc();
    if (location.hash) document.querySelector(location.hash)?.scrollIntoView();
  } catch (err) {
    console.error("[page] laden mislukt", err);
    main.innerHTML = "<section><h2>Deze pagina kon niet laden</h2><p>Open de site via een webserver of GitHub Pages. Een lokaal geopend bestand mag het Word-bestand niet inlezen.</p></section>";
  }
  document.getElementById("download-page").addEventListener("click", () => downloadStandalone(main, toc, htmlName));
}

// 1 self-contained file without scripts, checks or site navigation: this is what Ruben submits.
async function downloadStandalone(main, toc, fileName) {
  const css = await (await fetch("style.css")).text();
  const body = main.cloneNode(true);
  body.querySelectorAll(".check").forEach((el) => el.remove());
  const title = body.querySelector("h1")?.textContent || document.title;
  const html = `<!DOCTYPE html>
<html lang="nl-BE" data-theme="light">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title.replace(/</g, "&lt;")}</title>
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,700&family=Source+Sans+3:wght@400;600;700&display=swap" rel="stylesheet">
<style>${css}</style>
</head>
<body>
<header class="topbar"><span class="brand"><span class="lion" aria-hidden="true"></span>Sociaal beleid 2026-2027</span></header>
<div class="layout">
<aside class="toc"><p class="toc-title">Inhoud</p>${toc.outerHTML}</aside>
${body.outerHTML}
</div>
</body>
</html>`;
  saveBlob(new Blob([html], { type: "text/html" }), fileName);
}
