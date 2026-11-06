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
];

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

async function renderSyllabus(arrayBuffer, main) {
  const { value, messages } = await mammoth.convertToHtml({ arrayBuffer }, { styleMap: STYLE_MAP });
  const tmp = document.createElement("div");
  tmp.innerHTML = value;

  tmp.querySelectorAll("td > p:only-child, th > p:only-child").forEach((p) => p.replaceWith(...p.childNodes));
  tmp.querySelectorAll("p").forEach((p) => { if (!p.textContent.trim()) p.remove(); });
  tmp.querySelectorAll(".note-warn").forEach((d) => { d.className = "note warn"; });
  tmp.querySelectorAll("p.stat").forEach((p) => splitKey(p, "num", "lbl"));
  tmp.querySelectorAll("ol.chain > li").forEach((li) => splitKey(li, "lvl", "what"));
  tmp.querySelectorAll("ol.timeline > li").forEach((li) => splitKey(li, "year", "what"));
  tmp.querySelectorAll("table").forEach((t) => { if (!t.querySelector("thead")) t.classList.add("plain"); });

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

// 1 self-contained file without scripts, checks or site navigation: this is what Ruben submits.
async function downloadStandalone(main, toc) {
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
  saveBlob(new Blob([html], { type: "text/html" }), "syllabus-sociaal-beleid-ruben-mergaert.html");
}
