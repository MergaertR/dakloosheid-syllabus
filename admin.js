// The GitHub token lives in admin-key.json, encrypted with a key derived from the password.
// The repo never holds the password or a readable token. tools/set-password.mjs writes the file.
const REPO = { owner: "MergaertR", name: "dakloosheid-syllabus", branch: "main" };
const PAGES = {
  syllabus: { name: "syllabus", path: "content/syllabus.docx", initial: "content/syllabus-initial.docx", url: "index.html", file: "syllabus-sociaal-beleid-ruben-mergaert" },
  onderzoek: { name: "onderzoek", path: "content/onderzoek.docx", initial: "content/onderzoek-initial.docx", url: "onderzoek.html", file: "onderzoeksdossier-dakloosheid-brugge" },
};
let page = PAGES.syllabus;

const $ = (id) => document.getElementById(id);
const fromB64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));

function toB64(bytes) {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

async function unlockToken(password) {
  const res = await fetch("admin-key.json?v=" + Date.now(), { cache: "no-store" });
  if (!res.ok) throw new Error("no-key");
  const box = await res.json();
  const material = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveKey"]);
  const key = await crypto.subtle.deriveKey(
    { name: "PBKDF2", hash: "SHA-256", salt: fromB64(box.salt), iterations: box.iterations },
    material, { name: "AES-GCM", length: 256 }, false, ["decrypt"],
  );
  const plain = await crypto.subtle.decrypt({ name: "AES-GCM", iv: fromB64(box.iv) }, key, fromB64(box.data));
  return new TextDecoder().decode(plain);
}

async function github(token, method, body) {
  const url = `https://api.github.com/repos/${REPO.owner}/${REPO.name}/contents/${page.path}`;
  return fetch(method === "GET" ? `${url}?ref=${REPO.branch}` : url, {
    method,
    headers: { Authorization: `Bearer ${token}`, Accept: "application/vnd.github+json" },
    body: body && JSON.stringify(body),
  });
}

let token = null;
let fileBytes = null;

function choosePage(key) {
  page = PAGES[key];
  const link = $("download-current");
  link.href = page.path;
  link.download = page.file + ".docx";
  fileBytes = null;
  $("preview").hidden = true;
  $("status").hidden = true;
  $("publish").disabled = true;
  $("download-page").disabled = true;
  for (const id of ["publish-msg", "reset-msg"]) $(id).textContent = "";
  disarmReset();
}
document.querySelectorAll('input[name="page"]').forEach((r) => r.addEventListener("change", () => choosePage(r.value)));

$("login-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const msg = $("login-msg");
  msg.textContent = "Even geduld...";
  msg.className = "msg";
  try {
    token = await unlockToken($("password").value);
    $("login").hidden = true;
    $("workspace").hidden = false;
  } catch (err) {
    msg.className = "msg error";
    msg.textContent = err.message === "no-key"
      ? "Er is nog geen wachtwoord ingesteld. Voer eerst tools/set-password.mjs uit."
      : "Fout wachtwoord.";
  }
});

async function loadFile(file) {
  const status = $("status");
  status.hidden = false;
  if (!file.name.toLowerCase().endsWith(".docx")) {
    status.className = "status error";
    status.textContent = "Dit is geen .docx-bestand. Sla je document in Word op als Word-document (.docx).";
    return;
  }
  const buffer = await file.arrayBuffer();
  const main = $("syllabus");
  const result = await renderDocx(buffer.slice(0), main);
  buildToc(main, $("toc-list"));
  $("preview").hidden = false;
  fileBytes = new Uint8Array(buffer);

  const lines = [`<strong>${file.name.replace(/</g, "&lt;")}</strong> is ingeladen. Het voorbeeld staat hieronder.`];
  if (result.checks) lines.push(`Er staan nog ${result.checks} checks in je tekst. Verwijder ze voor je indient.`);
  const unknown = result.unknownStyles.filter((s) => !/^(Normal|List Paragraph|Standaard|Lijstalinea)$/i.test(s));
  if (unknown.length) lines.push(`Deze Word-stijlen kent de site niet en worden gewone tekst: ${unknown.join(", ")}.`);
  status.className = result.checks || unknown.length ? "status warn" : "status ok";
  status.innerHTML = lines.map((l) => `<p>${l}</p>`).join("");
  $("publish").disabled = false;
  $("download-page").disabled = false;
  $("publish-msg").textContent = "";
}

const drop = $("drop");
$("file").addEventListener("change", (e) => e.target.files[0] && loadFile(e.target.files[0]));
drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("over"); });
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", (e) => {
  e.preventDefault();
  drop.classList.remove("over");
  if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]);
});

$("download-page").addEventListener("click", () => downloadStandalone($("syllabus"), $("toc-list"), page.file + ".html"));

async function publish(bytes, commitMessage) {
  const current = await github(token, "GET");
  if (!current.ok && current.status !== 404) throw new Error(`GitHub antwoordde ${current.status}`);
  const sha = current.ok ? (await current.json()).sha : undefined;
  const res = await github(token, "PUT", { message: commitMessage, content: toB64(bytes), branch: REPO.branch, sha });
  if (!res.ok) throw new Error(`GitHub antwoordde ${res.status}: ${(await res.json()).message}`);
}


async function run(btn, msg, busyText, task) {
  btn.disabled = true;
  msg.className = "msg";
  msg.textContent = busyText;
  try {
    await task();
    msg.className = "msg ok";
    msg.innerHTML = `Gepubliceerd. Over 1 à 2 minuten staat de nieuwe versie op <a href="${page.url}">de pagina ${page.name}</a>.`;
  } catch (err) {
    console.error("[admin] mislukt", err);
    msg.className = "msg error";
    msg.textContent = `Mislukt. ${err.message}. Probeer opnieuw, of bewaar deze melding.`;
  } finally {
    btn.disabled = false;
  }
}

$("publish").addEventListener("click", () =>
  run($("publish"), $("publish-msg"), "Bezig met publiceren...", () => publish(fileBytes, `Update ${page.name}`)));

// Reset takes 2 clicks: the 1st saves a backup of the live version, the 2nd overwrites it.
const resetBtn = $("reset");
let resetArmed = false;
function disarmReset() {
  resetArmed = false;
  resetBtn.textContent = "Zet terug naar eerste versie";
}
resetBtn.addEventListener("click", async () => {
  const msg = $("reset-msg");
  if (!resetArmed) {
    const res = await fetch(page.path + "?v=" + Date.now(), { cache: "no-store" });
    if (!res.ok) {
      msg.className = "msg error";
      msg.textContent = "De huidige versie kon niet downloaden. Er gebeurde niets.";
      return;
    }
    saveBlob(await res.blob(), `${page.name}-backup-${new Date().toISOString().slice(0, 10)}.docx`);
    resetArmed = true;
    resetBtn.textContent = "Back-up gedownload. Klik nogmaals om terug te zetten";
    msg.className = "msg";
    msg.textContent = "Controleer of de back-up in je Downloads staat.";
    return;
  }
  disarmReset();
  await run(resetBtn, msg, "Bezig met terugzetten...", async () => {
    const res = await fetch(page.initial, { cache: "no-store" });
    if (!res.ok) throw new Error("De eerste versie is niet gevonden");
    await publish(new Uint8Array(await res.arrayBuffer()), `Reset ${page.name} to initial version`);
  });
});
