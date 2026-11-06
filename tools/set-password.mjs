#!/usr/bin/env node
// Sets or changes the admin password. Writes admin-key.json: the GitHub token, encrypted with
// a key derived from the password. Commit and push that file to activate the new password.
//
//   node tools/set-password.mjs
//
// Token: a fine-grained GitHub token for this repo only, permission "Contents: Read and write".
import { createInterface } from "node:readline/promises";
import { existsSync, readFileSync, writeFileSync } from "node:fs";

const KEY_FILE = new URL("../admin-key.json", import.meta.url);
const REPO = "MergaertR/dakloosheid-syllabus";
const ITERATIONS = 600_000;
const { subtle } = globalThis.crypto;

async function ask(question, hidden = true) {
  const rl = createInterface({ input: process.stdin, output: process.stdout, terminal: true });
  if (hidden) rl._writeToOutput = (s) => { if (s.includes(question)) rl.output.write(question); };
  const answer = await rl.question(question);
  rl.close();
  if (hidden) process.stdout.write("\n");
  return answer.trim();
}

async function deriveKey(password, salt, iterations, usage) {
  const material = await subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveKey"]);
  return subtle.deriveKey({ name: "PBKDF2", hash: "SHA-256", salt, iterations }, material, { name: "AES-GCM", length: 256 }, false, [usage]);
}

async function decryptExisting(password) {
  const box = JSON.parse(readFileSync(KEY_FILE, "utf8"));
  const key = await deriveKey(password, Buffer.from(box.salt, "base64"), box.iterations, "decrypt");
  const plain = await subtle.decrypt({ name: "AES-GCM", iv: Buffer.from(box.iv, "base64") }, key, Buffer.from(box.data, "base64"));
  return new TextDecoder().decode(plain);
}

const fail = (msg) => { console.error(msg); process.exit(1); };

let token = await ask("GitHub token (Enter = keep the current token): ");
if (!token) {
  if (!existsSync(KEY_FILE)) fail("There is no admin-key.json yet, so a token is required.");
  const current = await ask("Current password: ");
  token = await decryptExisting(current).catch(() => fail("Wrong current password."));
}

const check = await fetch(`https://api.github.com/repos/${REPO}`, { headers: { Authorization: `Bearer ${token}` } });
if (!check.ok) fail(`GitHub rejected the token (HTTP ${check.status}).`);
const perms = (await check.json()).permissions;
if (perms && !perms.push) fail("The token cannot write to this repo. Give it 'Contents: Read and write'.");

const password = await ask("New password: ");
if (!password) fail("The password cannot be empty.");
if (password !== await ask("Repeat new password: ")) fail("The passwords do not match.");

const salt = crypto.getRandomValues(new Uint8Array(16));
const iv = crypto.getRandomValues(new Uint8Array(12));
const key = await deriveKey(password, salt, ITERATIONS, "encrypt");
const data = new Uint8Array(await subtle.encrypt({ name: "AES-GCM", iv }, key, new TextEncoder().encode(token)));
const b64 = (bytes) => Buffer.from(bytes).toString("base64");

writeFileSync(KEY_FILE, JSON.stringify({ iterations: ITERATIONS, salt: b64(salt), iv: b64(iv), data: b64(data) }, null, 2) + "\n");
console.log("Wrote admin-key.json. Commit and push it to activate the new password.");
