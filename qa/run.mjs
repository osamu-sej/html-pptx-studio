// Runs the browser QA scripts the way CI does: starts the studio server on its own port, runs each script against it,
// retries a failed script once (a second failure is real; a pass on the retry is reported as flaky), and prints a table.
// A script passes only when it exits 0, prints no "errors:" block and ends with "no errors" — so a script that forgets
// to set its exit code cannot hide a failure.
// Usage: node qa/run.mjs [--only=studio-smoke,studio-format,layouts] [--retry=1] [--port=8791] [--base=http://127.0.0.1:8787] [--list]
//   no --only            every script, in the order of CLAUDE.md, then the layout check (every layout in every theme)
//   --base               use a server that is already running instead of starting one
import { spawn } from "node:child_process";
import { readdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { join } from "node:path";

const root = fileURLToPath(new URL("..", import.meta.url));
const args = Object.fromEntries(process.argv.slice(2).map((arg) => arg.replace(/^--/, "").split("=")));
const retries = Math.max(0, Number(args.retry ?? 1));
const port = Number(args.port || 8791);
const perScript = Number(args.timeout || 900) * 1000;
const inCI = process.env.GITHUB_ACTIONS === "true";

/** The layout check shoots every layout in every theme; --strict makes a fit note (text that does not fit) fail the run. */
const LAYOUTS = [
  ["--themes=clarity,midnight,editorial,mono,forest,sunset,aurora,kinari", "--deck=/qa/all-layouts.json", "--per=8", "--strict"],
];

const found = (await readdir(join(root, "qa"))).filter((f) => /^studio-.*\.mjs$/.test(f)).map((f) => f.replace(/\.mjs$/, "")).sort();
const ORDER = ["studio-smoke", "studio-editing", "studio-ai-mock", "studio-motion", "studio-drill", "studio-interactive", "studio-objects", "studio-animations", "studio-tables", "studio-convert", "studio-import", "studio-interact", "studio-layout", "studio-motion-ide", "studio-powerpoint", "studio-media", "studio-graphics", "studio-coedit", "studio-draw", "studio-show", "studio-review", "studio-more", "studio-format", "studio-present", "studio-rehearse", "studio-3d", "studio-live", "studio-tools"];
// A script added without joining ORDER still runs (after the known ones) so it cannot be forgotten.
const all = [...ORDER.filter((n) => found.includes(n)), ...found.filter((n) => !ORDER.includes(n)), "layouts"];
if ("list" in args) { console.log(all.join("\n")); process.exit(0); }

const wanted = args.only ? String(args.only).split(",").map((n) => n.trim()).filter(Boolean) : all;
for (const name of wanted) if (!all.includes(name)) { console.error(`unknown script: ${name} (known: ${all.join(", ")})`); process.exit(2); }

const log = (line) => console.log(line);
const group = (title) => { if (inCI) log(`::group::${title}`); else log(`\n── ${title}`); };
const endGroup = () => { if (inCI) log("::endgroup::"); };

function run(command, commandArgs, { env = {} } = {}) {
  return new Promise((resolve) => {
    const child = spawn(command, commandArgs, { cwd: root, env: { ...process.env, ...env }, stdio: ["ignore", "pipe", "pipe"] });
    let out = "";
    child.stdout.on("data", (d) => { out += d; });
    child.stderr.on("data", (d) => { out += d; });
    const timer = setTimeout(() => { out += `\ntimeout after ${perScript / 1000}s\n`; child.kill("SIGKILL"); }, perScript);
    child.on("close", (code, signal) => { clearTimeout(timer); resolve({ code: code ?? (signal ? 137 : 1), out }); });
  });
}

/** A run is good only when the exit code, the "errors:" block and the closing line all agree. */
function verdict(name, { code, out }) {
  const lines = out.trim().split("\n");
  if (code !== 0) return `exit ${code}`;
  if (/^errors:/m.test(out)) return "reported errors";
  if (name === "layouts") {
    if (/^fit notes:$/m.test(out)) return "fit findings";
    return out.includes("no fit notes") ? null : "no layout-check result";
  }
  return /no errors\s*$/.test(lines[lines.length - 1] || "") ? null : `unexpected last line: ${(lines[lines.length - 1] || "").slice(0, 120)}`;
}

async function startServer() {
  if (args.base) return { base: args.base, stop: () => {} };
  const child = spawn(process.execPath, ["server.js"], { cwd: root, env: { ...process.env, PORT: String(port), HOST: "127.0.0.1" }, stdio: ["ignore", "pipe", "pipe"] });
  let out = "";
  child.stdout.on("data", (d) => { out += d; });
  child.stderr.on("data", (d) => { out += d; });
  const base = `http://127.0.0.1:${port}`;
  for (let i = 0; i < 80; i++) {
    try { if ((await fetch(`${base}/healthz`)).ok) return { base, stop: () => child.kill("SIGTERM"), output: () => out }; } catch { /* not up yet */ }
    if (child.exitCode !== null) break;
    await new Promise((r) => setTimeout(r, 250));
  }
  child.kill("SIGKILL");
  console.error(`the server did not start on ${base}\n${out}`);
  process.exit(2);
}

const server = await startServer();
log(`server: ${server.base}  scripts: ${wanted.length}  retries: ${retries}`);
const rows = [];
for (const name of wanted) {
  const attempts = [];
  let reason = "failed";
  for (let attempt = 0; attempt <= retries; attempt++) {
    const started = Date.now();
    let result;
    if (name === "layouts") {
      // Each shoot failing fails the check.
      let out = "";
      let code = 0;
      for (const shootArgs of LAYOUTS) {
        const r = await run(process.execPath, ["qa/shoot.mjs", ...shootArgs]);
        out += r.out;
        if (r.code !== 0) code = r.code;
      }
      result = { code, out };
    } else {
      result = await run(process.execPath, [`qa/${name}.mjs`, `--base=${server.base}`], { env: { PYTHON_BIN: process.env.PYTHON_BIN || "python3" } });
    }
    const bad = verdict(name, result);
    const seconds = Math.round((Date.now() - started) / 1000);
    attempts.push({ bad, seconds });
    group(`${name} — attempt ${attempt + 1}: ${bad ? `FAILED (${bad})` : "ok"} in ${seconds}s`);
    // Passing output is long and uninteresting; a failure shows everything the script said.
    if (bad) log(result.out.trim()); else log(result.out.trim().split("\n").slice(-3).join("\n"));
    endGroup();
    if (!bad) break;
    reason = bad;
  }
  const last = attempts[attempts.length - 1];
  const status = !last.bad ? (attempts.length > 1 ? "flaky" : "ok") : "FAIL";
  if (status === "flaky") log(`${inCI ? "::warning::" : "!! "}${name} failed once (${attempts[0].bad}) and passed on the retry — look for a timing-dependent step`);
  if (status === "FAIL") log(`${inCI ? "::error::" : "!! "}${name} failed: ${reason}`);
  rows.push({ name, status, seconds: attempts.reduce((a, b) => a + b.seconds, 0), tries: attempts.length });
}
server.stop();

log("\nscript                      result  tries  seconds");
for (const r of rows) log(`${r.name.padEnd(27)} ${r.status.padEnd(7)} ${String(r.tries).padEnd(6)} ${r.seconds}`);
const failed = rows.filter((r) => r.status === "FAIL");
const flaky = rows.filter((r) => r.status === "flaky");
log(`\n${rows.length - failed.length}/${rows.length} passed${flaky.length ? `, ${flaky.length} flaky (${flaky.map((r) => r.name).join(", ")})` : ""}${failed.length ? `, FAILED: ${failed.map((r) => r.name).join(", ")}` : ""}`);
if (process.env.GITHUB_STEP_SUMMARY) {
  const { appendFileSync } = await import("node:fs");
  appendFileSync(process.env.GITHUB_STEP_SUMMARY, `| script | result | tries | seconds |\n|---|---|---|---|\n${rows.map((r) => `| ${r.name} | ${r.status} | ${r.tries} | ${r.seconds} |`).join("\n")}\n`);
}
process.exit(failed.length ? 1 : 0);
