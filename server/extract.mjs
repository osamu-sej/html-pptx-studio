// Source material from files: text for the AI (extract_text.py) and editable decks (import_deck.py).
import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { join } from "node:path";

export function pythonBinary(rootDir) {
  if (process.env.PYTHON_BIN) return process.env.PYTHON_BIN;
  const venv = join(rootDir, ".venv", "bin", "python");
  return existsSync(venv) ? venv : "python3";
}

function runScript(script, args, buffer, { rootDir, timeoutMs, timeoutMessage }) {
  return new Promise((resolve, reject) => {
    const child = spawn(pythonBinary(rootDir), [join(rootDir, "tools", script), ...args], { cwd: rootDir, stdio: ["pipe", "pipe", "pipe"] });
    const out = [];
    let stderr = "";
    const timer = setTimeout(() => { child.kill("SIGKILL"); reject(new Error(timeoutMessage)); }, timeoutMs);
    child.stdout.on("data", (chunk) => out.push(chunk));
    child.stderr.on("data", (chunk) => { stderr += String(chunk); });
    child.once("error", (error) => { clearTimeout(timer); reject(error); });
    child.once("exit", (code) => {
      clearTimeout(timer);
      if (code !== 0) return reject(new Error((stderr.trim().split("\n").at(-1) || "読み込めませんでした").replace(/^\w*Error:\s*/, "")));
      try { resolve(JSON.parse(Buffer.concat(out).toString("utf8"))); }
      catch { reject(new Error("読み込み結果を解析できませんでした。")); }
    });
    child.stdin.end(buffer);
  });
}

export function extractText(buffer, fileName, { rootDir, timeoutMs = 60_000 } = {}) {
  return runScript("extract_text.py", [fileName], buffer, { rootDir, timeoutMs, timeoutMessage: "読み込みがタイムアウトしました。" });
}

export function importDeck(buffer, fileName, { rootDir, timeoutMs = 120_000, mode = "layout" } = {}) {
  return runScript("import_deck.py", [fileName, mode], buffer, { rootDir, timeoutMs, timeoutMessage: "取り込みがタイムアウトしました。" });
}

/** An Excel workbook's sheets as table cells (formulas, number formats, merged cells): { sheets: [{ name, rows, cols }] }. */
export function importSheets(buffer, fileName, { rootDir, timeoutMs = 60_000 } = {}) {
  return runScript("xlsx_table.py", [fileName], buffer, { rootDir, timeoutMs, timeoutMessage: "Excelの読み込みがタイムアウトしました。" });
}
