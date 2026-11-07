import assert from "node:assert/strict";
import test from "node:test";

import { leanMachine, machine } from "../server/pptx.mjs";

test("the export knows how big its machine is, and PPTX_LEAN decides by hand", () => {
  const { cpus, memory } = machine();
  assert.ok(cpus > 0 && Number.isFinite(cpus), String(cpus));
  assert.ok(memory > 64 * 1024 ** 2, String(memory));
  const before = process.env.PPTX_LEAN;
  try {
    process.env.PPTX_LEAN = "1";
    assert.equal(leanMachine(), true);
    process.env.PPTX_LEAN = "0";
    assert.equal(leanMachine(), false);
    delete process.env.PPTX_LEAN;
    assert.equal(leanMachine(), cpus < 1 || memory < 1.5 * 1024 ** 3);
  } finally {
    if (before === undefined) delete process.env.PPTX_LEAN; else process.env.PPTX_LEAN = before;
  }
});
