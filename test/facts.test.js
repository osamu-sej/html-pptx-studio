import assert from "node:assert/strict";
import test from "node:test";

import { numbersIn, unsourcedNumbers } from "../server/facts.mjs";

test("figures are compared by value, whatever their spelling", () => {
  assert.deepEqual([...numbersIn("１，２００店で42.0%、2026年")], ["1200", "42", "2026"]);
  assert.deepEqual([...numbersIn("3.50倍")], ["3.5"]);
});

test("only figures the material lacks are reported; counts, dates and layout settings are not", () => {
  const slides = [
    { type: "title", title: "報告", date: "2026年9月" },
    { type: "kpi", title: "3つの成果", items: [{ label: "作業", value: "-120h/月" }, { label: "発注", value: "43%" }], notes: "STEP 2で説明" },
    { type: "gantt", title: "計画", periods: ["4月", "5月"], items: [{ title: "設計", start: 0, span: 2 }, { title: "開発", start: 1, span: 1 }] },
    { type: "imageText", title: "推移", image: { chartType: "line", data: { items: [{ label: "4月", value: 95 }, { label: "5月", value: 77 }] } } },
  ];
  const issues = unsourcedNumbers(slides, "1店あたり月120時間の削減。導入前95分");
  assert.deepEqual(issues.map((issue) => [issue.slide, issue.field, issue.values]), [
    [1, "items[1].value", ["43%"]],
    [3, "image.data.items[1].value", ["77"]],
  ]);
  assert.deepEqual(unsourcedNumbers(slides, "", { only: new Set([0, 2]) }), []);
});
