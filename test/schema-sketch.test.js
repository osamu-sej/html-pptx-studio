import assert from "node:assert/strict";
import test from "node:test";

import { codexDeckSchema, codexSlideSchema } from "../server/schemas.mjs";
import { schemaSketch } from "../server/schema-sketch.mjs";

test("the deck format written out short names every slide type and its fields", () => {
  const sketch = schemaSketch(codexDeckSchema);
  // Far shorter than the schema itself (which is sent as the output format).
  assert.ok(sketch.length < JSON.stringify(codexDeckSchema).length / 10, `${sketch.length}`);
  for (const variant of codexDeckSchema.properties.slideData.items.anyOf) {
    assert.match(sketch, new RegExp(`type: "${variant.properties.type.const}"`));
  }
  assert.match(sketch, /^\{deckTitle: string, purpose: string, audience: string, slideData: SlideDataItem\[\]\}/);
  // Nullable fields are optional; required ones are not; long enums used again and again get a name.
  assert.match(sketch, /\{type: "compare", leftTitle: string, rightTitle: string, leftItems: string\[\], rightItems: string\[\]\}/);
  assert.match(sketch, /twoColumn\?: boolean/);
  assert.match(sketch, /icon\?: Icon/);
  assert.match(sketch, /^Icon = "bulb"\|/m);
  assert.doesNotMatch(sketch, /\|null|,null/);
});

test("a schema without tagged variants is written as one type", () => {
  const sketch = schemaSketch({ type: "object", properties: { slide: codexSlideSchema.properties.slide, n: { type: "integer" } } });
  assert.match(sketch, /n: int/);
  assert.match(sketch, /"title"/);
});
