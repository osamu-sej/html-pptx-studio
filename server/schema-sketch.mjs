// A short, readable sketch of a JSON schema (TypeScript-like), for asking the AI for JSON without
// sending the schema itself as the turn's output format (see codex-app-server.mjs: a deck answer that
// was cut off at the output limit is asked again without the strict format).
//
// Nullable fields are optional (`name?:`), long enums used more than once get a name (`Icon = "bulb"|…`),
// and the fields most slide types share are written once.

const nullable = (schema) => (Array.isArray(schema?.anyOf) && schema.anyOf.some((item) => item.type === "null"))
  || (Array.isArray(schema?.type) && schema.type.includes("null"));

export function schemaSketch(schema) {
  const enumNames = new Map();
  const enumUses = new Map();
  // First pass: count the enums, so the long ones used more than once can be named.
  (function walk(node, key) {
    if (!node || typeof node !== "object") return;
    if (Array.isArray(node.enum)) {
      const text = node.enum.map((value) => JSON.stringify(value)).join("|");
      const use = enumUses.get(text) ?? { count: 0, key };
      use.count += 1;
      enumUses.set(text, use);
    }
    for (const [childKey, child] of Object.entries(node.properties ?? {})) walk(child, childKey);
    for (const child of node.anyOf ?? []) walk(child, key);
    if (node.items) walk(node.items, key);
  })(schema, "value");
  const taken = new Set();
  for (const [text, use] of enumUses) {
    if (use.count < 2 || text.length < 60) continue;
    let name = use.key.replace(/^./, (c) => c.toUpperCase()).replace(/s$/, "");
    while (taken.has(name)) name += "_";
    taken.add(name);
    enumNames.set(text, name);
  }

  // Unions of tagged objects (a slide is one of 44 kinds): written once below, by name.
  const tagged = (node) => Array.isArray(node?.anyOf) && node.anyOf.filter((item) => item.type !== "null").length > 2
    && node.anyOf.filter((item) => item.type !== "null").every((item) => item.properties?.type?.const !== undefined);
  const unions = new Map();
  const unionName = (node, key) => {
    const text = JSON.stringify(node.anyOf);
    if (!unions.has(text)) {
      let name = `${key.replace(/^./, (c) => c.toUpperCase())}Item`;
      while ([...unions.values()].some((union) => union.name === name)) name += "_";
      unions.set(text, { name, items: node.anyOf.filter((item) => item.type !== "null") });
    }
    return unions.get(text).name;
  };

  const type = (node, key = "value") => {
    if (!node || typeof node !== "object") return "any";
    if (Array.isArray(node.type)) return node.type.filter((name) => name !== "null").map((name) => type({ ...node, type: name }, key)).join("|") || "null";
    if (tagged(node)) return unionName(node, key);
    if (node.anyOf) return node.anyOf.filter((item) => item.type !== "null").map((item) => type(item, key)).join("|") || "null";
    if (node.const !== undefined) return JSON.stringify(node.const);
    if (Array.isArray(node.enum)) {
      const text = node.enum.map((value) => JSON.stringify(value)).join("|");
      return enumNames.get(text) ?? text;
    }
    if (node.type === "array") {
      const inner = type(node.items, key);
      return `${inner.includes("|") ? `(${inner})` : inner}[]`;
    }
    if (node.type === "object") return `{${fields(node).join(", ")}}`;
    return node.type === "integer" ? "int" : node.type ?? "any";
  };
  const field = (key, node) => `${key}${nullable(node) ? "?" : ""}: ${type(node, key)}`;
  const fields = (node) => Object.entries(node.properties ?? {}).map(([key, child]) => field(key, child));

  const lines = [type(schema)];
  // Each union: the fields most of its kinds share once, then each kind's own fields.
  for (const { name, items } of unions.values()) {
    const counts = new Map();
    for (const item of items) for (const text of fields(item)) counts.set(text, (counts.get(text) ?? 0) + 1);
    const common = new Set([...counts].filter(([text, count]) => count >= items.length / 2 && !text.startsWith("type:")).map(([text]) => text));
    lines.push("", `${name} は次のどれか（type で種類を決める）。多くの種類に共通の項目: {${[...common].join(", ")}}`);
    for (const item of items) lines.push(`- {${fields(item).filter((text) => !common.has(text)).join(", ")}}`);
  }
  if (enumNames.size) {
    lines.push("", "名前を付けた選択肢:");
    for (const [text, name] of enumNames) lines.push(`${name} = ${text}`);
  }
  return lines.join("\n");
}
