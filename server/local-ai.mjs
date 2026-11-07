// The stand-by AI: a model on this machine (Gemma through Ollama), used only when Codex cannot answer
// (usage limit reached, logged out, not starting, failing). The web deployment has no local model, so
// it says so plainly instead of pretending there is a fallback.
//
//   LOCAL_AI_URL    Ollama's address (default http://127.0.0.1:11434)
//   LOCAL_AI_MODEL  model tag (default gemma4:12b)
//   LOCAL_AI=off    never use a local model

const STATUS_TTL_MS = 20_000;

export class LocalModel {
  constructor({ url = process.env.LOCAL_AI_URL || "http://127.0.0.1:11434", model = process.env.LOCAL_AI_MODEL || "gemma4:12b",
    disabled = process.env.LOCAL_AI === "off", web = Boolean(process.env.RENDER) } = {}) {
    this.url = url.replace(/\/+$/, "");
    this.model = model;
    this.disabled = disabled;
    this.web = web;
    this.cached = null;
  }

  /** { available, model, reason, web } — probed at most every 20 seconds. */
  async status({ fresh = false } = {}) {
    if (this.disabled) return { available: false, model: this.model, web: this.web, reason: "予備AI（Gemma）は設定で無効になっています。" };
    if (!fresh && this.cached && Date.now() - this.cached.at < STATUS_TTL_MS) return this.cached.value;
    let value;
    try {
      const response = await fetch(`${this.url}/api/tags`, { signal: AbortSignal.timeout(2500) });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const names = ((await response.json()).models ?? []).map((m) => m.name ?? m.model).filter(Boolean);
      const family = this.model.split(":")[0];
      const model = names.find((name) => name === this.model || name === `${this.model}:latest`) ?? names.find((name) => name.startsWith(`${family}:`));
      value = model
        ? { available: true, model, web: this.web, reason: "" }
        : { available: false, model: this.model, web: this.web, reason: `Ollamaに ${this.model} が入っていません（ollama pull ${this.model}）。` };
    } catch {
      value = {
        available: false, model: this.model, web: this.web,
        reason: this.web
          ? "Web版では手元のPCのGemmaに接続できないため、予備AIは使えません。Codexが使えないときはAI機能が止まります。"
          : `Gemma（Ollama）に接続できません（${this.url}）。Ollamaを起動すると予備AIとして使えます。`,
      };
    }
    this.cached = { at: Date.now(), value };
    return value;
  }

  /**
   * One chat turn with structured output. `messages` is the whole conversation so far; `onText(text)`
   * receives the answer as it grows. Resolves with the final text.
   */
  async chat(messages, schema, onText = () => {}) {
    const { available, model, reason } = await this.status();
    if (!available) throw new Error(reason);
    const response = await fetch(`${this.url}/api/chat`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      // Long prompts (layout guide + material) need more than Ollama's small default context.
      body: JSON.stringify({ model, messages, format: schema, stream: true, options: { temperature: 0.3, num_ctx: 32768 } }),
      signal: AbortSignal.timeout(15 * 60 * 1000),
    });
    if (!response.ok || !response.body) throw new Error(`Gemmaが応答しませんでした（HTTP ${response.status}）。`);
    let text = "";
    let buffer = "";
    const decoder = new TextDecoder();
    for await (const chunk of response.body) {
      buffer += decoder.decode(chunk, { stream: true });
      let newline;
      while ((newline = buffer.indexOf("\n")) >= 0) {
        const line = buffer.slice(0, newline).trim();
        buffer = buffer.slice(newline + 1);
        if (!line) continue;
        const part = JSON.parse(line);
        if (part.error) throw new Error(`Gemmaでエラーが発生しました：${part.error}`);
        if (part.message?.content) {
          text += part.message.content;
          onText(text);
        }
      }
    }
    if (!text.trim()) throw new Error("Gemmaから回答がありませんでした。");
    return text;
  }
}
