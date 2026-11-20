'use strict';
/*
 * Search by meaning, with whatever AI the user set up in Hostwerk – never a fixed cloud:
 *   1. embeddings: if the local Ollama has an embedding model (nomic-embed-text, mxbai-embed-large, bge-*, all-minilm,
 *      snowflake-arctic-embed …), every library file gets up to 4 chunk vectors (int8, incremental by mtime, a few files
 *      per pass, never blocking) → cosine ranking;
 *   2. otherwise the configured chat AI turns the question into search words (synonyms, other languages) and the
 *      full-text index answers them, then the AI may re-rank the candidates;
 *   3. without any AI: the plain full-text search.
 *
 *   new Semantic({ file, ai: () => ctx.ai, index, log }) · .search(q, { limit, dir }) → { hits: [{ rel, score, snippet }], mode }
 *   .refresh(maxFiles) · .state()
 */
const { readJson, writeJson } = require('./store');

const EMBED = /(embed|bge-|all-minilm|e5-|gte-|arctic-embed|granite-embedding)/i;
const CHUNK = 1200; const CHUNKS = 4; const MAX_FILES = 4000;

function quantize(v) { let m = 0; for (const x of v) m = Math.max(m, Math.abs(x)); const s = m || 1; const q = Buffer.alloc(v.length); for (let i = 0; i < v.length; i++) q.writeInt8(Math.round((v[i] / s) * 127), i); return { s, q: q.toString('base64') }; }
function dequant({ s, q }) { const b = Buffer.from(q, 'base64'); const v = new Float32Array(b.length); for (let i = 0; i < b.length; i++) v[i] = (b.readInt8(i) / 127) * s; return v; }
function cosine(a, b) { let d = 0; let na = 0; let nb = 0; const n = Math.min(a.length, b.length); for (let i = 0; i < n; i++) { d += a[i] * b[i]; na += a[i] * a[i]; nb += b[i] * b[i]; } return na && nb ? d / Math.sqrt(na * nb) : 0; }
function chunks(text) { const t = String(text || '').replace(/\s+/g, ' ').trim(); if (!t) return []; const out = []; for (let i = 0; i < t.length && out.length < CHUNKS; i += CHUNK) out.push(t.slice(i, i + CHUNK)); return out; }

class Semantic {
  constructor({ file, ai, index, log = () => {} }) { this.file = file; this.ai = ai; this.index = index; this.log = log; this.data = readJson(file, { model: null, files: {} }); this.model = undefined; this.modelAt = 0; this.busy = null; this.vec = new Map(); }
  async embedModel() {
    if (this.model !== undefined && Date.now() - this.modelAt < 10 * 60000) return this.model;
    this.modelAt = Date.now(); this.model = null;
    const ai = this.ai();
    if (!ai || typeof ai.api !== 'function') return null;
    try {
      const j = await (await ai.api('/api/tags', { timeout: 3000 })).json();
      const m = (j.models || []).find((x) => EMBED.test(x.name) || (Array.isArray(x.capabilities) && x.capabilities.includes('embedding')));
      this.model = m ? m.name : null;
    } catch { this.model = null; }
    return this.model;
  }
  async embed(model, input) {
    const r = await this.ai().api('/api/embed', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ model, input }), timeout: 120000 });
    const j = await r.json();
    if (!Array.isArray(j.embeddings)) throw new Error('no embeddings');
    return j.embeddings;
  }
  state() { return { model: this.data.model, files: Object.keys(this.data.files).length, busy: !!this.busy }; }
  /** index a few changed files (called from a timer); returns the number done */
  refresh(maxFiles = 15) {
    if (this.busy) return this.busy;
    this.busy = (async () => {
      const model = await this.embedModel();
      if (!model) return 0;
      if (this.data.model !== model) { this.data = { model, files: {} }; this.vec.clear(); }
      const all = this.index.data.files || {};
      let n = 0;
      for (const [rel, rec] of Object.entries(all)) {
        if (n >= maxFiles) break;
        const cur = this.data.files[rel];
        if (cur && cur.mtime === rec.mtime) continue;
        if (Object.keys(this.data.files).length >= MAX_FILES && !cur) break;
        const parts = chunks(`${rel.split('/').pop()}\n${rec.text || ''}`);
        if (!parts.length) { this.data.files[rel] = { mtime: rec.mtime, v: [] }; continue; }
        try { const e = await this.embed(model, parts); this.data.files[rel] = { mtime: rec.mtime, v: e.map(quantize) }; this.vec.delete(rel); n++; } catch (x) { this.log('semantic embed failed', x.message); break; }
        await new Promise((ok) => setImmediate(ok));
      }
      for (const k of Object.keys(this.data.files)) if (!all[k]) { delete this.data.files[k]; this.vec.delete(k); }
      if (n) writeJson(this.file, this.data);
      return n;
    })().catch(() => 0).finally(() => { this.busy = null; });
    return this.busy;
  }
  vectors(rel) { let v = this.vec.get(rel); if (!v) { const f = this.data.files[rel]; v = f ? f.v.map(dequant) : []; this.vec.set(rel, v); } return v; }

  async search(q, { limit = 30, dir = '' } = {}) {
    const text = String(q || '').trim().slice(0, 500);
    if (!text) return { hits: [], mode: 'none' };
    const inDir = (rel) => !dir || rel === dir || rel.startsWith(`${dir}/`);
    const model = await this.embedModel().catch(() => null);
    if (model && this.data.model === model && Object.keys(this.data.files).length) {
      try {
        const [qv] = await this.embed(model, [text]);
        const out = [];
        for (const rel of Object.keys(this.data.files)) {
          if (!inDir(rel)) continue;
          let best = 0; for (const v of this.vectors(rel)) best = Math.max(best, cosine(qv, v));
          if (best > 0.3) out.push({ rel, score: Math.round(best * 1000) / 1000 });
        }
        out.sort((a, b) => b.score - a.score);
        const X = require('./extract');
        return { hits: out.slice(0, limit).map((h) => { const rec = this.index.get(h.rel); return { ...h, snippet: rec ? X.snippet(rec.text, text) : '' }; }), mode: 'embeddings', model };
      } catch (x) { this.log('semantic query failed', x.message); }
    }
    // AI-assisted: question → search words → full text
    const ai = this.ai();
    let words = null;
    try {
      if (ai && ai.chats && typeof ai.chats.oneShot === 'function' && (!ai.readyState || (await ai.readyState()).ready !== false)) {
        const out = await ai.chats.oneShot([
          { role: 'system', content: 'You turn a search question into search words for a full-text search over office documents. Answer with ONE line: 8 to 15 single words or short terms separated by commas – synonyms, related terms, German and English variants. No explanations.' },
          { role: 'user', content: text },
        ], { timeout: 20000, temperature: 0, fast: true, maxTokens: 120 });
        words = String(out || '').split(/[,;\n]/).map((w) => w.trim().toLowerCase()).filter((w) => w && w.length < 40).slice(0, 15);
      }
    } catch (x) { this.log('semantic: AI not available', x.message); }
    const terms = [...new Set([...(words || []), ...text.toLowerCase().split(/\s+/).filter((w) => w.length > 2)])];
    const score = new Map();
    for (const t of terms) for (const h of this.index.search(t, { dir, limit: 100 })) { const c = score.get(h.rel) || { rel: h.rel, score: 0, snippet: h.snippet }; c.score += h.score; score.set(h.rel, c); }
    const hits = [...score.values()].sort((a, b) => b.score - a.score).slice(0, limit);
    return { hits, mode: words ? 'ai-terms' : 'fulltext', terms };
  }
}

module.exports = { Semantic, quantize, dequant, cosine, chunks };
