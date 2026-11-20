'use strict';
/*
 * Rules & notifications for spreadsheets: "when Sheet1!B2 > 100 then notify …". Stored per file in rules.json,
 * evaluated after every save / form append and on a schedule (only files whose mtime changed are read again).
 * A rule fires on the change from "not met" to "met" (no message flood while it stays met); 'changed' fires on every
 * change of the watched values.
 *
 *   rule = { id, name?, on: true, when: { sheet, cell | range, op, value, agg: 'any'|'all'|'sum'|'count'|'avg'|'min'|'max' },
 *            then: { notify: { channel: 'hostwerk'|'discord'|'ntfy'|'mail'|'all', text } }, state?: { met, hash, at, value } }
 *   validateRules(list) → normalised list (throws 400)
 *   evaluate(book, rule) → { met, value, hash }
 *   new Rules({ file, read(loc) → { buf, ext, mtime }, notify(rule, info) }) · .get(key) · .set(key, loc, list) · .check(key, { force })
 */
const crypto = require('crypto');
const { readJson, writeJson } = require('./store');
const S = require('./sheet');

const OPS = ['>', '<', '>=', '<=', '=', '!=', 'contains', 'notContains', 'empty', 'notEmpty', 'changed'];
const AGGS = ['any', 'all', 'sum', 'count', 'avg', 'min', 'max'];
const CHANNELS = ['hostwerk', 'discord', 'ntfy', 'mail', 'all'];
const fail = (msg, vars) => Object.assign(new Error(msg), { status: 400, vars });

function validateRules(list) {
  if (!Array.isArray(list)) throw fail('Invalid rules');
  if (list.length > 50) throw fail('At most 50 rules per file.');
  return list.map((r0, i) => {
    const r = r0 || {}; const w = r.when || {}; const n = (r.then && r.then.notify) || {};
    const cell = w.cell ? String(w.cell).toUpperCase().replace(/\$/g, '') : '';
    const range = w.range ? String(w.range).toUpperCase().replace(/\$/g, '') : '';
    if (cell && !S.parseRef(cell)) throw fail('Rule {n}: invalid cell', { n: i + 1 });
    if (range && !range.split(':').every((x) => S.parseRef(x))) throw fail('Rule {n}: invalid range', { n: i + 1 });
    if (!cell && !range) throw fail('Rule {n}: choose a cell or a range', { n: i + 1 });
    const op = String(w.op || '>');
    if (!OPS.includes(op)) throw fail('Rule {n}: invalid condition', { n: i + 1 });
    const agg = w.agg ? String(w.agg) : 'any';
    if (!AGGS.includes(agg)) throw fail('Rule {n}: invalid condition', { n: i + 1 });
    const channel = String(n.channel || 'hostwerk');
    if (!CHANNELS.includes(channel)) throw fail('Rule {n}: unknown channel', { n: i + 1 });
    return { id: /^[\w-]{1,40}$/.test(String(r.id || '')) ? String(r.id) : crypto.randomBytes(5).toString('hex'), name: String(r.name || '').slice(0, 120), on: r.on !== false,
      when: { sheet: String(w.sheet || '').slice(0, 120), ...(cell ? { cell } : { range }), op, value: w.value == null ? '' : String(w.value).slice(0, 500), agg },
      then: { notify: { channel, text: String(n.text || '').slice(0, 1000) } } };
  });
}

const num = (v) => (typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v.replace(',', '.'))) ? Number(v.replace(',', '.')) : null);
function test(v, op, want) {
  if (op === 'empty') return v === null || v === '';
  if (op === 'notEmpty') return !(v === null || v === '');
  if (op === 'contains') return String(v ?? '').toLowerCase().includes(String(want).toLowerCase());
  if (op === 'notContains') return !String(v ?? '').toLowerCase().includes(String(want).toLowerCase());
  const a = num(v); const b = num(want);
  if (a !== null && b !== null) return { '>': a > b, '<': a < b, '>=': a >= b, '<=': a <= b, '=': a === b, '!=': a !== b }[op];
  const x = String(v ?? '').toLowerCase(); const y = String(want).toLowerCase();
  return { '=': x === y, '!=': x !== y, '>': x > y, '<': x < y, '>=': x >= y, '<=': x <= y }[op] || false;
}
function evaluate(book, rule) {
  const w = rule.when;
  const values = w.cell ? [S.cellValue(book, w.sheet, w.cell)] : S.rangeValues(book, w.sheet, w.range);
  const hash = crypto.createHash('sha1').update(JSON.stringify(values)).digest('hex');
  if (w.op === 'changed') return { met: null, value: w.cell ? values[0] : values.length, hash };
  let value; let met;
  if (w.cell || w.agg === 'any' || w.agg === 'all') {
    const list = w.cell ? values : values.length ? values : [null];
    met = w.agg === 'all' ? list.every((v) => test(v, w.op, w.value)) : list.some((v) => test(v, w.op, w.value));
    value = w.cell ? values[0] : list.find((v) => test(v, w.op, w.value)) ?? null;
  } else {
    const nums = values.map(num).filter((x) => x !== null);
    value = { sum: nums.reduce((a, b) => a + b, 0), count: values.length, avg: nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : null, min: nums.length ? Math.min(...nums) : null, max: nums.length ? Math.max(...nums) : null }[w.agg];
    met = test(value, w.op, w.value);
  }
  return { met: !!met, value, hash };
}
const fill = (text, vars) => String(text || '').replace(/\{(\w+)\}/g, (m, k) => (vars[k] !== undefined && vars[k] !== null ? String(vars[k]) : m));

class Rules {
  constructor({ file, read, stat = null, notify, log = () => {} }) { this.file = file; this.read = read; this.stat = stat; this.notify = notify; this.log = log; this.data = readJson(file, { files: {} }); if (!this.data.files) this.data.files = {}; }
  save() { writeJson(this.file, this.data); }
  get(key) { const e = this.data.files[key]; return e ? e.rules.map(({ state, ...r }) => ({ ...r, last: state ? { met: state.met, at: state.at, value: state.value } : null })) : []; }
  set(key, loc, list) {
    const rules = validateRules(list);
    const old = this.data.files[key];
    for (const r of rules) { const o = old && old.rules.find((x) => x.id === r.id); if (o && o.state && JSON.stringify(o.when) === JSON.stringify(r.when)) r.state = o.state; }
    if (!rules.length) delete this.data.files[key]; else this.data.files[key] = { root: loc.root || null, path: loc.path, rules, mtime: null };
    this.save();
    if (rules.length) this.check(key, { force: true, silent: true }).catch(() => {});
    return this.get(key);
  }
  keys() { return Object.keys(this.data.files); }
  rekey(a, b) { let ch = false; for (const k of Object.keys(this.data.files)) if (k === a || k.startsWith(`${a}/`)) { const e = this.data.files[k]; delete this.data.files[k]; const nk = b + k.slice(a.length); if (!e.root) e.path = nk; this.data.files[nk] = e; ch = true; } if (ch) this.save(); }
  /** silent: only remember the state (first evaluation after creating a rule does not notify) */
  async check(key, { force = false, silent = false } = {}) {
    const e = this.data.files[key]; if (!e) return [];
    if (!force && this.stat) { try { const st = await this.stat({ root: e.root, path: e.path }); if (st && st.mtime === e.mtime) return []; } catch { return []; } }
    let f; try { f = await this.read({ root: e.root, path: e.path }); } catch { return []; }
    if (!force && e.mtime === f.mtime) return [];
    e.mtime = f.mtime;
    let book; try { book = S.readBook(f.buf, f.ext); } catch (x) { this.log('rules: cannot read', key, x.message); return []; }
    const fired = [];
    for (const r of e.rules) {
      if (!r.on) continue;
      const ev = evaluate(book, r);
      const prev = r.state || null;
      const now = Date.now();
      const fire = r.when.op === 'changed' ? !!prev && prev.hash !== ev.hash : ev.met && !(prev && prev.met);
      r.state = { met: ev.met, hash: ev.hash, at: fire ? now : (prev && prev.at) || null, value: ev.value };
      if (fire && !silent) {
        const vars = { value: ev.value, file: e.path.split('/').pop(), cell: r.when.cell || r.when.range, sheet: r.when.sheet || '', rule: r.name };
        fired.push({ rule: r, vars });
        try { await this.notify(r, vars, e); } catch (x) { this.log('rules: notify failed', x.message); }
      }
    }
    this.save();
    return fired;
  }
  async checkAll() { for (const k of this.keys()) { await this.check(k); await new Promise((ok) => setImmediate(ok)); } }
}

module.exports = { Rules, validateRules, evaluate, fill, test, OPS };
