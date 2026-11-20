'use strict';
/*
 * Live fields for documents and slides: date/time, user name (only when the user shares it), Hostwerk values (through
 * the data sources), the value of a cell in another office file, document properties.
 *
 *   new Fields({ t, data (DataSources), userName(), readFile(loc) → { buf, ext, mtime, key }, info(buf, ext) })
 *   .catalog() → [{ id, label, group, params: [{ id, label, type, options? }] }]
 *   .value(id, params) → { value, text, updatedAt, ttl }
 */
const S = require('./sheet');

class Fields {
  constructor({ t = (s) => s, data, userName = () => null, readFile, info }) { this.t = t; this.data = data; this.userName = userName; this.readFile = readFile; this.info = info; this.cache = new Map(); }
  catalog() {
    const t = this.t;
    const file = [{ id: 'path', label: t('File'), type: 'file' }];
    return [
      { id: 'date', label: t('Date'), group: 'time', params: [{ id: 'format', label: t('Format'), type: 'choice', options: ['short', 'long', 'iso'] }] },
      { id: 'time', label: t('Time'), group: 'time', params: [] },
      { id: 'datetime', label: t('Date and time'), group: 'time', params: [] },
      { id: 'user', label: t('User name'), group: 'document', params: [] },
      { id: 'service', label: t('State of a service'), group: 'hostwerk', params: [{ id: 'service', label: t('Service'), type: 'service' }] },
      { id: 'players', label: t('Number of players online'), group: 'hostwerk', params: [{ id: 'server', label: t('Server (empty = all)'), type: 'minecraft' }] },
      { id: 'system', label: t('System load'), group: 'hostwerk', params: [{ id: 'metric', label: t('Value'), type: 'choice', options: ['cpu', 'mem', 'diskFree', 'servicesRunning'] }] },
      { id: 'hostwerk', label: t('Hostwerk'), group: 'hostwerk', params: [{ id: 'metric', label: t('Value'), type: 'choice', options: ['version', 'uptimeHours'] }] },
      { id: 'cell', label: t('Cell of another spreadsheet'), group: 'files', params: [...file, { id: 'sheet', label: t('Sheet'), type: 'text' }, { id: 'cell', label: t('Cell'), type: 'text' }] },
      { id: 'docprop', label: t('Document property'), group: 'files', params: [...file, { id: 'prop', label: t('Property'), type: 'choice', options: ['title', 'author', 'created', 'modified', 'pages', 'words', 'slides', 'sheets', 'name', 'size'] }] },
    ];
  }
  async value(id, p = {}) {
    const now = new Date();
    const lang = p.lang === 'en' ? 'en-US' : 'de-DE';
    const done = (value, ttl = 60, text) => ({ value, text: text != null ? text : value == null ? '' : String(value), updatedAt: now.toISOString(), ttl });
    switch (id) {
      case 'date': return done(now.toISOString().slice(0, 10), 60, p.format === 'iso' ? now.toISOString().slice(0, 10) : now.toLocaleDateString(lang, p.format === 'long' ? { day: 'numeric', month: 'long', year: 'numeric' } : undefined));
      case 'time': return done(now.toISOString(), 30, now.toLocaleTimeString(lang, { hour: '2-digit', minute: '2-digit' }));
      case 'datetime': return done(now.toISOString(), 30, now.toLocaleString(lang, { dateStyle: 'medium', timeStyle: 'short' }));
      case 'user': { const n = this.userName(); return done(n || null, 3600, n || ''); }
      case 'service': return done((await this.data.query('service', { service: String(p.service || '') })).value ?? null, 10);
      case 'players': return done((await this.data.query('playerCount', { server: String(p.server || '') })).value ?? 0, 10);
      case 'system': return done((await this.data.query('system', { metric: String(p.metric || 'cpu') })).value ?? null, 10);
      case 'hostwerk': return done((await this.data.query('hostwerk', { metric: String(p.metric || 'version') })).value ?? null, 60);
      case 'cell': {
        const f = await this.readFile({ root: p.root || null, path: String(p.path || '') });
        const ck = `${f.key}\n${f.mtime}`;
        let book = this.cache.get(ck);
        if (!book) { book = S.readBook(f.buf, f.ext); this.cache.set(ck, book); if (this.cache.size > 30) this.cache.delete(this.cache.keys().next().value); }
        if (!S.parseRef(String(p.cell || ''))) throw Object.assign(new Error('Invalid cell'), { status: 400 });
        return done(S.cellValue(book, p.sheet || '', String(p.cell).toUpperCase().replace(/\$/g, '')), 30);
      }
      case 'docprop': {
        const f = await this.readFile({ root: p.root || null, path: String(p.path || '') });
        const prop = String(p.prop || 'title');
        if (prop === 'name') return done(String(p.path || '').split('/').pop(), 300);
        if (prop === 'size') return done(f.buf.length, 300);
        const i = await this.info(f.buf, f.ext);
        return done(i && i[prop] != null ? i[prop] : null, 300);
      }
      default: throw Object.assign(new Error('Unknown field'), { status: 404 });
    }
  }
}

module.exports = { Fields };
