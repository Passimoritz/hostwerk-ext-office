'use strict';
/*
 * Live data from Hostwerk for spreadsheets/documents (read only): services, resources, Minecraft, backups, tasks,
 * websites, databases, certificates, uptime. Read through Hostwerk's own API (host.get → loopback, same data as the
 * UI) – only chosen columns leave this module, strings are additionally redacted (URL credentials, token-like values).
 *
 *   new DataSources({ get, t, now })
 *   .sources() → [{ id, label, kind: 'table'|'value', params: [{ id, label, options? }] }]
 *   .query(id, params) → { columns: [{ id, label, type }], rows: [[…]], value?, updatedAt }
 *   .watch(bus, broadcast) → off()  – pushes { type: 'office-data', source, updatedAt } after core changes
 *     (only for sources queried in the last 15 minutes, at most every 3 s per source)
 */
const MASK = '••••••';
const SECRET_VALUE = /^(sk-|ghp_|gho_|github_pat_|xox[abp]-|AKIA|AIza)[\w-]{8,}|-----BEGIN [A-Z ]*PRIVATE KEY-----/;
function redact(v) {
  if (typeof v !== 'string') return v;
  if (SECRET_VALUE.test(v)) return MASK;
  return v.replace(/(\w+:\/\/[^/\s:@]*):[^@/\s]+@/g, `$1:${MASK}@`);
}
const col = (id, label, type = 'string') => ({ id, label, type });
const iso = (t) => (t ? new Date(t).toISOString() : null);
const mb = (b) => (b == null ? null : Math.round((Number(b) / 1048576) * 10) / 10);
const pct = (x) => (x == null ? null : Math.round(Number(x) * 10) / 10);
const arr = (x) => (Array.isArray(x) ? x : []);

class DataSources {
  constructor({ get, t = (s) => s, now = Date.now, ctx = null }) { this.get = get; this.t = t; this.now = now; this.ctx = ctx; this.cache = new Map(); this.asked = new Map(); }
  sources() {
    const t = this.t;
    return [
      { id: 'services', label: t('Hostwerk services'), kind: 'table', params: [] },
      { id: 'service', label: t('State of a service'), kind: 'value', params: [{ id: 'service', label: t('Service') }] },
      { id: 'resources', label: t('Resources per service (CPU, RAM)'), kind: 'table', params: [] },
      { id: 'system', label: t('System load'), kind: 'value', params: [{ id: 'metric', label: t('Value'), options: ['cpu', 'mem', 'memUsed', 'diskFree', 'diskUsed', 'services', 'servicesRunning'] }] },
      { id: 'disks', label: t('Disks'), kind: 'table', params: [] },
      { id: 'minecraft', label: t('Minecraft servers'), kind: 'table', params: [] },
      { id: 'players', label: t('Players online'), kind: 'table', params: [{ id: 'server', label: t('Server (empty = all)') }] },
      { id: 'playerCount', label: t('Number of players online'), kind: 'value', params: [{ id: 'server', label: t('Server (empty = all)') }] },
      { id: 'backups', label: t('Backups'), kind: 'table', params: [] },
      { id: 'tasks', label: t('Scheduled tasks'), kind: 'table', params: [] },
      { id: 'websites', label: t('Websites and domains'), kind: 'table', params: [] },
      { id: 'databases', label: t('Database servers'), kind: 'table', params: [] },
      { id: 'certificates', label: t('Certificates'), kind: 'table', params: [] },
      { id: 'uptime', label: t('Uptime checks'), kind: 'table', params: [] },
      { id: 'hostwerk', label: t('Hostwerk'), kind: 'value', params: [{ id: 'metric', label: t('Value'), options: ['version', 'uptime', 'uptimeHours', 'osUptimeHours'] }] },
    ];
  }
  async api(p, ttl = 5000) {
    const c = this.cache.get(p);
    if (c && this.now() - c.t < ttl) return c.v;
    const v = await this.get(p);
    this.cache.set(p, { t: this.now(), v });
    if (this.cache.size > 50) this.cache.delete(this.cache.keys().next().value);
    return v;
  }
  async query(id, params = {}) {
    const fn = this[`q_${id}`];
    if (typeof fn !== 'function' || !this.sources().some((s) => s.id === id)) throw Object.assign(new Error('Unknown data source'), { status: 404 });
    this.asked.set(id, this.now());
    const r = await fn.call(this, params || {});
    return { columns: r.columns || [], rows: (r.rows || []).slice(0, 5000).map((row) => row.map(redact)), ...(r.value !== undefined ? { value: redact(r.value) } : {}), updatedAt: iso(this.now()) };
  }
  async overview() { return this.api('/overview'); }
  async flat() {
    const o = await this.overview();
    const out = [];
    const proc = (x) => (x && x.proc) || {};
    for (const w of arr(o.web)) out.push({ id: w.id, kind: 'web', name: w.name || w.id, state: proc(w).state || 'stopped', port: w.httpPort || null, ...proc(w) });
    for (const d of arr(o.db)) out.push({ id: d.id, kind: 'db', name: d.name || d.id, state: proc(d).state || 'stopped', port: d.port || null, ...proc(d) });
    for (const a of arr(o.apps)) out.push({ id: a.id, kind: a.kind === 'docker' ? 'docker' : 'app', name: a.name || a.id, state: proc(a).state || 'stopped', port: a.port || null, ...proc(a) });
    if (o.mail && typeof o.mail === 'object' && (o.mail.installed || o.mail.proc)) out.push({ id: 'mail', kind: 'mail', name: 'Mail', state: proc(o.mail).state || (o.mail.running ? 'running' : 'stopped'), port: null, ...proc(o.mail) });
    for (const k of Object.values(o.kinds || {})) for (const it of arr(k.items)) out.push({ id: it.id, kind: (k.sdk && k.sdk.kind) || k.kind, name: it.name || it.id, state: proc(it).state || 'stopped', port: it.port || null, ...proc(it) });
    return out;
  }
  async q_services() {
    const list = await this.flat();
    return { columns: [col('name', this.t('Name')), col('kind', this.t('Type')), col('state', this.t('State')), col('port', this.t('Port'), 'number'), col('cpu', this.t('CPU %'), 'number'), col('memory', this.t('RAM (MB)'), 'number'), col('since', this.t('Running since'), 'date'), col('id', 'ID')],
      rows: list.map((s) => [s.name, s.kind, s.state, s.port, s.stats ? pct(s.stats.cpu) : null, s.stats ? mb(s.stats.memory) : null, iso(s.startedAt), s.id]) };
  }
  async q_service({ service }) {
    const list = await this.flat();
    const s = list.find((x) => x.id === service || x.name === service);
    return { columns: [col('state', this.t('State'))], rows: s ? [[s.state]] : [], value: s ? s.state : null };
  }
  async q_resources() {
    const r = await this.api('/system/resources');
    const sys = r.system || {};
    const rows = arr(r.procs).map((p) => [p.name || p.id, p.state || null, pct(p.cpu), mb(p.memory), p.workers || null, iso(p.startedAt), p.id]);
    const last = arr(sys.history)[arr(sys.history).length - 1] || {};
    rows.unshift([this.t('System'), null, pct(last.cpu), sys.totalMem != null && sys.freeMem != null ? mb(sys.totalMem - sys.freeMem) : null, sys.cpus || null, null, 'system']);
    return { columns: [col('name', this.t('Name')), col('state', this.t('State')), col('cpu', this.t('CPU %'), 'number'), col('memory', this.t('RAM (MB)'), 'number'), col('workers', this.t('Processes'), 'number'), col('since', this.t('Running since'), 'date'), col('id', 'ID')], rows };
  }
  async q_system({ metric = 'cpu' }) {
    const s = await this.api('/summary');
    const v = { cpu: pct(s.cpu), mem: s.mem && s.mem.total ? pct((s.mem.used / s.mem.total) * 100) : null, memUsed: s.mem ? mb(s.mem.used) : null, diskFree: s.disk ? mb(s.disk.free) : null, diskUsed: s.disk && s.disk.total ? pct(((s.disk.total - s.disk.free) / s.disk.total) * 100) : null, services: s.services ? s.services.total : null, servicesRunning: s.services ? s.services.running : null }[metric];
    return { columns: [col('value', String(metric), 'number')], rows: [[v ?? null]], value: v ?? null };
  }
  async q_disks() {
    const d = await this.api('/system/diskguard', 30000);
    return { columns: [col('label', this.t('Disk')), col('path', this.t('Path')), col('free', this.t('Free (GB)'), 'number'), col('total', this.t('Size (GB)'), 'number'), col('freePct', this.t('Free %'), 'number')],
      rows: arr(d.disks).map((x) => [x.label || x.id, x.path || null, x.free != null ? Math.round(x.free / 1073741824 * 10) / 10 : null, x.total != null ? Math.round(x.total / 1073741824 * 10) / 10 : null, pct(x.pct)]) };
  }
  async mcServers() {
    const o = await this.overview();
    const list = arr(o.minecraft).length ? o.minecraft : arr(((o.kinds || {}).mc || {}).items);
    return list.map((s) => { const p = s.proc || {}; const ex = p.extra || {}; return { id: s.id, name: s.name || s.id, state: p.state || 'stopped', version: s.version || null, type: s.typeLabel || s.type || null, players: arr(ex.players), count: ex.playerCount != null ? ex.playerCount : arr(ex.players).length, max: ex.maxPlayersLive || s.maxPlayers || null, tps: ex.tps != null ? ex.tps : null, port: s.port || null }; });
  }
  async q_minecraft() {
    const list = await this.mcServers();
    return { columns: [col('name', this.t('Name')), col('state', this.t('State')), col('players', this.t('Players'), 'number'), col('max', this.t('Max. players'), 'number'), col('tps', 'TPS', 'number'), col('version', this.t('Version')), col('type', this.t('Type')), col('port', this.t('Port'), 'number')],
      rows: list.map((s) => [s.name, s.state, s.count, s.max, s.tps, s.version, s.type, s.port]) };
  }
  async q_players({ server = '' }) {
    const list = (await this.mcServers()).filter((s) => !server || s.id === server || s.name === server);
    const rows = [];
    for (const s of list) for (const p of s.players) rows.push([typeof p === 'string' ? p : p && p.name, s.name]);
    return { columns: [col('player', this.t('Player')), col('server', this.t('Server'))], rows };
  }
  async q_playerCount({ server = '' }) {
    const list = (await this.mcServers()).filter((s) => !server || s.id === server || s.name === server);
    const n = list.reduce((a, s) => a + (Number(s.count) || 0), 0);
    return { columns: [col('value', this.t('Players'), 'number')], rows: [[n]], value: n };
  }
  async q_backups() {
    const b = await this.api('/backups/all', 30000);
    return { columns: [col('time', this.t('Time'), 'date'), col('service', this.t('Service')), col('kind', this.t('Type')), col('label', this.t('Backup')), col('size', this.t('Size (MB)'), 'number'), col('auto', this.t('Automatic'), 'boolean')],
      rows: arr(b.items).slice(0, 2000).map((x) => [iso(x.t), x.svcName || x.svc || null, x.kind || null, x.label || x.id, mb(x.size), !!x.auto]) };
  }
  async q_tasks() {
    const r = await this.api('/tasks', 10000);
    return { columns: [col('name', this.t('Name')), col('action', this.t('Action')), col('enabled', this.t('On'), 'boolean'), col('next', this.t('Next run'), 'date'), col('last', this.t('Last run'), 'date'), col('status', this.t('Last result')), col('success', this.t('Success rate %'), 'number')],
      rows: arr(r.list).map((x) => [x.name, x.actionLabel || (x.action && x.action.type) || null, !!x.enabled, iso(x.nextRun), iso(x.lastRun), x.lastStatus || null, x.successRate != null ? pct(x.successRate) : null]) };
  }
  async q_websites() {
    const list = await this.api('/web', 10000);
    const rows = [];
    for (const w of arr(list)) {
      const st = (w.proc && w.proc.state) || 'stopped';
      if (!arr(w.vhosts).length) rows.push([w.name || w.id, null, st, w.httpPort || null, !!w.https, w.docRoot || null]);
      for (const v of arr(w.vhosts)) rows.push([w.name || w.id, v.domain || null, st, w.httpPort || null, !!(v.tls || w.https), v.proxy ? `→ ${v.proxy}` : v.root || null]);
    }
    return { columns: [col('server', this.t('Web server')), col('domain', this.t('Domain')), col('state', this.t('State')), col('port', this.t('Port'), 'number'), col('https', 'HTTPS', 'boolean'), col('target', this.t('Folder / target'))], rows };
  }
  async q_databases() {
    const list = await this.api('/dbs', 10000);
    return { columns: [col('name', this.t('Name')), col('type', this.t('Type')), col('version', this.t('Version')), col('port', this.t('Port'), 'number'), col('state', this.t('State')), col('id', 'ID')],
      rows: arr(list).map((d) => [d.name || d.id, d.typeLabel || d.type, d.version || null, d.port || null, (d.proc && d.proc.state) || 'stopped', d.id]) };
  }
  async q_certificates() {
    const c = await this.api('/live/certs', 30000);
    const rows = arr(c.items).map((x) => [x.domain, x.web || null, iso(x.expiresAt), x.daysLeft != null ? x.daysLeft : null, x.staging ? 'staging' : "Let's Encrypt"]);
    for (const m of arr(c.missing)) rows.push([m.domain, m.web || null, null, null, this.t('missing')]);
    if (c.local && c.local.ca) rows.push([this.t('Local CA'), null, iso(c.local.ca.expiresAt), c.local.ca.daysLeft, this.t('local')]);
    return { columns: [col('domain', this.t('Domain')), col('server', this.t('Web server')), col('expires', this.t('Expires'), 'date'), col('days', this.t('Days left'), 'number'), col('kind', this.t('Type'))], rows };
  }
  async q_uptime() {
    const u = await this.api('/live/uptime', 10000);
    return { columns: [col('name', this.t('Name')), col('url', this.t('Address')), col('ok', this.t('Reachable'), 'boolean'), col('uptime24', this.t('Uptime 24 h %'), 'number'), col('ms', this.t('Response (ms)'), 'number'), col('since', this.t('Since'), 'date'), col('error', this.t('Error'))],
      rows: arr(u.checks).map((x) => [x.name, x.url, !!x.ok, x.uptime24 != null ? pct(x.uptime24 <= 1 ? x.uptime24 * 100 : x.uptime24) : null, x.ms != null ? x.ms : x.avgMs, iso(x.since), x.error || null]) };
  }
  async q_hostwerk({ metric = 'version' }) {
    const o = await this.api('/system/overview', 10000);
    const h = o.host || {};
    const v = { version: h.version || null, uptime: h.uptime != null ? Math.floor(h.uptime / 1000) : null, uptimeHours: h.uptime != null ? Math.round(h.uptime / 360000) / 10 : null, osUptimeHours: h.osUptime != null ? Math.round(h.osUptime / 360000) / 10 : null }[metric];
    return { columns: [col('value', String(metric))], rows: [[v ?? null]], value: v ?? null };
  }

  /** Push "changed" for sources somebody looks at – driven by Hostwerk's own broadcasts */
  watch(bus, broadcast) {
    const MAP = { status: ['services', 'service', 'minecraft', 'players', 'playerCount', 'databases', 'websites'], stats: ['resources', 'system'], 'mc-player-join': ['players', 'playerCount', 'minecraft'], 'mc-player-leave': ['players', 'playerCount', 'minecraft'], tasks: ['tasks'], 'task-failed': ['tasks'], backups: ['backups'], 'backup-failed': ['backups'], 'uptime-down': ['uptime'], 'cert-watch': ['certificates'], 'web-health': ['websites'] };
    const last = new Map(); const timers = new Map();
    const fire = (src) => {
      if (timers.has(src)) return;
      const wait = Math.max(0, 3000 - (this.now() - (last.get(src) || 0)));
      const t = setTimeout(() => { timers.delete(src); last.set(src, this.now()); for (const k of [...this.cache.keys()]) this.cache.delete(k); try { broadcast({ type: 'office-data', source: src, updatedAt: iso(this.now()) }); } catch { /* egal */ } }, wait);
      if (t.unref) t.unref(); timers.set(src, t);
    };
    const l = (m) => { if (!m || !MAP[m.type]) return; for (const src of MAP[m.type]) { const a = this.asked.get(src); if (a && this.now() - a < 15 * 60000) fire(src); } };
    bus.on('broadcast', l);
    return () => { bus.off('broadcast', l); for (const t of timers.values()) clearTimeout(t); timers.clear(); };
  }
}

module.exports = { DataSources, redact };
