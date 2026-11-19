'use strict';
/*
 * Full editor with one click: an ONLYOFFICE Document Server container through Hostwerk's Docker support
 * (host.ctx.docker – install Docker under Apps → Docker if it is missing). ONLYOFFICE is the only full editor.
 * The container listens on 127.0.0.1 only; it reaches Hostwerk's ONLYOFFICE routes through host.docker.internal.
 */
const crypto = require('crypto');

const APPS = {
  onlyoffice: { name: 'hostwerk-office-onlyoffice', image: 'onlyoffice/documentserver:latest', port: 9981, inner: 80, label: 'ONLYOFFICE Docs' },
};

/** docker run arguments (array – never a shell string) */
function runArgs(app, { secret, port }) {
  const a = APPS[app];
  if (!a) throw new Error('Unknown office server');
  const p = Number(port) || a.port;
  const base = ['run', '-d', '--name', a.name, '--restart', 'unless-stopped', '-p', `127.0.0.1:${p}:${a.inner}`, '--add-host', 'host.docker.internal:host-gateway', '--label', 'hostwerk.office=1'];
  return [...base, '-e', 'JWT_ENABLED=true', '-e', `JWT_SECRET=${secret}`, '-e', 'JWT_HEADER=Authorization', '-e', 'ALLOW_PRIVATE_IP_ADDRESS=true', '-e', 'ALLOW_META_IP_ADDRESS=true', a.image];
}
const newSecret = () => crypto.randomBytes(24).toString('base64url');

class Servers {
  /** deps: { docker: () => ctx.docker | null, runJob, t, settings(), save(patch), hostPort(), log } */
  constructor(deps) { this.d = deps; this.jobs = {}; }
  docker() { const d = this.d.docker(); if (!d || typeof d.x !== 'function') throw Object.assign(new Error('Hostwerk\'s Docker support is not available.'), { status: 501 }); return d; }
  async inspect(app) {
    const d = this.d.docker();
    if (!d) return null;
    const r = await d.x(['inspect', '--format', '{{.State.Status}}|{{.Config.Image}}', APPS[app].name], { timeout: 20000 }).catch(() => null);
    if (!r || r.code !== 0) return null;
    const [state, image] = r.stdout.trim().split('|');
    return { state, image };
  }
  async state() {
    let docker = null;
    try { const d = this.d.docker(); docker = d ? await d.status() : null; } catch (e) { docker = { error: e.message }; }
    const s = this.d.settings().integrations;
    const out = { docker: docker ? { installed: !!docker.cli, running: !!docker.running, version: docker.version || null, engine: docker.engine || null } : { installed: false, running: false }, apps: {} };
    for (const app of Object.keys(APPS)) {
      const c = docker && docker.running ? await this.inspect(app) : null;
      const j = this.jobs[app];
      const running = j && j.state === 'running';
      const url = (s[app] || {}).url || '';
      // shell states: missing | pulling | installing | starting | ready | stopped | error
      const st = running ? (j.progress_ < 80 ? 'pulling' : j.progress_ < 90 ? 'installing' : 'starting') : j && j.state === 'error' ? 'error' : c ? (c.state === 'running' ? 'ready' : c.state === 'restarting' || c.state === 'created' ? 'starting' : 'stopped') : url ? 'ready' : 'missing';
      out[app] = { state: st, url, external: !c && !!url, progress: running ? { percent: j.progress_, text: j.text } : null, error: j && j.state === 'error' ? j.error : null, jobId: j ? j.id : null, container: APPS[app].name, image: APPS[app].image, label: APPS[app].label, containerState: c ? c.state : null };
    }
    out.dockerInfo = out.docker;
    out.docker = !!(out.dockerInfo.installed && out.dockerInfo.running);
    if (!out.docker) out.dockerHint = this.d.t(out.dockerInfo.installed ? 'Docker is not running – start it under Apps → Docker, then try again.' : 'Docker is not installed – install it under Apps → Docker (one click), then try again.');
    delete out.apps;
    return out;
  }
  /** Create (or re-create) and start the container as a Hostwerk job */
  setup(app) {
    const a = APPS[app];
    if (!a) throw Object.assign(new Error('Unknown office server'), { status: 400 });
    const d = this.docker();
    const T = this.d.t;
    const job = this.d.runJob(T('Set up {name}', { name: a.label }), async (j) => {
      j.progress(2, T('Checking Docker …'));
      const st = await d.status();
      if (!st.cli) throw Object.assign(new Error(T('Docker is not installed – install it under Apps → Docker (one click), then try again.')), { status: 409 });
      if (!st.running) throw Object.assign(new Error(T('Docker is not running – start it under Apps → Docker, then try again.')), { status: 409 });
      j.progress(5, T('Downloading {image} …', { image: a.image }));
      if (typeof d.pullInto === 'function') await d.pullInto(j, a.image, [5, 80]); else await d.run(['pull', a.image], 30 * 60000);
      await d.x(['rm', '-f', a.name], { timeout: 60000 }).catch(() => {});
      const secret = newSecret();
      const hostUrl = `http://host.docker.internal:${this.d.hostPort()}`;
      j.progress(85, T('Starting {name} …', { name: a.label }));
      const r = await d.x(runArgs(app, { secret }), { timeout: 120000 });
      if (r.code !== 0) throw new Error((r.stderr || r.stdout || 'docker run failed').trim().split('\n').pop());
      const url = `http://localhost:${a.port}`;
      const patch = { integrations: { hostUrl, [app]: { url, secret } } };
      if (!this.d.settings().integrations.defaultOpen || this.d.settings().integrations.defaultOpen === 'builtin') patch.integrations.defaultOpen = 'builtin';
      this.d.save(patch);
      j.progress(90, T('Waiting until {name} answers …', { name: a.label }));
      const probe = `${url}/healthcheck`;
      const until = Date.now() + 5 * 60000;
      for (;;) {
        try { const res = await fetch(probe, { signal: AbortSignal.timeout(5000) }); if (res.ok) break; } catch { /* starting */ }
        if (Date.now() > until) throw new Error(T('{name} was started but does not answer yet – look at the container log under Apps → Docker.', { name: a.label }));
        await new Promise((x) => setTimeout(x, 3000));
        if (j.checkpoint) await j.checkpoint();
      }
      j.progress(100, T('{name} is ready.', { name: a.label }));
      return { url };
    }, { key: `office-server-${app}` });
    this.jobs[app] = job;
    return job;
  }
  async action(app, act) {
    const a = APPS[app];
    if (!a || !['start', 'stop', 'restart'].includes(act)) throw Object.assign(new Error('Unknown office server'), { status: 400 });
    const r = await this.docker().x([act, a.name], { timeout: 120000 });
    if (r.code !== 0) throw Object.assign(new Error((r.stderr || '').trim() || 'docker failed'), { status: 409 });
    return this.state();
  }
  async remove(app) {
    const a = APPS[app];
    if (!a) throw Object.assign(new Error('Unknown office server'), { status: 400 });
    const d = this.d.docker();
    if (d) await d.x(['rm', '-f', a.name], { timeout: 120000 }).catch(() => {});
    const s = this.d.settings().integrations;
    const patch = { integrations: { [app]: { url: '', secret: '' } } };
    if (s.defaultOpen === app) patch.integrations.defaultOpen = 'builtin';
    this.d.save(patch);
    return this.state();
  }
}

module.exports = { Servers, runArgs, APPS };
