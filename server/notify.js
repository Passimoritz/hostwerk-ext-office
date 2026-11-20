'use strict';
/*
 * Sending a message to one channel the user picked (rules, forms, scheduled reports):
 *   'hostwerk' → host.notifications (own event types with a switch in Hostwerk's settings)
 *   'discord' | 'ntfy' | 'mail' → the addresses configured under Hostwerk → Live → Alerts (settings.live.alerts)
 *   'all' → every configured alert channel (ctx.monitor.send – quiet hours respected)
 * Attachments (reports): Discord as a file upload, mail as MIME attachment through Hostwerk's mail relay.
 *
 *   new Notifier({ host }) · .channels() → [{ id, label, ready }] · .send(channel, { event, title, text, kind, url, attachment: { name, data, type } })
 */
const crypto = require('crypto');
const os = require('os');

const DISCORD = /^https:\/\/(discord\.com|discordapp\.com|ptb\.discord\.com|canary\.discord\.com)\/api\/webhooks\/\d+\/[\w-]+$/;
const b64h = (s) => `=?UTF-8?B?${Buffer.from(String(s)).toString('base64')}?=`;

class Notifier {
  constructor({ host }) { this.host = host; this.ctx = host.ctx || {}; }
  alerts() { try { return (this.ctx.monitor && this.ctx.monitor.alerts) || ((this.ctx.settings.get().live || {}).alerts) || {}; } catch { return {}; } }
  relay() { try { return (this.ctx.settings.get().mail || {}).relay || {}; } catch { return {}; } }
  channels() {
    const a = this.alerts();
    return [
      { id: 'hostwerk', label: this.host.t('Hostwerk notification'), ready: true },
      { id: 'discord', label: 'Discord', ready: DISCORD.test(String(a.discord || '')) },
      { id: 'ntfy', label: 'ntfy', ready: /^https?:\/\//.test(String(a.ntfy || '')) },
      { id: 'mail', label: this.host.t('E-mail'), ready: /^[^\s@]+@[^\s@]+$/.test(String(a.email || '')) },
      { id: 'all', label: this.host.t('All alert channels'), ready: true },
    ];
  }
  async send(channel, { event = 'rule', title, text = '', kind = 'info', url = null, attachment = null } = {}) {
    const a = this.alerts();
    const t = String(title || 'Office').slice(0, 200); const body = String(text || '').slice(0, 3000);
    const hostName = os.hostname();
    if (channel === 'hostwerk' || !channel) { this.host.notifications.notify(event, { title: t, text: body.slice(0, 400), kind, url: url || '/ext/office' }); return { hostwerk: true }; }
    if (channel === 'all') { if (!this.ctx.monitor || typeof this.ctx.monitor.send !== 'function') throw Object.assign(new Error('Alerts are not available.'), { status: 409 }); return { all: await this.ctx.monitor.send(t, body, { kind, force: true, event: 'office' }) }; }
    if (channel === 'discord') {
      if (!DISCORD.test(String(a.discord || ''))) throw Object.assign(new Error('No Discord webhook is set up (Hostwerk → Live → Alerts).'), { status: 409 });
      const content = `**${t}**\n${body}\n_${hostName}_`.slice(0, 1900);
      let r;
      if (attachment && attachment.data && attachment.data.length <= 8 * 1048576 && typeof FormData === 'function') {
        const fd = new FormData();
        fd.append('payload_json', JSON.stringify({ username: 'Hostwerk', content }));
        fd.append('files[0]', new Blob([attachment.data], { type: attachment.type || 'application/octet-stream' }), String(attachment.name || 'file').replace(/[^\w. ()-]/g, '_'));
        r = await fetch(a.discord, { method: 'POST', body: fd, signal: AbortSignal.timeout(30000) });
      } else r = await fetch(a.discord, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ username: 'Hostwerk', content }), signal: AbortSignal.timeout(10000) });
      if (!r.ok) throw new Error(`Discord: HTTP ${r.status}`);
      return { discord: true };
    }
    if (channel === 'ntfy') {
      if (!/^https?:\/\//.test(String(a.ntfy || ''))) throw Object.assign(new Error('No ntfy address is set up (Hostwerk → Live → Alerts).'), { status: 409 });
      const r = await fetch(a.ntfy, { method: 'POST', headers: { Title: t.replace(/[^\x20-\x7e]/g, '?').slice(0, 200), Tags: kind === 'error' ? 'rotating_light' : 'memo' }, body: `${body}\n(${hostName})`, signal: AbortSignal.timeout(10000) });
      if (!r.ok) throw new Error(`ntfy: HTTP ${r.status}`);
      return { ntfy: true };
    }
    if (channel === 'mail') {
      const to = String(a.email || '');
      if (!/^[^\s@]+@[^\s@]+$/.test(to)) throw Object.assign(new Error('No e-mail address is set up (Hostwerk → Live → Alerts).'), { status: 409 });
      const rel = this.relay();
      if (!attachment || !rel.host) {
        const nm = this.host.core('notifymail');
        return { mail: await nm.sendNotice(this.ctx, { to, subject: `[Hostwerk] ${t}`, text: `${body}\n\n-- Hostwerk Office (${hostName})` }) };
      }
      const from = rel.from || rel.user;
      const bnd = `----hw${crypto.randomBytes(8).toString('hex')}`;
      const raw = [`From: Hostwerk <${from}>`, `To: ${to}`, `Subject: ${b64h(`[Hostwerk] ${t}`)}`, `Date: ${new Date().toUTCString()}`, `Message-ID: <${crypto.randomUUID()}@hostwerk>`, 'MIME-Version: 1.0', `Content-Type: multipart/mixed; boundary="${bnd}"`, '',
        `--${bnd}`, 'Content-Type: text/plain; charset=UTF-8', 'Content-Transfer-Encoding: base64', '', Buffer.from(`${body}\n\n-- Hostwerk Office (${hostName})`).toString('base64').replace(/.{76}/g, '$&\r\n'),
        `--${bnd}`, `Content-Type: ${attachment.type || 'application/octet-stream'}; name="${b64h(attachment.name)}"`, 'Content-Transfer-Encoding: base64', `Content-Disposition: attachment; filename="${b64h(attachment.name)}"`, '', attachment.data.toString('base64').replace(/.{76}/g, '$&\r\n'),
        `--${bnd}--`, ''].join('\r\n');
      const smtp = this.host.core('smtp-client'); const sec = this.host.core('secrets');
      await smtp.sendRaw({ host: rel.host, port: Number(rel.port) || 587, secure: !!rel.secure, user: rel.user, pass: sec.open(rel.pass), from, to, raw });
      return { mail: true };
    }
    throw Object.assign(new Error('Unknown channel'), { status: 400 });
  }
}

module.exports = { Notifier, DISCORD };
