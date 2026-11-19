/*
 * E-mail view for .eml (mime.js) and .msg (msg.js): header, sanitised HTML body in a sandboxed frame (remote images
 * blocked until "Load images"), cid: images inline, plain text fallback, attachments with download.
 */
import { useEffect, useMemo, useState } from 'react';
import DOMPurify from 'dompurify';
import { Mail, Paperclip, Download, ImageOff, Image as ImageIcon, FileText, Code } from 'lucide-react';
import { Btn, Callout, Seg } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { readMail, htmlToText } from './mime.js';
import { readMsg } from './msg.js';
import { ActionRow, SafeFrame, fmtSize, fmtDate } from './common.jsx';
import { Keep } from '../_shared/keep.jsx';

export function parseMailFile(bytes, ext) { return ext === 'msg' ? readMsg(bytes) : readMail(bytes); }

const people = (list) => (list || []).map((p) => (p.name && p.email ? `${p.name} <${p.email}>` : p.name || p.email)).join(', ');

/** sanitise mail HTML; remote resources are dropped unless allowed; cid: → blob URLs */
export function cleanMailHtml(html, { cids = {}, remote = false } = {}) {
  let blocked = 0;
  const purify = DOMPurify();
  purify.addHook('afterSanitizeAttributes', (node) => {
    for (const attr of ['src', 'background', 'poster', 'srcset']) {
      if (!node.hasAttribute || !node.hasAttribute(attr)) continue;
      const v = node.getAttribute(attr).trim();
      if (/^cid:/i.test(v)) { const u = cids[v.slice(4).replace(/^<|>$/g, '')]; if (u) node.setAttribute(attr, u); else node.removeAttribute(attr); } else if (/^(https?:)?\/\//i.test(v) && !remote) { node.removeAttribute(attr); node.setAttribute('data-of-blocked', '1'); blocked++; }
    }
    if (node.hasAttribute && node.hasAttribute('style') && !remote) {
      const s = node.getAttribute('style');
      if (/url\(\s*['"]?(https?:)?\/\//i.test(s)) { node.setAttribute('style', s.replace(/url\([^)]*\)/gi, 'none')); blocked++; }
    }
    if (node.tagName === 'A') node.removeAttribute('target');
  });
  const body = purify.sanitize(html, { WHOLE_DOCUMENT: false, FORBID_TAGS: ['script', 'iframe', 'object', 'embed', 'form', 'input', 'button', 'meta', 'link', 'base'], FORBID_ATTR: ['onerror', 'onload'], ADD_ATTR: ['data-of-blocked'] });
  const csp = `default-src 'none'; img-src data: blob:${remote ? ' https: http:' : ''}; style-src 'unsafe-inline'; font-src data:${remote ? ' https:' : ''};`;
  // paper look: mails are designed for white backgrounds
  const doc = `<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="${csp}"><style>html,body{margin:0;background:#fff;color:#111;font:14px/1.5 -apple-system,'Segoe UI',Roboto,Arial,sans-serif}body{padding:16px;overflow-wrap:anywhere}img{max-width:100%;height:auto}pre{white-space:pre-wrap}blockquote{margin:0 0 0 8px;padding-left:10px;border-left:3px solid #ccc;color:#444}a{color:#1f5fbf}</style></head><body>${body}</body></html>`;
  return { html: doc, blocked };
}

const textToHtml = (s) => `<pre style="font:13px/1.55 ui-monospace,Consolas,monospace;white-space:pre-wrap">${String(s || '').replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]).replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1">$1</a>')}</pre>`;

export default function MailView({ file, bytes, host }) {
  const mail = useMemo(() => parseMailFile(bytes, file.ext), [bytes, file.ext]);
  const [remote, setRemote] = useState(false);
  const [mode, setMode] = useState(mail.html ? 'html' : 'text');
  const [headers, setHeaders] = useState(false);
  const cidUrls = useMemo(() => {
    const m = {};
    for (const a of mail.attachments) if (a.cid && a.bytes) m[a.cid] = URL.createObjectURL(new Blob([a.bytes], { type: a.mime }));
    return m;
  }, [mail]);
  useEffect(() => () => Object.values(cidUrls).forEach((u) => URL.revokeObjectURL(u)), [cidUrls]);
  const usedCids = useMemo(() => new Set([...String(mail.html).matchAll(/cid:([^"'\s)>]+)/gi)].map((m) => m[1])), [mail]);
  const shown = useMemo(() => (mode === 'html' && mail.html ? cleanMailHtml(mail.html, { cids: cidUrls, remote }) : { html: cleanMailHtml(textToHtml(mail.text)).html, blocked: 0 }), [mail, mode, remote, cidUrls]);
  const atts = mail.attachments.filter((a) => !(a.cid && usedCids.has(a.cid)));
  const openLink = (url) => { if (/^(https?:|mailto:)/i.test(url)) { if (host.openLink) host.openLink(url); } };
  useEffect(() => { host.status(atts.length ? t('{n} attachments', { n: atts.length }) : ''); }, [atts.length]); // eslint-disable-line react-hooks/exhaustive-deps

  return (
    <div className="of-view-mail">
      <div className="of-view-mailhead card">
        <div className="card-body stack">
          <h2 className="of-view-subject"><Mail size={18} /> {mail.subject ? <Keep>{mail.subject}</Keep> : t('(no subject)')}</h2>
          <table className="of-view-hdr"><tbody>
            <tr><th>{t('From')}</th><td><Keep>{people(mail.from) || '–'}</Keep></td></tr>
            {mail.to.length > 0 && <tr><th>{t('To')}</th><td><Keep>{people(mail.to)}</Keep></td></tr>}
            {mail.cc.length > 0 && <tr><th>{t('Cc')}</th><td><Keep>{people(mail.cc)}</Keep></td></tr>}
            {mail.bcc && mail.bcc.length > 0 && <tr><th>{t('Bcc')}</th><td><Keep>{people(mail.bcc)}</Keep></td></tr>}
            {mail.replyTo && mail.replyTo.length > 0 && <tr><th>{t('Reply to')}</th><td><Keep>{people(mail.replyTo)}</Keep></td></tr>}
            {mail.date && <tr><th>{t('Date')}</th><td>{fmtDate(mail.date)}</td></tr>}
          </tbody></table>
          {atts.length > 0 && (
            <div className="of-view-atts">
              {atts.map((a, i) => (
                <button key={i} type="button" className="of-view-att" disabled={!a.bytes} title={a.bytes ? t('Download {name}', { name: a.name }) : t('Embedded message – open the .msg in Outlook')} onClick={() => host.download(a.bytes, a.name, a.mime)}>
                  <Paperclip size={14} /><span className="of-view-ellipsis"><Keep>{a.name}</Keep></span><span className="small muted">{fmtSize(a.size)}</span>{a.bytes && <Download size={12} />}
                </button>
              ))}
            </div>
          )}
        </div>
      </div>
      {shown.blocked > 0 && !remote && (
        <div className="of-view-pad0"><Callout kind="info" icon={ImageOff} dismissible={false}>
          <span className="row">{t('Images from the internet are blocked to protect your privacy.')} <Btn className="sm" icon={ImageIcon} onClick={() => setRemote(true)}>Load images</Btn></span>
        </Callout></div>
      )}
      {mail.rtfOnly && <div className="of-view-pad0"><Callout kind="info">{t('This message only has a rich-text (RTF) body, which is not shown here. Open it in Outlook or convert it.')}</Callout></div>}
      <div className="of-view-mailbody">
        <SafeFrame html={shown.html} name={mail.subject || t('E-mail')} onLink={openLink} minHeight={300} />
      </div>
      {headers && <pre className="of-view-raw mono small"><Keep>{(mail.headers || []).map(([k, v]) => `${k}: ${v}`).join('\n')}</Keep></pre>}
      <ActionRow file={file} bytes={bytes} host={host}>
        {mail.html && mail.text ? <Seg value={mode} onChange={setMode} options={[{ id: 'html', label: t('Formatted'), icon: Code }, { id: 'text', label: t('Plain text'), icon: FileText }]} /> : null}
        {mail.headers && mail.headers.length > 0 && <Btn className="sm ghost" onClick={() => setHeaders(!headers)}>{headers ? 'Hide headers' : 'Show all headers'}</Btn>}
      </ActionRow>
    </div>
  );
}

export function mailText(bytes, ext) {
  const m = parseMailFile(bytes, ext);
  return [`${t('Subject')}: ${m.subject}`, `${t('From')}: ${people(m.from)}`, `${t('To')}: ${people(m.to)}`, m.date ? `${t('Date')}: ${m.date.toISOString()}` : '', '', m.text || htmlToText(m.html), m.attachments.length ? `\n${t('Attachments')}: ${m.attachments.map((a) => a.name).join(', ')}` : ''].filter((x) => x !== null).join('\n');
}
export function mailPreview(bytes, ext) { const m = parseMailFile(bytes, ext); return { title: m.subject || t('(no subject)'), sub: people(m.from), line: (m.text || '').slice(0, 140) }; }
