/*
 * Dialogs and panels of the PDF editor's "Tools" tab: protect (AES-256), compress, OCR, redact, compare.
 * The heavy lifting is in the pure modules (security.js, compress.js, ocr.js/textlayer.js, redact.js, compare.js);
 * these components collect options, show progress and results.
 */
import { useEffect, useMemo, useRef, useState } from 'react';
import { Lock, LockOpen, ShieldCheck, Gauge, ScanText, EyeOff, GitCompare, Trash2, Search, CheckCircle2, AlertTriangle, FileUp, Eye, KeyRound } from 'lucide-react';
import { Modal, Btn, Field, Switch, Callout, Seg, Progress } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { KINDS } from './redact.js';
import { comparePages, pixelDiff } from './compare.js';
import { openDoc } from './open.js';
import { fmtSize } from './Sidebar.jsx';
import { Keep } from '../_shared/keep.jsx';

/* ---------------- protect ---------------- */
export function ProtectDialog({ protection, onApply, onClose, onSaveDecrypted }) {
  const [pw, setPw] = useState('');
  const [pw2, setPw2] = useState('');
  const [owner, setOwner] = useState('');
  const [perm, setPerm] = useState({ print: true, copy: true, modify: true, annotate: true, fill: true, assemble: true });
  const [show, setShow] = useState(false);
  const mismatch = pw2 && pw !== pw2;
  const can = (pw || owner) && !mismatch && (!pw || pw === pw2);
  const strength = pw.length >= 12 && /[^a-zA-Z]/.test(pw) ? 'strong' : pw.length >= 8 ? 'ok' : pw ? 'weak' : '';
  const box = (k, label) => <label className="of-pdf-check"><input type="checkbox" checked={perm[k]} onChange={(e) => setPerm({ ...perm, [k]: e.target.checked })} /> {label}</label>;
  return (
    <Modal title="Protect PDF" subtitle="AES-256 encryption – opens in every PDF program with the password" onClose={onClose}
      footer={<>
        {protection && <Btn className="ghost" icon={LockOpen} onClick={() => { onApply(null); onClose(); }}>Remove protection</Btn>}
        <span className="grow" />
        <Btn onClick={onClose}>Cancel</Btn>
        <Btn className="primary" icon={Lock} disabled={!can} onClick={() => { onApply({ userPassword: pw, ownerPassword: owner || pw, permissions: perm }); onClose(); }}>Protect</Btn>
      </>}>
      <div className="stack">
        {protection && <Callout kind="info" icon={ShieldCheck} dismissible={false}>{protection.keep ? t('This PDF is protected ({method}). Saving keeps its passwords. Set new ones here or remove the protection.', { method: protection.keep.method === 'rc4' ? 'RC4' : protection.keep.method === 'aes128' ? 'AES-128' : 'AES-256' }) : t('New protection is applied when you save.')}</Callout>}
        <Field label="Password to open" hint={strength ? { weak: t('Weak – use at least 8 characters'), ok: t('Good'), strong: t('Strong') }[strength] : t('Empty = anyone can open it; only the permissions below are restricted')}>
          <span className="row"><input className="input" type={show ? 'text' : 'password'} value={pw} onChange={(e) => setPw(e.target.value)} autoFocus autoComplete="new-password" aria-label="Password to open" />
            <button type="button" className="btn sm ghost icon" title={show ? 'Hide password' : 'Show password'} aria-label={show ? 'Hide password' : 'Show password'} onClick={() => setShow(!show)}><Eye size={15} /></button></span>
          {strength && <div className={`of-pdf-strength ${strength}`}><span /></div>}
        </Field>
        <Field label="Repeat password" hint={mismatch ? t('The passwords are not the same') : undefined}>
          <input className="input" type={show ? 'text' : 'password'} value={pw2} onChange={(e) => setPw2(e.target.value)} autoComplete="new-password" aria-label="Repeat password" />
        </Field>
        <Field label="Permissions password (optional)" hint="Needed to change the permissions – empty = same as the password to open">
          <input className="input" type={show ? 'text' : 'password'} value={owner} onChange={(e) => setOwner(e.target.value)} autoComplete="new-password" aria-label="Permissions password" />
        </Field>
        <div className="of-pdf-perms">
          <span className="small muted">{t('Allowed without the permissions password')}</span>
          {box('print', t('Print'))}{box('copy', t('Copy text and images'))}{box('modify', t('Change the document'))}{box('annotate', t('Add comments'))}{box('fill', t('Fill in forms'))}{box('assemble', t('Insert, rotate, delete pages'))}
        </div>
        {protection && onSaveDecrypted && <div className="row"><KeyRound size={14} /><span className="small grow">{t('Need it without a password?')}</span><Btn className="sm" onClick={onSaveDecrypted}>Save an unprotected copy</Btn></div>}
      </div>
    </Modal>
  );
}

/* ---------------- compress ---------------- */
const LEVELS = { small: { maxPx: 1200, quality: 0.6 }, balanced: { maxPx: 1800, quality: 0.74 }, high: { maxPx: 2600, quality: 0.85 } };
export async function recodeImage(input, { maxPx, quality = 0.74 }) {
  let bmp;
  if (input.kind === 'jpeg') bmp = await createImageBitmap(new Blob([input.data], { type: 'image/jpeg' }));
  else {
    const { w, h, comps, data } = input;
    const id = new ImageData(w, h); const px = id.data;
    for (let i = 0, j = 0; i < w * h; i++, j += comps) { const r = data[j]; px[i * 4] = r; px[i * 4 + 1] = comps === 3 ? data[j + 1] : r; px[i * 4 + 2] = comps === 3 ? data[j + 2] : r; px[i * 4 + 3] = 255; }
    bmp = await createImageBitmap(id);
  }
  const s = Math.min(1, maxPx / Math.max(bmp.width, bmp.height));
  const W = Math.max(1, Math.round(bmp.width * s)); const H = Math.max(1, Math.round(bmp.height * s));
  const c = typeof OffscreenCanvas !== 'undefined' ? new OffscreenCanvas(W, H) : Object.assign(document.createElement('canvas'), { width: W, height: H });
  const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
  ctx.drawImage(bmp, 0, 0, W, H); bmp.close && bmp.close();
  const blob = c.convertToBlob ? await c.convertToBlob({ type: 'image/jpeg', quality }) : await new Promise((r) => c.toBlob(r, 'image/jpeg', quality));
  return { data: new Uint8Array(await blob.arrayBuffer()), w: W, h: H, gray: false };
}

export function CompressDialog({ run, onApply, onClose, hasConverter }) {
  const [level, setLevel] = useState('balanced');
  const [busy, setBusy] = useState(false);
  const [res, setRes] = useState(null);
  const [err, setErr] = useState('');
  const go = async () => {
    setBusy(true); setErr(''); setRes(null);
    try { setRes(await run(LEVELS[level])); } catch (e) { setErr(e.message || String(e)); } finally { setBusy(false); }
  };
  const saved = res ? Math.max(0, 1 - res.after / res.before) : 0;
  return (
    <Modal title="Make PDF smaller" subtitle="Pictures are scaled down and stored more compactly – text and vector graphics stay sharp" onClose={onClose}
      footer={<><Btn onClick={onClose}>Cancel</Btn>{res && res.after < res.before ? <Btn className="primary" onClick={() => { onApply(res.bytes); onClose(); }}>Apply</Btn> : <Btn className="primary" icon={Gauge} busy={busy} onClick={go}>Compress</Btn>}</>}>
      <div className="stack">
        <Seg value={level} onChange={(v) => { setLevel(v); setRes(null); }} options={[{ id: 'small', label: t('Smallest') }, { id: 'balanced', label: t('Balanced') }, { id: 'high', label: t('High quality') }]} />
        <p className="small muted">{{ small: t('About 110 dpi – for e-mail and the screen.'), balanced: t('About 150–200 dpi – good for printing at home.'), high: t('About 250–300 dpi – almost no visible difference.') }[level]}</p>
        {busy && <Progress indeterminate />}
        {res && (
          <div className="of-pdf-result card">
            <div className="of-pdf-bars">
              <div><span className="small muted">{t('Before')}</span><div className="of-pdf-bar"><span style={{ width: '100%' }} /></div><b>{fmtSize(res.before)}</b></div>
              <div><span className="small muted">{t('After')}</span><div className="of-pdf-bar after"><span style={{ width: `${Math.max(2, (res.after / res.before) * 100)}%` }} /></div><b>{fmtSize(res.after)}</b></div>
            </div>
            <p className="small">{res.after < res.before ? t('{p} % smaller · {n} pictures recompressed', { p: Math.round(saved * 100), n: res.images }) : t('This PDF is already compact – nothing to gain.')}</p>
          </div>
        )}
        {err && <Callout kind="err">{err}</Callout>}
        {!hasConverter && <p className="small muted">{t('PDF/A (long-term archiving) needs the Office engine – it is set up under Settings → Office engine.')}</p>}
      </div>
    </Modal>
  );
}

/* ---------------- OCR ---------------- */
export function OcrDialog({ count, current, selection, run, onClose }) {
  const [scope, setScope] = useState('auto');
  const [lang, setLang] = useState('deu+eng');
  const [state, setState] = useState(null); // { done, total, page, words }
  const [err, setErr] = useState('');
  const [finished, setFinished] = useState(null);
  const stop = useRef(false);
  const go = async () => {
    setErr(''); setFinished(null); stop.current = false;
    try {
      const r = await run({ scope, lang, onProgress: (s) => setState(s), stopped: () => stop.current });
      setFinished(r);
    } catch (e) { setErr(e.message || String(e)); } finally { setState(null); }
  };
  return (
    <Modal title="Recognize text (OCR)" subtitle="Scanned pages get an invisible text layer – searchable and copyable in every PDF program" onClose={() => { stop.current = true; onClose(); }}
      footer={finished ? <Btn className="primary" onClick={onClose}>Done</Btn> : <><Btn onClick={() => { stop.current = true; onClose(); }}>{state ? 'Stop' : 'Cancel'}</Btn><Btn className="primary" icon={ScanText} busy={!!state} onClick={go}>Recognize</Btn></>}>
      <div className="stack">
        <Field label="Pages">
          <Seg value={scope} onChange={setScope} options={[{ id: 'auto', label: t('Pages without text') }, { id: 'all', label: t('All {n}', { n: count }) }, { id: 'current', label: selection > 1 ? t('{n} selected', { n: selection }) : t('Page {n}', { n: current }) }]} />
        </Field>
        <Field label="Language">
          <Seg value={lang} onChange={setLang} options={[{ id: 'deu+eng', label: t('German + English') }, { id: 'deu', label: t('German') }, { id: 'eng', label: t('English') }]} />
        </Field>
        {state && (
          <div className="stack">
            <Progress value={(state.done / Math.max(1, state.total)) * 100} />
            <span className="small muted">{state.loading ? t('Loading the OCR engine …') : t('Page {page} – {done} of {total} done', state)}</span>
          </div>
        )}
        {finished && <Callout kind="ok" icon={CheckCircle2} dismissible={false}>{t('{words} words recognized on {pages} pages. Save to keep the text layer.', finished)}</Callout>}
        {err && <Callout kind="err">{err}</Callout>}
        <p className="small muted">{t('The OCR engine (Tesseract, ≈ 23 MB with German and English) is downloaded once and runs entirely on this computer – nothing leaves Hostwerk.')}</p>
      </div>
    </Modal>
  );
}

/* ---------------- redaction panel ---------------- */
export function RedactPanel({ marks, pageNo, onFind, onRemove, onClear, onApply, onGo, busy }) {
  const [kinds, setKinds] = useState(() => new Set(['email', 'iban', 'phone']));
  const [terms, setTerms] = useState('');
  const [found, setFound] = useState(null);
  const list = useMemo(() => Object.entries(marks).flatMap(([pid, l]) => l.map((m) => ({ ...m, pid }))), [marks]);
  const toggle = (k) => { const s = new Set(kinds); if (s.has(k)) s.delete(k); else s.add(k); setKinds(s); };
  return (
    <aside className="of-pdf-redact-panel" aria-label="Redact">
      <div className="of-pdf-panel-head"><EyeOff size={16} /><b>{t('Redact')}</b></div>
      <div className="of-pdf-panel-body stack">
        <p className="small muted">{t('Drag over areas on the page, or let Hostwerk find personal data. “Apply” removes the text, graphics and pixels under the marks for good.')}</p>
        <div className="of-pdf-kinds">
          {Object.entries(KINDS).map(([k, v]) => <label key={k} className="of-pdf-check"><input type="checkbox" checked={kinds.has(k)} onChange={() => toggle(k)} /> {t(v.label)}</label>)}
        </div>
        <Field label="Own words and names" hint="Separate with commas"><input className="input" value={terms} onChange={(e) => setTerms(e.target.value)} placeholder="Erika Mustermann, Projekt X" aria-label="Own words and names" /></Field>
        <Btn icon={Search} busy={busy === 'find'} onClick={async () => setFound(await onFind({ kinds: [...kinds], terms: terms.split(/[,;\n]/).map((x) => x.trim()).filter(Boolean) }))}>Find and mark</Btn>
        {found != null && <span className="small muted">{found ? t('{n} places marked', { n: found }) : t('Nothing found.')}</span>}
        <div className="of-pdf-marks">
          {list.length === 0 ? <span className="small muted">{t('No marks yet.')}</span> : list.map((m) => (
            <div key={m.id} className="of-pdf-mark">
              <button type="button" className="of-pdf-mark-go" onClick={() => onGo(m)} title={t('Show')}>
                <span className="of-pdf-mark-swatch" />
                <span className="small of-pdf-ellipsis">{m.value ? <Keep>{m.value}</Keep> : t('Marked area')}</span>
                <span className="small muted">{t('p. {n}', { n: pageNo(m.pid) })}</span>
              </button>
              <button type="button" className="btn xs ghost icon" title="Remove mark" aria-label="Remove mark" onClick={() => onRemove(m.id)}><Trash2 size={12} /></button>
            </div>
          ))}
        </div>
      </div>
      <div className="of-pdf-panel-foot">
        <Btn className="sm ghost" disabled={!list.length} onClick={onClear}>Clear all</Btn>
        <span className="grow" />
        <Btn className="sm danger" icon={EyeOff} disabled={!list.length} busy={busy === 'apply'} onClick={onApply}>{t('Apply ({n})', { n: list.length })}</Btn>
      </div>
    </aside>
  );
}

export function RedactResult({ result, onClose }) {
  return (
    <Modal title={result.ok ? 'Redaction checked' : 'Please check the redaction'} onClose={onClose} footer={<Btn className="primary" onClick={onClose}>OK</Btn>}>
      <div className="stack">
        {result.ok
          ? <Callout kind="ok" icon={CheckCircle2} dismissible={false}>{t('{n} areas on {p} pages are removed for good. The text of the result was extracted again: nothing of the marked content is left. Save to keep it.', result)}</Callout>
          : <Callout kind="warn" icon={AlertTriangle} dismissible={false}>{t('Text was still found under {k} marks – those pages were flattened without a text layer. Check the result before you share it.', { k: result.leaks })}</Callout>}
        <p className="small muted">{t('Redacted pages are pictures with an invisible text layer of the remaining text. Annotations, form fields and alternative texts under the marks were removed too.')}</p>
      </div>
    </Modal>
  );
}

/* ---------------- compare ---------------- */
async function pageTexts(doc) {
  const out = [];
  for (let i = 1; i <= doc.numPages; i++) out.push((await (await doc.getPage(i)).getTextContent()).items.map((x) => x.str + (x.hasEOL ? '\n' : '')).join(''));
  return out;
}
async function renderTo(doc, n, width) {
  const page = await doc.getPage(n);
  const base = page.getViewport({ scale: 1 });
  const vp = page.getViewport({ scale: width / base.width });
  const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
  await page.render({ canvasContext: c.getContext('2d'), viewport: vp, background: '#ffffff' }).promise;
  return c;
}

/**
 * Side by side comparison of two PDFs (also the chunk's `Compare` export for host.compare – versions, other files):
 * page list with +/− counts, word diff of the text, and a visual overlay (red = only before, blue = only now).
 * a, b: { bytes, label } – a is the older one.
 */
export function CompareView({ a, b }) {
  const [docs, setDocs] = useState(null);
  const [err, setErr] = useState('');
  const [sel, setSel] = useState(0);
  const [view, setView] = useState('text');
  const canvas = useRef(null);
  const [diffInfo, setDiffInfo] = useState(null);
  useEffect(() => {
    let gone = false; let opened = [];
    (async () => {
      const A = await openDoc(a.bytes); const B = await openDoc(b.bytes);
      opened = [A, B];
      const pairs = comparePages(await pageTexts(A), await pageTexts(B));
      if (gone) return;
      setDocs({ a: A, b: B, pairs }); setSel(Math.max(0, pairs.findIndex((p) => !p.same)));
    })().catch((e) => { if (!gone) setErr(e.name === 'PasswordException' ? t('Protected PDFs cannot be compared – remove the protection first.') : e.message); });
    return () => { gone = true; opened.forEach((d) => d.destroy()); };
  }, [a, b]);
  const pair = docs && docs.pairs[sel];
  useEffect(() => {
    if (!pair || view !== 'visual' || !canvas.current) return undefined;
    let gone = false;
    (async () => {
      const W = 760;
      const ca = pair.a != null ? await renderTo(docs.a, pair.a + 1, W) : null;
      const cb = pair.b != null ? await renderTo(docs.b, pair.b + 1, W) : null;
      if (gone) return;
      const w = Math.max(ca ? ca.width : 0, cb ? cb.width : 0); const h = Math.max(ca ? ca.height : 0, cb ? cb.height : 0);
      const out = canvas.current; out.width = w; out.height = h;
      const grab = (c) => { const x = document.createElement('canvas'); x.width = w; x.height = h; const g = x.getContext('2d', { willReadFrequently: true }); g.fillStyle = '#fff'; g.fillRect(0, 0, w, h); if (c) g.drawImage(c, 0, 0); return g.getImageData(0, 0, w, h); };
      const A = grab(ca); const B = grab(cb);
      const d = pixelDiff(A.data, B.data, w, h);
      const ctx = out.getContext('2d'); const img = ctx.createImageData(w, h);
      for (let p = 0, q = 0; p < w * h; p++, q += 4) {
        const la = (A.data[q] + A.data[q + 1] + A.data[q + 2]) / 3; const lb = (B.data[q] + B.data[q + 1] + B.data[q + 2]) / 3;
        if (d.mask[p]) { if (la < lb) { img.data[q] = 220; img.data[q + 1] = 40; img.data[q + 2] = 50; } else { img.data[q] = 30; img.data[q + 1] = 110; img.data[q + 2] = 230; } } else { const v = 255 - (255 - lb) * 0.45; img.data[q] = v; img.data[q + 1] = v; img.data[q + 2] = v; }
        img.data[q + 3] = 255;
      }
      ctx.putImageData(img, 0, 0);
      setDiffInfo(d.ratio);
    })().catch((e) => setErr(e.message));
    return () => { gone = true; };
  }, [pair, view, docs]);
  if (err) return <Callout kind="err">{err}</Callout>;
  if (!docs) return <div className="of-pdf-compare-empty"><Progress indeterminate /><span className="small muted">{t('Comparing …')}</span></div>;
  const changed = docs.pairs.filter((p) => !p.same).length;
  return (
    <div className="of-pdf-compare">
      <div className="of-pdf-compare-list" role="listbox" aria-label="Pages">
        <div className="small muted of-pdf-compare-sum">{changed ? t('{n} of {m} pages differ', { n: changed, m: docs.pairs.length }) : t('No differences in the text.')}</div>
        {docs.pairs.map((p, i) => (
          <button key={i} type="button" role="option" aria-selected={i === sel} className={`of-pdf-compare-row${i === sel ? ' on' : ''}${p.same ? '' : ' diff'}`} onClick={() => setSel(i)}>
            <span>{p.a != null ? p.a + 1 : '–'} ↔ {p.b != null ? p.b + 1 : '–'}</span>
            {p.same ? <span className="small muted">{t('same')}</span> : <span className="small"><span className="of-pdf-add">+{p.added}</span> <span className="of-pdf-del">−{p.removed}</span></span>}
          </button>
        ))}
      </div>
      <div className="of-pdf-compare-main">
        <div className="row of-pdf-compare-bar">
          <Seg value={view} onChange={setView} options={[{ id: 'text', label: t('Text') }, { id: 'visual', label: t('Overlay') }]} />
          <span className="small muted of-pdf-ellipsis"><Keep>{`${a.label || ''} → ${b.label || ''}`}</Keep></span>
          <span className="grow" />
          {view === 'visual' && diffInfo != null && <span className="small muted">{t('{p} % of the page differ · red = only before, blue = only now', { p: (diffInfo * 100).toFixed(1) })}</span>}
        </div>
        {view === 'text' ? (
          <div className="of-pdf-difftext">{pair && pair.ops.map((o, i) => <span key={i} className={o.op > 0 ? 'ins' : o.op < 0 ? 'del' : ''}><Keep>{o.text}</Keep></span>)}</div>
        ) : <div className="of-pdf-diffcanvas"><canvas ref={canvas} /></div>}
      </div>
    </div>
  );
}

export function CompareDialog({ name, other, getBytes, pick, onClose }) {
  const [pair, setPair] = useState(null);
  const [busy, setBusy] = useState(false);
  const load = async (f) => {
    if (!f) return;
    setBusy(true);
    try { setPair({ a: { bytes: f.bytes, label: f.name }, b: { bytes: await getBytes(), label: name } }); } finally { setBusy(false); }
  };
  useEffect(() => { load(other); }, [other]); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <Modal title="Compare PDFs" subtitle="Text differences and a visual overlay, page by page" onClose={onClose} size="xl"
      footer={<><Btn icon={FileUp} busy={busy} onClick={async () => load(await pick())}>Other file …</Btn><span className="grow" /><Btn className="primary" onClick={onClose}>Close</Btn></>}>
      {pair ? <CompareView a={pair.a} b={pair.b} /> : <div className="of-pdf-compare-empty"><GitCompare size={40} /><Btn className="primary" icon={FileUp} busy={busy} onClick={async () => load(await pick())}>Choose a PDF …</Btn></div>}
    </Modal>
  );
}
