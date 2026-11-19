/* ZIP archives: file list (folders, sizes, dates, compression) with search and single-file download. */
import { useEffect, useMemo, useState } from 'react';
import JSZip from 'jszip';
import { Folder, File as FileIcon, Download, Search, Archive, Lock } from 'lucide-react';
import { Loading, Callout, TileEmpty } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { ActionRow, fmtSize, fmtDate } from './common.jsx';
import { FORMATS, extOf } from '../../lib/formats.js';
import { Keep } from '../_shared/keep.jsx';

export async function listZip(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const out = [];
  zip.forEach((path, f) => {
    // eslint-disable-next-line no-underscore-dangle
    const d = f._data || {};
    out.push({ path, dir: f.dir, size: d.uncompressedSize || 0, packed: d.compressedSize || 0, date: f.date, f });
  });
  out.sort((a, b) => a.path.localeCompare(b.path, undefined, { numeric: true }));
  return { zip, entries: out };
}

export default function ZipView({ file, bytes, host }) {
  const [state, setState] = useState(null);
  const [err, setErr] = useState(null);
  const [q, setQ] = useState('');
  useEffect(() => { let gone = false; listZip(bytes).then((s) => !gone && setState(s)).catch((e) => !gone && setErr(e)); return () => { gone = true; }; }, [bytes]);
  const files = useMemo(() => (state ? state.entries.filter((e) => !e.dir && (!q || e.path.toLowerCase().includes(q.toLowerCase()))) : []), [state, q]);
  const total = state ? state.entries.reduce((s, e) => s + e.size, 0) : 0;
  useEffect(() => { if (state) host.status(t('{n} files · {size} unpacked', { n: state.entries.filter((e) => !e.dir).length, size: fmtSize(total) })); }, [state]); // eslint-disable-line react-hooks/exhaustive-deps
  if (err) return <div className="of-view-pad"><Callout kind="err" icon={/encrypt/i.test(err.message) ? Lock : undefined} title={t('This archive cannot be opened')}>{/encrypt/i.test(err.message) ? t('The archive is password-protected – unpack it on your computer.') : err.message}</Callout><ActionRow file={file} bytes={bytes} host={host} /></div>;
  if (!state) return <Loading text={t('Reading archive …')} />;
  const get = async (e) => { const data = await e.f.async('uint8array'); const name = e.path.split('/').pop(); host.download(data, name, (FORMATS[extOf(name)] || {}).mime || 'application/octet-stream'); };
  return (
    <div className="of-view-zip">
      <div className="of-view-toolbar">
        <h2 className="of-view-title"><Archive size={18} /> <Keep>{file.name}</Keep></h2>
        <span className="small muted">{t('{n} files', { n: state.entries.filter((e) => !e.dir).length })} · {fmtSize(total)} → {fmtSize(bytes.length)}</span>
        <span className="grow" />
        <span className="of-view-search"><Search size={14} /><input className="input" value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search in archive" aria-label="Search in archive" /></span>
      </div>
      <div className="of-view-tablewrap">
        <table className="table">
          <thead><tr><th>{t('Name')}</th><th className="num">{t('Size')}</th><th className="num">{t('Packed')}</th><th>{t('Modified')}</th><th aria-label="Actions" /></tr></thead>
          <tbody>
            {files.slice(0, 5000).map((e) => {
              const parts = e.path.split('/'); const name = parts.pop();
              return (
                <tr key={e.path}>
                  <td><span className="of-view-zipname"><FileIcon size={14} className="muted" />{parts.length > 0 && <span className="muted small"><Folder size={12} /> <Keep>{`${parts.join('/')}/`}</Keep></span>}<span><Keep>{name}</Keep></span></span></td>
                  <td className="num">{fmtSize(e.size)}</td>
                  <td className="num muted">{e.size ? `${Math.round((e.packed / e.size) * 100)} %` : ''}</td>
                  <td className="muted small">{fmtDate(e.date)}</td>
                  <td><button type="button" className="btn xs ghost icon" title={t('Download {name}', { name })} aria-label={t('Download {name}', { name })} onClick={() => get(e)}><Download size={12} /></button></td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {!files.length && <TileEmpty icon={Search} text={q ? t('Nothing found.') : t('This archive is empty.')} />}
      </div>
      <ActionRow file={file} bytes={bytes} host={host} />
    </div>
  );
}
