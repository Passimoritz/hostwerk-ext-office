/*
 * Slides editor chunk (ARCHITECTURE.md "Editor chunk contract"): presentations – pptx/pptm/ppsx/potx… (own OOXML
 * reader/writer, lossless patching of the original package), odp/otp/fodp (own ODF reader/writer), ppt/pps/pot
 * (built-in 97–2003 reader), export to PDF/PNG, slide show with presenter view.
 */
import { useEffect, useRef, useState } from 'react';
import Editor from './Editor.jsx';
import SlideView from './render/SlideView.jsx';
import { readAny, writeAny } from './io/index.js';
import { buildTemplate } from './templates.js';
import { docToText } from './model.js';
import { S } from './units.js';
import { t, pl } from '@hostwerk/i18n';
import './style.css';
import './style11.css';

async function create({ ext = 'pptx', template = 'blank', settings = {} } = {}) {
  const ss = settings.slides || {};
  const doc = buildTemplate(template || 'blank', { size: ss.size || '16:9', theme: ss.theme || 'hostwerk', font: ss.font, author: (settings.general || {}).author || '' });
  if (ss.transition && ss.transition !== 'none') {
    const type = ss.transition === 'slide' ? 'push' : ss.transition;
    doc.slides = doc.slides.map((s) => ({ ...s, transition: { ...(s.transition || {}), type: s.transition && s.transition.type !== 'none' ? s.transition.type : type } }));
  }
  if (ss.loop || ss.autoAdvance) doc.show = { ...doc.show, loop: !!ss.loop, autoAdvance: ss.autoAdvance || 0 };
  const e = String(ext || 'pptx').toLowerCase();
  const fmt = ['odp', 'otp', 'fodp'].includes(e) ? (e === 'otp' ? 'otp' : 'odp') : ['pptx', 'pptm', 'ppsx', 'potx', 'ppsm', 'potm'].includes(e) ? e : 'pptx';
  return writeAny(doc, fmt);
}

async function toText(bytes, ext) {
  const { doc } = await readAny(bytes, ext);
  return docToText(doc);
}

function Preview({ bytes, ext }) {
  const [doc, setDoc] = useState(null);
  const [err, setErr] = useState(false);
  const box = useRef(null);
  const [w, setW] = useState(320);
  useEffect(() => {
    let alive = true;
    readAny(bytes, ext).then((r) => { if (alive) setDoc(r.doc); }).catch(() => alive && setErr(true));
    return () => { alive = false; };
  }, [bytes, ext]);
  useEffect(() => {
    if (!box.current) return undefined;
    const m = () => { if (box.current) setW(box.current.clientWidth || 320); };
    m(); const ro = new ResizeObserver(m); ro.observe(box.current);
    return () => ro.disconnect();
  }, []);
  const slide = doc && (doc.slides.find((s) => !s.hidden) || doc.slides[0]);
  return (
    <div className="of-slides-preview" ref={box}>
      {err && <div className="small muted">{t('Preview not available.')}</div>}
      {slide && <SlideView doc={doc} slide={slide} index={doc.slides.indexOf(slide)} scale={w / S(doc.size.w)} mode="thumb" />}
      {doc && <div className="small muted of-slides-preview-meta">{pl(doc.slides.length, '{n} slide', '{n} slides')}</div>}
    </div>
  );
}

export default { kind: 'slides', Editor, create, toText, Preview };
