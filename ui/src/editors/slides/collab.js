/*
 * Co-editing of a presentation over host.collab (Yjs, element level). The shared Y.Doc holds
 *   Y.Map 'deck'   key → JSON: 'base' (size, theme, layouts, sections, footer, show, meta …), 'order' (slide ids)
 *   Y.Map 'slides' slide id → JSON of the slide without elements, plus 'els' (element ids in z-order)
 *   Y.Map 'els'    `${slideId}/${elId}` → JSON of one element
 * Local edits are diffed against the last synced doc (unchanged objects are skipped by identity, so typing in one
 * text box sends one element). Remote changes rebuild the doc. Awareness: { user: { name, color }, slide, sel }.
 * The Y module comes with the session (one Yjs instance for the whole app – nothing bundled here).
 */
const strCache = new WeakMap();
const js = (o) => { if (o && typeof o === 'object') { let s = strCache.get(o); if (s === undefined) { s = JSON.stringify(o); strCache.set(o, s); } return s; } return JSON.stringify(o); };
const baseOf = (doc) => { const { slides, ...rest } = doc; return rest; }; // eslint-disable-line no-unused-vars
const slideMeta = (s) => { const { elements, ...rest } = s; return { ...rest, els: elements.map((e) => e.id) }; };

export const COLORS = ['#E5484D', '#3E63DD', '#30A46C', '#F76B15', '#8E4EC6', '#12A594', '#D6409F', '#AD7F58'];
export const colorFor = (name) => { let h = 0; for (const c of String(name || '')) h = (h * 31 + c.charCodeAt(0)) >>> 0; return COLORS[h % COLORS.length]; };

/** write the whole doc into empty shared state */
export function seed(Y, ydoc, doc) {
  const deck = ydoc.getMap('deck'); const slides = ydoc.getMap('slides'); const els = ydoc.getMap('els');
  ydoc.transact(() => {
    deck.set('base', js(baseOf(doc)));
    deck.set('order', JSON.stringify(doc.slides.map((s) => s.id)));
    for (const s of doc.slides) { slides.set(s.id, JSON.stringify(slideMeta(s))); for (const e of s.elements) els.set(`${s.id}/${e.id}`, js(e)); }
  }, 'local');
}
/** the doc of the shared state (null when empty) */
export function readShared(ydoc) {
  const deck = ydoc.getMap('deck'); const slides = ydoc.getMap('slides'); const els = ydoc.getMap('els');
  const b = deck.get('base'); const o = deck.get('order');
  if (!b || !o) return null;
  const base = JSON.parse(b);
  const order = JSON.parse(o);
  const out = [];
  for (const id of order) {
    const m = slides.get(id);
    if (!m) continue;
    const meta = JSON.parse(m);
    const { els: ids, ...rest } = meta;
    const elements = (ids || []).map((eid) => els.get(`${id}/${eid}`)).filter(Boolean).map((x) => JSON.parse(x));
    out.push({ ...rest, elements });
  }
  return { ...base, slides: out };
}
/** push the difference prev → next into the shared state */
export function pushDiff(ydoc, prev, next) {
  if (prev === next) return;
  const deck = ydoc.getMap('deck'); const slides = ydoc.getMap('slides'); const els = ydoc.getMap('els');
  ydoc.transact(() => {
    const pb = prev ? js(baseOf(prev)) : null; const nb = js(baseOf(next));
    if (pb !== nb) deck.set('base', nb);
    const po = prev ? prev.slides.map((s) => s.id).join(',') : ''; const no = next.slides.map((s) => s.id).join(',');
    if (po !== no) deck.set('order', JSON.stringify(next.slides.map((s) => s.id)));
    const prevById = new Map((prev ? prev.slides : []).map((s) => [s.id, s]));
    for (const s of next.slides) {
      const p = prevById.get(s.id);
      prevById.delete(s.id);
      if (p === s) continue;
      const meta = JSON.stringify(slideMeta(s));
      if (!p || JSON.stringify(slideMeta(p)) !== meta) slides.set(s.id, meta);
      const pe = new Map(((p && p.elements) || []).map((e) => [e.id, e]));
      for (const e of s.elements) { const old = pe.get(e.id); pe.delete(e.id); if (old !== e && (!old || js(old) !== js(e))) els.set(`${s.id}/${e.id}`, js(e)); }
      for (const id of pe.keys()) els.delete(`${s.id}/${id}`);
    }
    for (const [id, s] of prevById) { slides.delete(id); for (const e of s.elements) els.delete(`${id}/${e.id}`); }
  }, 'local');
}

/**
 * connect an editor: session from host.collab.join; getDoc() → current doc; setDoc(doc) applies a remote doc;
 * onPeers(list) for presence. Returns { push(doc), presence(state), close() }.
 */
export function connect(session, { getDoc, setDoc, onPeers, me }) {
  const { ydoc, awareness } = session;
  let synced = getDoc();
  const shared = readShared(ydoc);
  if (shared && shared.slides.length) { synced = shared; setDoc(shared, true); } else seed(null, ydoc, synced);
  let timer = null;
  const onUpdate = (ev, tr) => {
    if (tr && tr.origin === 'local') return;
    clearTimeout(timer);
    timer = setTimeout(() => { const d = readShared(ydoc); if (d) { synced = d; setDoc(d, false); } }, 16);
  };
  ydoc.on('afterTransaction', (tr) => { if (tr.origin !== 'local' && tr.changed && tr.changed.size) onUpdate(null, tr); });
  const peers = () => {
    if (!awareness) return;
    const list = [];
    awareness.getStates().forEach((st, id) => { if (id !== awareness.clientID && st && st.user) list.push({ id, ...st }); });
    onPeers(list);
  };
  if (awareness) { awareness.setLocalStateField('user', { name: me.name, color: me.color }); awareness.on('change', peers); peers(); }
  return {
    push(doc) { if (doc && doc !== synced) { pushDiff(ydoc, synced, doc); synced = doc; } },
    presence(state) { if (awareness) for (const [k, v] of Object.entries(state)) awareness.setLocalStateField(k, v); },
    close() { clearTimeout(timer); if (awareness) awareness.off('change', peers); session.close(); },
  };
}
