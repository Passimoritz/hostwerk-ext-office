/*
 * PowerPoint animation timing (p:timing) ⇄ the model's animations (see anim.js ANIM_EFFECTS):
 *   { id, el, cls: 'entr'|'emph'|'exit'|'path', effect, dir?, trigger, duration, delay, path? ('M 0 0 L 0.25 0 E',
 *     fractions of the slide), scale? (grow), raw? (an effect we can't edit: its p:par XML, re-emitted as it is) }
 * timingXml() writes PowerPoint's structure (main sequence → click groups → after-chains → effects); animationsOf()
 * reads it back. Pure (Node tests).
 */
import { serialize } from '../../_shared/xmltree.js';
import { PRESETS, presetOf } from '../anim.js';

const I = (n) => Math.round(Number(n) || 0);
const ln = (n) => n.slice(n.indexOf(':') + 1);
const KS = (el, name) => (el && el.c ? el.c.filter((x) => typeof x !== 'string' && (!name || ln(x.n) === name)) : []);
const K = (el, name) => KS(el, name)[0];
function FIND(el, name, out = []) { for (const x of (el && el.c) || []) { if (typeof x === 'string') continue; if (ln(x.n) === name) out.push(x); FIND(x, name, out); } return out; }

const FLY_SUB = { b: 4, d: 4, t: 1, u: 1, l: 8, r: 2 };
const SUB_DIR = { 4: 'b', 1: 't', 8: 'l', 2: 'r', 3: 't', 6: 'b', 9: 'l', 12: 'b' };

/* ---------------------------------------------------------------- write */
function behaviours(a, tgt, N) {
  const dur = Math.max(1, I(a.duration || 500));
  const cls = a.cls || 'entr';
  const vis = (val, delay = 0) => `<p:set><p:cBhvr><p:cTn id="${N()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${I(delay)}"/></p:stCondLst></p:cTn>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${val}"/></p:to></p:set>`;
  const fx = (dir, filter) => `<p:animEffect transition="${dir}" filter="${filter}"><p:cBhvr><p:cTn id="${N()}" dur="${dur}"/>${tgt}</p:cBhvr></p:animEffect>`;
  const anim = (attr, v0, v1, additive) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr${additive ? ' additive="base"' : ''}><p:cTn id="${N()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${v0}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${v1}"/></p:val></p:tav></p:tavLst></p:anim>`;
  const e = a.effect;
  if (cls === 'entr' || cls === 'exit') {
    const into = cls === 'entr';
    const sub = e === 'fly' ? FLY_SUB[a.dir] || 4 : e === 'wipe' ? 4 : e === 'zoom' ? 16 : e === 'split' ? 21 : e === 'wheel' ? 1 : 0;
    const out = into ? '' : vis('hidden', dur - 1);
    let b = '';
    if (e === 'appear') return { sub, xml: into ? vis('visible') : vis('hidden') };
    if (e === 'fade') b = fx(into ? 'in' : 'out', 'fade');
    else if (e === 'wipe') b = fx(into ? 'in' : 'out', into ? 'wipe(down)' : 'wipe(up)');
    else if (e === 'split') b = fx(into ? 'in' : 'out', into ? 'barn(outVertical)' : 'barn(inVertical)');
    else if (e === 'wheel') b = fx(into ? 'in' : 'out', 'wheel(1)');
    else if (e === 'float') b = (into ? anim('ppt_y', '#ppt_y+.1', '#ppt_y', true) : anim('ppt_y', '#ppt_y', '#ppt_y-.1', true)) + fx(into ? 'in' : 'out', 'fade');
    else if (e === 'fly') {
      const horiz = sub === 8 || sub === 2;
      const far = { 4: '1+#ppt_h/2', 1: '0-#ppt_h/2', 8: '0-#ppt_w/2', 2: '1+#ppt_w/2' }[sub];
      const ax = horiz ? (into ? [far, '#ppt_x'] : ['#ppt_x', far]) : ['#ppt_x', '#ppt_x'];
      const ay = horiz ? ['#ppt_y', '#ppt_y'] : (into ? [far, '#ppt_y'] : ['#ppt_y', far]);
      b = anim('ppt_x', ax[0], ax[1], true) + anim('ppt_y', ay[0], ay[1], true);
    } else if (e === 'zoom') {
      b = (into ? anim('ppt_w', '0', '#ppt_w') + anim('ppt_h', '0', '#ppt_h') : anim('ppt_w', '#ppt_w', '0') + anim('ppt_h', '#ppt_h', '0')) + fx(into ? 'in' : 'out', 'fade');
    } else b = fx(into ? 'in' : 'out', 'fade');
    return { sub, xml: (into ? vis('visible') : '') + b + out };
  }
  if (cls === 'emph') {
    if (e === 'spin') return { sub: 0, xml: `<p:animRot by="${I((a.angle || 360) * 60000)}"><p:cBhvr><p:cTn id="${N()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>` };
    if (e === 'grow') { const s = I((a.scale || 1.5) * 100000); return { sub: 0, xml: `<p:animScale><p:cBhvr><p:cTn id="${N()}" dur="${dur}" fill="hold"/>${tgt}</p:cBhvr><p:by x="${s}" y="${s}"/></p:animScale>` }; }
    if (e === 'pulse') return { sub: 0, xml: `<p:animScale><p:cBhvr><p:cTn id="${N()}" dur="${I(dur / 2)}" autoRev="1" fill="hold"/>${tgt}</p:cBhvr><p:by x="105000" y="105000"/></p:animScale>` };
    if (e === 'teeter') return { sub: 0, xml: `<p:animRot by="240000"><p:cBhvr><p:cTn id="${N()}" dur="${I(dur / 4)}" autoRev="1" repeatCount="2000" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>` };
    if (e === 'transparency') return { sub: 0, xml: `<p:set><p:cBhvr><p:cTn id="${N()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${(a.opacity ?? 0.5)}"/></p:to></p:set>` };
  }
  if (cls === 'path') {
    const path = a.path || 'M 0 0 L 0.25 0 E';
    const pts = (path.match(/-?\d*\.?\d+(?:e-?\d+)?/gi) || []).length / 2;
    return { sub: 0, xml: `<p:animMotion origin="layout" path="${path}" pathEditMode="relative" rAng="0" ptsTypes="${'A'.repeat(Math.max(1, Math.round(pts)))}"><p:cBhvr><p:cTn id="${N()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:rCtr x="0" y="0"/></p:animMotion>` };
  }
  return { sub: 0, xml: vis('visible') };
}

/** p:timing for a slide. spids: element id → shape id; textIds: elements with text (build list) */
export function timingXml(anims, spids, textIds) {
  const list = (anims || []).filter((a) => spids[a.el] && (a.raw || presetOf(a)));
  if (!list.length) return '';
  let id = 3;
  const N = () => id++;
  const groups = [];
  list.forEach((a, i) => {
    const trig = a.trigger || 'click';
    if (trig === 'click' || !groups.length) groups.push({ start: i === 0 && trig !== 'click' ? 'auto' : 'click', chains: [[a]] });
    else if (trig === 'after') groups[groups.length - 1].chains.push([a]);
    else groups[groups.length - 1].chains[groups[groups.length - 1].chains.length - 1].push(a);
  });
  const grp = {}; const builds = [];
  const effect = (a, nodeType) => {
    const sp = spids[a.el];
    const g = grp[sp] = (grp[sp] ?? -1) + 1;
    if (textIds.has(a.el)) builds.push([sp, g]);
    if (a.raw) {
      // re-emit an effect we can't edit: fresh cTn ids, its shape's (new) id, our trigger/delay
      let x = a.raw.replace(/(<p:cTn\b[^>]*?\sid=")\d+"/g, (m, p) => `${p}${N()}"`).replace(/(<p:spTgt\s+spid=")\d+"/g, `$1${sp}"`);
      x = x.replace(/(<p:cTn\b[^>]*?)\snodeType="[^"]*"/, `$1 nodeType="${nodeType}"`).replace(/(<p:cTn\b[^>]*?)\sgrpId="\d+"/, `$1 grpId="${g}"`);
      return x;
    }
    const tgt = `<p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl>`;
    const b = behaviours(a, tgt, N);
    const pre = presetOf(a);
    return `<p:par><p:cTn id="${N()}" presetID="${pre.id}" presetClass="${a.cls || 'entr'}" presetSubtype="${b.sub}" fill="hold" grpId="${g}" nodeType="${nodeType}"><p:stCondLst><p:cond delay="${I(a.delay || 0)}"/></p:stCondLst><p:childTnLst>${b.xml}</p:childTnLst></p:cTn></p:par>`;
  };
  const gx = groups.map((g) => {
    let offset = 0;
    const chains = g.chains.map((ch, ci) => {
      const nodes = ch.map((a, k) => effect(a, ci === 0 && k === 0 ? (g.start === 'click' ? 'clickEffect' : ch[0].trigger === 'after' ? 'afterEffect' : 'withEffect') : k === 0 ? 'afterEffect' : 'withEffect')).join('');
      const x = `<p:par><p:cTn id="${N()}" fill="hold"><p:stCondLst><p:cond delay="${offset}"/></p:stCondLst><p:childTnLst>${nodes}</p:childTnLst></p:cTn></p:par>`;
      offset += Math.max(...ch.map((a) => I(a.delay || 0) + I(a.duration || 500)));
      return x;
    }).join('');
    const cond = g.start === 'click' ? '<p:cond delay="indefinite"/>' : '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>';
    return `<p:par><p:cTn id="${N()}" fill="hold"><p:stCondLst>${cond}</p:stCondLst><p:childTnLst>${chains}</p:childTnLst></p:cTn></p:par>`;
  }).join('');
  const bld = builds.map(([s, g]) => `<p:bldP spid="${s}" grpId="${g}"/>`).join('');
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${gx}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>${bld ? `<p:bldLst>${bld}</p:bldLst>` : ''}</p:timing>`;
}

/* ---------------------------------------------------------------- read */
/** animations of a p:timing node; idMap: shape id → element id; uid: id maker; flag(key) for things we can't keep */
export function animationsOf(timing, idMap, uid, flag = () => {}) {
  const out = [];
  if (FIND(timing, 'cTn').some((c) => c.a.nodeType === 'interactiveSeq')) flag('animations');
  const seq = FIND(timing, 'cTn').find((c) => c.a.nodeType === 'mainSeq');
  if (!seq) return out;
  const pars = [];
  const walkPars = (el) => { for (const x of KS(el)) { if (ln(x.n) === 'par') { const c = K(x, 'cTn'); if (c && c.a.presetClass) { pars.push([x, c]); continue; } } walkPars(x); } };
  walkPars(seq);
  for (const [par, c] of pars) {
    const cls = c.a.presetClass;
    const tgt = FIND(c, 'spTgt')[0];
    const el = tgt && idMap[tgt.a.spid];
    if (!el) continue;
    const beh = FIND(c, 'cTn').find((x) => x !== c && x.a.dur && x.a.dur !== '1' && x.a.dur !== 'indefinite');
    const cond = K(K(c, 'stCondLst'), 'cond');
    const trigger = c.a.nodeType === 'withEffect' ? 'with' : c.a.nodeType === 'afterEffect' ? 'after' : 'click';
    const base = { id: uid('a'), el, cls, trigger, duration: beh ? I(beh.a.dur) || 500 : 500, delay: cond && cond.a.delay && cond.a.delay !== 'indefinite' ? I(cond.a.delay) : 0 };
    const pre = PRESETS.find((p) => p.cls === cls && p.id === I(c.a.presetID) && (p.match ? p.match(c, FIND) : true));
    const one = FIND(timing, 'spTgt').filter((x) => x.a.spid === tgt.a.spid).length;
    // several targets inside one effect, or text-level targets (paragraph builds) → keep as it is
    const multi = FIND(c, 'spTgt').some((x) => x.a.spid !== tgt.a.spid) || FIND(c, 'txEl').length > 0;
    if (!pre || multi || one < 1) { out.push({ ...base, effect: 'raw', presetID: I(c.a.presetID), raw: serialize(par) }); continue; }
    const a = { ...base, effect: pre.effect };
    if (pre.effect === 'appear' || (cls === 'exit' && pre.effect === 'appear')) a.duration = 1;
    if (pre.effect === 'fly') a.dir = SUB_DIR[I(c.a.presetSubtype)] || 'b';
    if (cls === 'path') { const m = FIND(c, 'animMotion')[0]; a.path = (m && m.a.path) || 'M 0 0 L 0.25 0 E'; }
    if (pre.effect === 'grow') { const by = FIND(c, 'by')[0]; if (by && by.a.x) a.scale = I(by.a.x) / 100000; }
    if (pre.effect === 'spin') { const r = FIND(c, 'animRot')[0]; if (r && r.a.by) a.angle = I(r.a.by) / 60000; }
    if (pre.effect === 'pulse' && beh) a.duration = I(beh.a.dur) * 2 || 500;
    if (pre.effect === 'teeter' && beh) a.duration = I(beh.a.dur) * 4 || 1000;
    if (pre.effect === 'transparency') { const v = FIND(c, 'strVal').find((x) => /^[\d.]+$/.test(x.a.val || '')); if (v) a.opacity = parseFloat(v.a.val); }
    out.push(a);
  }
  return out;
}
