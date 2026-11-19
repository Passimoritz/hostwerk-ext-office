/*
 * VBA projects (vbaProject.bin in docm/dotm, the Macros storage in .doc): list the modules and decode their source
 * (MS-OVBA run-length decompression of the "dir" and module streams). Read-only – nothing is ever executed.
 */
import { readCfb } from '../../_shared/cfb.js';

export function decompressVba(u, start = 0) {
  const out = [];
  if (u[start] !== 1) throw new Error('Not a VBA compressed container');
  let p = start + 1;
  while (p + 2 <= u.length) {
    const header = u[p] | (u[p + 1] << 8);
    p += 2;
    const size = (header & 0x0fff) + 3;
    const compressed = (header >> 15) & 1;
    const end = Math.min(u.length, p - 2 + size);
    const chunkStart = out.length;
    if (!compressed) { for (let k = 0; k < 4096 && p < u.length; k++) out.push(u[p++]); continue; }
    while (p < end) {
      const flags = u[p++];
      for (let bit = 0; bit < 8 && p < end; bit++) {
        if (!((flags >> bit) & 1)) { out.push(u[p++]); continue; }
        const token = u[p] | (u[p + 1] << 8);
        p += 2;
        const diff = out.length - chunkStart;
        let bitCount = 4;
        while ((1 << bitCount) < diff) bitCount++;
        bitCount = Math.max(bitCount, 4);
        const lengthMask = 0xffff >> bitCount;
        const len = (token & lengthMask) + 3;
        const offset = (token >> (16 - bitCount)) + 1;
        const from = out.length - offset;
        for (let k = 0; k < len; k++) out.push(out[from + k]);
      }
    }
    p = end;
  }
  return new Uint8Array(out);
}

function decoder(cp) { try { return new TextDecoder(cp === 65001 ? 'utf-8' : `windows-${cp || 1252}`); } catch { return new TextDecoder('windows-1252'); } }

/** modules of a VBA project: [{ name, type: 'module'|'class'|'document', source }] */
export function vbaModules(bytesOrCfb, prefix = '') {
  const cfb = bytesOrCfb.streams ? bytesOrCfb : readCfb(bytesOrCfb);
  const find = (p) => {
    const want = (prefix ? `${prefix}/${p}` : p).toLowerCase();
    for (const [k, v] of cfb.streams) if (k.toLowerCase() === want) return v;
    for (const [k, v] of cfb.streams) if (k.toLowerCase().endsWith(`/${p.toLowerCase()}`)) return v;
    return null;
  };
  const dirRaw = find('VBA/dir');
  if (!dirRaw) return [];
  const dir = decompressVba(dirRaw);
  const dv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  let p = 0; let cp = 1252;
  const mods = []; let cur = null;
  while (p + 6 <= dir.length) {
    const id = dv.getUint16(p, true); const size = dv.getUint32(p + 2, true);
    p += 6;
    const len = id === 0x0009 ? 6 : size;
    const data = dir.subarray(p, p + len);
    p += len;
    if (id === 0x0003) cp = data[0] | (data[1] << 8);
    else if (id === 0x0019) { cur = { name: decoder(cp).decode(data), stream: null, offset: 0, type: 'module' }; mods.push(cur); }
    else if (id === 0x0047 && cur) cur.name = new TextDecoder('utf-16le').decode(data) || cur.name;
    else if (id === 0x001a && cur) cur.stream = decoder(cp).decode(data);
    else if (id === 0x0032 && cur) cur.stream = new TextDecoder('utf-16le').decode(data) || cur.stream;
    else if (id === 0x0031 && cur) cur.offset = new DataView(data.buffer, data.byteOffset, 4).getUint32(0, true);
    else if (id === 0x0022 && cur) cur.type = 'class';
    else if (id === 0x002b) cur = null;
  }
  return mods.map((m) => {
    let source = '';
    try {
      const s = find(`VBA/${m.stream || m.name}`);
      if (s) source = decoder(cp).decode(decompressVba(s, m.offset)).replace(/^Attribute VB_[^\n]*\n/gm, '');
    } catch { source = ''; }
    const type = m.type === 'class' && /^(ThisDocument|DieseArbeitsmappe|ThisWorkbook|Tabelle\d+|Sheet\d+)$/i.test(m.name) ? 'document' : m.type;
    return { name: m.name, type, source };
  });
}
