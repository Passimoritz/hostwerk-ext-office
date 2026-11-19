/*
 * Read-only view of VBA macros: vbaProject.bin (OLE compound file, read with SheetJS' CFB) → module names + source
 * (MS-OVBA decompression of the dir stream and module streams). Nothing is ever executed.
 */
import * as XLSX from 'xlsx';

export function decompress(buf, offset = 0) {
  if (buf[offset] !== 1) return new Uint8Array(0);
  const out = [];
  let i = offset + 1;
  while (i + 1 < buf.length) {
    const header = buf[i] | (buf[i + 1] << 8);
    const chunkEnd = Math.min(buf.length, i + 2 + (header & 0x0fff) + 1);
    i += 2;
    const start = out.length;
    if (!((header >> 15) & 1)) { for (let k = 0; k < 4096 && i < buf.length; k++) out.push(buf[i++]); continue; }
    while (i < chunkEnd) {
      const flag = buf[i++];
      for (let bit = 0; bit < 8 && i < chunkEnd; bit++) {
        if (!((flag >> bit) & 1)) { out.push(buf[i++]); continue; }
        const token = buf[i] | (buf[i + 1] << 8); i += 2;
        const d = out.length - start;
        let bc = 4; while ((1 << bc) < d) bc++;
        if (bc > 12) bc = 12;
        const lenMask = 0xffff >> bc;
        const len = (token & lenMask) + 3;
        const off = (token >> (16 - bc)) + 1;
        for (let k = 0; k < len; k++) out.push(out[out.length - off]);
      }
    }
  }
  return Uint8Array.from(out);
}

function decoder(cp) { try { return new TextDecoder(cp === 65001 ? 'utf-8' : `windows-${cp}`); } catch { return new TextDecoder('windows-1252'); } }

/** → [{ name, kind: 'module'|'class'|'document', code }] (empty when unreadable) */
export function readVba(bin) {
  try {
    const cfb = XLSX.CFB.read(bin, { type: 'array' });
    const find = (name) => { const i = cfb.FullPaths.findIndex((p) => p.replace(/\/$/, '').toLowerCase().endsWith(`/${name.toLowerCase()}`)); return i >= 0 ? cfb.FileIndex[i] : null; };
    const dirEntry = find('VBA/dir');
    if (!dirEntry) return [];
    const dir = decompress(Uint8Array.from(dirEntry.content));
    const dv = new DataView(dir.buffer);
    let p = 0; let cp = 1252;
    const mods = []; let cur = null;
    while (p + 6 <= dir.length) {
      const id = dv.getUint16(p, true); let size = dv.getUint32(p + 2, true); p += 6;
      if (id === 0x0009) size = 6; // PROJECTVERSION: size field lies
      const data = dir.subarray(p, p + size);
      if (id === 0x0003) cp = dv.getUint16(p, true);
      else if (id === 0x0019) { cur = { name: decoder(cp).decode(data), stream: null, offset: 0, kind: 'module' }; mods.push(cur); } else if (id === 0x001a && cur) cur.stream = decoder(cp).decode(data);
      else if (id === 0x0031 && cur) cur.offset = dv.getUint32(p, true);
      else if (id === 0x0022 && cur) cur.kind = 'class';
      p += size;
    }
    const dec = decoder(cp);
    return mods.map((m) => {
      const e = find(`VBA/${m.stream || m.name}`);
      let code = '';
      if (e) code = dec.decode(decompress(Uint8Array.from(e.content), m.offset)).replace(/^Attribute VB_.*\r?\n/gm, '');
      const kind = m.kind === 'class' && /^(ThisWorkbook|Sheet\d+|Tabelle\d+|DieseArbeitsmappe)$/i.test(m.name) ? 'document' : m.kind;
      return { name: m.name, kind, code };
    });
  } catch { return []; }
}
