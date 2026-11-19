'use strict';
/* JSON files in the data folder: read with fallback, write atomically (tmp + rename). */
const fs = require('fs');
const path = require('path');

function readJson(file, fallback) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return typeof fallback === 'function' ? fallback() : fallback; }
}
function writeJson(file, value) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = `${file}.${process.pid}.${Date.now().toString(36)}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(value));
  try { fs.renameSync(tmp, file); } catch {
    // Windows: target briefly locked by a virus scanner – write directly
    try { fs.writeFileSync(file, JSON.stringify(value)); } finally { fs.rmSync(tmp, { force: true }); }
  }
}
/** Write a file atomically (bytes) – keeps the old content if writing fails half way */
function writeFileAtomic(file, data) {
  const tmp = path.join(path.dirname(file), `.~office-${process.pid}-${Date.now().toString(36)}.tmp`);
  fs.writeFileSync(tmp, data);
  try { fs.renameSync(tmp, file); } catch {
    try { fs.copyFileSync(tmp, file); } finally { fs.rmSync(tmp, { force: true }); }
  }
}
const newId = () => `${Date.now().toString(36)}${require('crypto').randomBytes(5).toString('hex')}`;

module.exports = { readJson, writeJson, writeFileAtomic, newId };
