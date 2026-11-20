/*
 * Docs co-editing (1.2): the saver is the server's leader (the same for everyone at once); a save of the room makes the
 * co-editors clean; becoming the leader with unsaved room changes makes this editor save them.
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as Y from 'yjs';
import { Awareness } from 'y-protocols/awareness';
import Document from '@tiptap/extension-document';
import Paragraph from '@tiptap/extension-paragraph';
import Text from '@tiptap/extension-text';
import { register } from 'node:module';

// collab.js imports Hostwerk's i18n (given by Hostwerk at runtime) – a tiny stand-in for the test
const I18N = 'data:text/javascript,' + encodeURIComponent('export const t = (s) => s; export const pl = (n, a, b) => (n === 1 ? a : b);');
register('data:text/javascript,' + encodeURIComponent(`export async function resolve(spec, ctx, next) { if (spec === '@hostwerk/i18n') return { url: ${JSON.stringify(I18N)}, shortCircuit: true }; return next(spec, ctx); }`));
const { joinCollab } = await import('../ui/src/editors/docs/collab.js');

/** a provider like lib/collab.js's OfficeProvider (events, hello, leader) without a websocket */
function fakeHost({ leader, fresh = true }) {
  const ls = new Map();
  let provider = null;
  return {
    get provider() { return provider; },
    collab: {
      connect: async ({ doc }) => {
        provider = {
          doc, awareness: new Awareness(doc), synced: Promise.resolve(true), hello: { leader, fresh }, leader,
          on: (e, fn) => { if (!ls.has(e)) ls.set(e, new Set()); ls.get(e).add(fn); return () => ls.get(e).delete(fn); },
          emit: (e, v) => { if (e === 'leader') provider.leader = v; (ls.get(e) || []).forEach((f) => f(v)); },
          destroy: () => {},
        };
        return provider;
      },
    },
  };
}
const ext = [Document, Paragraph, Text];
const docJson = { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Hallo' }] }] };

test('docs collab: the leader is the server\'s leader – not the smallest client id', async () => {
  const host = fakeHost({ leader: false });
  const s = await joinCollab(host, { extensions: ext, initialDoc: docJson, model: {} });
  // a second editor whose own client id would be the smallest: still not the saver
  s.awareness.setLocalStateField('user', { name: 'B' });
  assert.equal(s.isLeader(), false);
  host.provider.emit('leader', true);
  assert.equal(s.isLeader(), true);
  s.awareness.destroy();
});

test('docs collab: room saves make co-editors clean, a hand-over with unsaved changes makes the new leader save', async () => {
  const host = fakeHost({ leader: false });
  const s = await joinCollab(host, { extensions: ext, initialDoc: docJson, model: {} });
  const saved = []; const lead = [];
  s.onSaved((m) => saved.push(m.mtime)); s.onLeader((v) => lead.push(v));
  s.fragment.insert(0, [new Y.XmlText('neu')]); // an edit (remote or local)
  assert.equal(s.unsaved(), true);
  host.provider.emit('saved', { mtime: 123, by: 'Anna' });
  assert.deepEqual(saved, [123]);
  assert.equal(s.unsaved(), false, 'the room saved everything');
  s.fragment.insert(0, [new Y.XmlText('mehr')]);
  host.provider.emit('leader', true); // the old leader left
  assert.deepEqual(lead, [true]);
  assert.equal(s.unsaved(), true, 'the new leader has changes to save');
  s.awareness.destroy();
});

test('docs collab: the editor wires leader + saved to the shell\'s dirty state', async () => {
  const fs = await import('node:fs');
  const src = fs.readFileSync(new URL('../ui/src/editors/docs/Editor.jsx', import.meta.url), 'utf8');
  assert.match(src, /collab\.onSaved\(\(\) => \{ if \(!collab\.isLeader\(\)\) \{ dirtyRef\.current = false; if \(host\.dirty\) host\.dirty\(false\)/);
  assert.match(src, /collab\.onLeader\(lead\)/);
  assert.match(src, /collab\.unsaved\(\)/);
});
