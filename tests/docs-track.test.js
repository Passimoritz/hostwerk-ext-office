/*
 * Docs track changes (1.2): accepting a deleted paragraph mark joins the paragraph with the next one (like Word);
 * rejecting it keeps both; an accepted inserted mark keeps the split.
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { register } from 'node:module';
import { getSchema } from '@tiptap/core';
import { EditorState } from '@tiptap/pm/state';
import Document from '@tiptap/extension-document';
import Paragraph from '@tiptap/extension-paragraph';
import Text from '@tiptap/extension-text';

const I18N = 'data:text/javascript,' + encodeURIComponent('export const t = (s) => s; export const pl = (n, a, b) => (n === 1 ? a : b);');
register('data:text/javascript,' + encodeURIComponent(`export async function resolve(spec, ctx, next) { if (spec === '@hostwerk/i18n') return { url: ${JSON.stringify(I18N)}, shortCircuit: true }; return next(spec, ctx); }`));
const { applyChanges, changeRanges, Insertion, Deletion, FormatChange } = await import('../ui/src/editors/docs/extensions/track.js');
const { RevisionAttrs } = await import('../ui/src/editors/docs/extensions/rich.js');

const schema = getSchema([Document, Paragraph, Text, RevisionAttrs, Insertion, Deletion, FormatChange]);
const rev = { id: '1', author: 'Anna', date: '2026-11-20T10:00:00Z' };
const p = (text, attrs = {}) => ({ type: 'paragraph', attrs, content: text ? [{ type: 'text', text }] : [] });
const run = (json, accept) => {
  const state = EditorState.create({ schema, doc: schema.nodeFromJSON(json) });
  const tr = applyChanges(state.tr, changeRanges(state.doc), accept);
  return tr.doc.toJSON().content.map((n) => ({ text: (n.content || []).map((x) => x.text).join(''), pmark: n.attrs && n.attrs.pmark }));
};

test('accept a deleted paragraph mark → the paragraphs join; reject → both stay, mark gone', () => {
  const doc = { type: 'doc', content: [p('Eins ', { pmark: { type: 'del', ...rev } }), p('zwei ', { pmark: { type: 'del', ...rev } }), p('drei'), p('vier')] };
  assert.deepEqual(run(doc, true), [{ text: 'Eins zwei drei', pmark: null }, { text: 'vier', pmark: null }]);
  assert.deepEqual(run(doc, false), [{ text: 'Eins ', pmark: null }, { text: 'zwei ', pmark: null }, { text: 'drei', pmark: null }, { text: 'vier', pmark: null }]);
});

test('an inserted paragraph mark: accept keeps the split, reject joins', () => {
  const doc = { type: 'doc', content: [p('Teil A', { pmark: { type: 'ins', ...rev } }), p('Teil B')] };
  assert.deepEqual(run(doc, true).map((x) => x.text), ['Teil A', 'Teil B']);
  assert.deepEqual(run(doc, false).map((x) => x.text), ['Teil ATeil B']);
});
