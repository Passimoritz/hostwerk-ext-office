import test from 'node:test';
import assert from 'node:assert/strict';
import { nextToCollapse, nextToExpand, levelForWidth, firstIcon } from '../ui/src/editors/_shared/ribbonfit.js';
import { sameContent } from '../ui/src/editors/slides/collab.js';

test('ribbon fit: the right-most group collapses first, kept groups never', () => {
  const g = [{ name: 'Slides' }, { name: 'Font' }, { name: 'Paragraph' }, { name: 'Designer', keep: true }];
  assert.equal(nextToCollapse(g, []), 'Paragraph');
  assert.equal(nextToCollapse(g, ['Paragraph']), 'Font');
  assert.equal(nextToCollapse(g, ['Paragraph', 'Font', 'Slides']), null);
});

test('ribbon fit: the left-most collapsed group that fits opens again, blocked ones stay closed', () => {
  const c = [{ name: 'Font', extra: 260 }, { name: 'Paragraph', extra: 180 }, { name: 'Drawing', extra: 90 }];
  assert.equal(nextToExpand(c, 300), 'Font');
  assert.equal(nextToExpand(c, 200), 'Paragraph');
  assert.equal(nextToExpand(c, 200, ['Paragraph']), 'Drawing');
  assert.equal(nextToExpand(c, 50), null);
});

test('fit level: steps undo when the width the body needed is available again', () => {
  assert.equal(levelForWidth([1200, 1000], 2, 900), 2);
  assert.equal(levelForWidth([1200, 1000], 2, 1000), 1);
  assert.equal(levelForWidth([1200, 1000], 2, 1300), 0);
});

test('first icon of a group (the collapsed button shows it)', () => {
  const A = () => null; const B = () => null;
  const tree = [{ props: { children: [null, { props: { icon: A } }] } }, { props: { icon: B } }];
  assert.equal(firstIcon(tree), A);
  assert.equal(firstIcon('text'), null);
});

test('collab: a room holding the saved file with other ids is the same content', () => {
  const a = { size: { w: 1, h: 1 }, slides: [{ id: 's1', elements: [{ id: 'e1', type: 'shape', text: 'x' }], animations: [{ el: 'e1', effect: 'fade' }] }] };
  const b = { size: { w: 1, h: 1 }, slides: [{ id: 'zz', elements: [{ id: 'q9', type: 'shape', text: 'x' }], animations: [{ el: 'q9', effect: 'fade' }] }] };
  assert.equal(sameContent(a, b), true);
  const c = { ...b, slides: [{ ...b.slides[0], elements: [{ id: 'q9', type: 'shape', text: 'changed' }] }] };
  assert.equal(sameContent(a, c), false);
});
