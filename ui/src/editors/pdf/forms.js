/*
 * AcroForm bridge between pdf.js (shows and edits the fields as HTML inputs, keeps the values in its
 * annotationStorage by widget id) and pdf-lib (writes them by field name – see edit.js applyFormValues).
 */

export const WIDGET = 20; // pdf.js AnnotationType.WIDGET

/** widget id → { name, fieldType: 'Tx'|'Btn'|'Ch'|'Sig', radio, checkbox, buttonValue, exportValue, multi } */
export async function collectWidgets(doc) {
  const map = new Map();
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    for (const a of await page.getAnnotations({ intent: 'display' })) {
      if (a.annotationType !== WIDGET || !a.fieldName) continue;
      map.set(a.id, { name: a.fieldName, fieldType: a.fieldType, radio: !!a.radioButton, checkbox: !!a.checkBox, buttonValue: a.buttonValue, exportValue: a.exportValue, multi: !!a.multiSelect, page: i - 1, readOnly: !!a.readOnly });
    }
  }
  return map;
}

/** pdf.js storage ({ [widgetId]: { value } }) → { fieldName: value } for pdf-lib */
export function valuesFromStorage(widgets, storage) {
  const out = {};
  for (const [id, v] of Object.entries(storage || {})) {
    const w = widgets.get(id);
    if (!w || !v || !('value' in v) || w.readOnly) continue;
    if (w.fieldType === 'Tx') out[w.name] = v.value == null ? '' : String(v.value);
    else if (w.fieldType === 'Btn') {
      if (w.radio) { if (v.value === true || (typeof v.value === 'string' && v.value !== 'Off')) out[w.name] = w.buttonValue; else if (!(w.name in out)) out[w.name] = ''; }
      else if (w.checkbox) out[w.name] = v.value === true || (typeof v.value === 'string' && v.value !== 'Off');
    } else if (w.fieldType === 'Ch') {
      const val = Array.isArray(v.value) ? v.value : v.value == null ? [] : [v.value];
      out[w.name] = w.multi ? val.map(String) : (val[0] == null ? '' : String(val[0]));
    }
  }
  return out;
}

/** short summary for the Forms tab */
export function summarize(widgets) {
  const names = new Map();
  for (const w of widgets.values()) if (!names.has(w.name)) names.set(w.name, w.fieldType);
  const c = { text: 0, check: 0, choice: 0, signature: 0 };
  for (const t of names.values()) { if (t === 'Tx') c.text++; else if (t === 'Btn') c.check++; else if (t === 'Ch') c.choice++; else if (t === 'Sig') c.signature++; }
  return { fields: names.size, ...c };
}
