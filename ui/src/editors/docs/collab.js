/*
 * Co-editing (host.collab, y-prosemirror): the document is one Y.XmlFragment ("prosemirror"), comments / page setup /
 * headers live in the Y.Map "office". Cursors with names and colours come from the awareness (host.collab sets the
 * user). Conflict-free by Yjs; undo is per person (y-undo). Only the LEADER (the co-editor with the smallest client id)
 * reports unsaved changes to the shell, so exactly one editor saves the file – when it leaves, the next one takes over.
 * Solo editing never loads any of this: co-editing starts from Review → "Edit together", or by itself when someone
 * else is already editing the file.
 */
import * as Y from 'yjs';
import { Extension, getSchema } from '@tiptap/core';
import { ySyncPlugin, yCursorPlugin, yUndoPlugin, undoCommand, redoCommand, prosemirrorJSONToYXmlFragment, ySyncPluginKey } from 'y-prosemirror';
import { t } from '@hostwerk/i18n';

export { ySyncPluginKey };

/** connect, wait for the first sync, seed the shared document when we are the first → session */
export async function joinCollab(host, { extensions, initialDoc, model, timeout = 7000 }) {
  if (!host || !host.collab || typeof host.collab.connect !== 'function') throw new Error(t('Co-editing is not available in this Hostwerk version.'));
  const doc = new Y.Doc();
  const provider = await host.collab.connect({ doc });
  const synced = await Promise.race([provider.synced.then(() => true), new Promise((r) => setTimeout(() => r(false), timeout))]);
  if (!synced) { provider.destroy(); throw new Error(t('The co-editing server did not answer.')); }
  const awareness = provider.awareness;
  const fragment = doc.getXmlFragment('prosemirror');
  const meta = doc.getMap('office');
  let first = false;
  if (fragment.length === 0) {
    const others = [...awareness.getStates().keys()].filter((id) => id !== doc.clientID);
    // the server says whether we opened an empty room (hello.fresh) – only then we seed; else the content is on its way
    const fresh = !provider.hello || provider.hello.fresh !== false;
    if (others.length || !fresh) await new Promise((res) => { const tm = setTimeout(res, fresh ? 2500 : 6000); fragment.observeDeep(() => { clearTimeout(tm); res(); }); });
    if (fragment.length === 0) {
      first = true;
      doc.transact(() => {
        prosemirrorJSONToYXmlFragment(getSchema(extensions), initialDoc, fragment);
        for (const k of ['comments', 'page', 'header', 'footer', 'headerFirst', 'footerFirst', 'headerEven', 'footerEven', 'titlePage', 'evenOdd']) if (model[k] !== undefined) meta.set(k, JSON.parse(JSON.stringify(model[k])));
      });
    }
  }
  awareness.setLocalStateField('editor', true);
  // changes since the room's last save (local or remote) – a co-editor that becomes the saver saves them
  let unsaved = false;
  doc.on('update', () => { unsaved = true; });
  const session = {
    doc, provider, awareness, fragment, meta, first,
    /**
     * the saver: the server's leader (hello / leader control messages – the same for everybody at once, also before the
     * cursors arrive); an older server without them: the co-editor with the smallest client id
     */
    isLeader: () => {
      if (provider.hello && typeof provider.leader === 'boolean') return provider.leader;
      const ids = [...awareness.getStates().entries()].filter(([, s]) => s && s.editor).map(([id]) => id); return !ids.length || Math.min(...ids) === doc.clientID;
    },
    unsaved: () => unsaved,
    /** any save of the file in the room → fn({ mtime, by }); the changes are saved now */
    onSaved: (fn) => (typeof provider.on === 'function' ? provider.on('saved', (m) => { unsaved = false; fn(m || {}); }) : () => {}),
    /** this editor became / stopped being the saver → fn(isLeader) */
    onLeader: (fn) => (typeof provider.on === 'function' ? provider.on('leader', (v) => fn(!!v)) : () => {}),
    peers: () => [...awareness.getStates().entries()].filter(([id, s]) => id !== doc.clientID && s && s.user).map(([id, s]) => ({ id, name: s.user.name, color: s.user.color })),
    destroy: () => { try { provider.destroy(); } catch { /* gone */ } try { doc.destroy(); } catch { /* gone */ } },
  };
  return session;
}

/** the tiptap extension of a session (StarterKit's history must be off) */
export function collabExtension(session) {
  return Extension.create({
    name: 'ofCollab',
    priority: 1100,
    addCommands() {
      return {
        undo: () => ({ state, dispatch }) => undoCommand(state, dispatch),
        redo: () => ({ state, dispatch }) => redoCommand(state, dispatch),
      };
    },
    addKeyboardShortcuts() {
      return { 'Mod-z': () => this.editor.commands.undo(), 'Mod-y': () => this.editor.commands.redo(), 'Mod-Shift-z': () => this.editor.commands.redo() };
    },
    addProseMirrorPlugins() {
      return [
        ySyncPlugin(session.fragment),
        yCursorPlugin(session.awareness, {
          cursorBuilder: (user) => {
            const c = document.createElement('span'); c.className = 'of-docs-ycursor'; c.style.setProperty('--of-peer', user.color || '#2b6fd6');
            const l = document.createElement('span'); l.className = 'of-docs-ycursor-label'; l.textContent = user.name || '?';
            c.append(l); return c;
          },
          selectionBuilder: (user) => ({ style: `background-color: ${user.color || '#2b6fd6'}33`, class: 'of-docs-yselection' }),
        }),
        yUndoPlugin(),
      ];
    },
  });
}
