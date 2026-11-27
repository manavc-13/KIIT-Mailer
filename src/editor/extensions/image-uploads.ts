// Re-hosts images flagged `rehost` (pasted Google Docs / external URLs, pasted or dropped
// files carried as data: URIs) and swaps in the hosted URL once the upload finishes.
import { Extension, type Editor } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import type { EditorView } from '@tiptap/pm/view';
import type { UploadRef } from '../../../shared/types';
import { dataUriToBlob } from '@/lib/uploads';

export interface ImageUploadsOptions {
  uploadFile: (file: Blob, filename: string) => Promise<UploadRef>;
  rehostImage: (url: string) => Promise<UploadRef>;
  onError?: (message: string) => void;
  onBusyChange?: (busy: boolean) => void;
}

export const imageUploadsKey = new PluginKey('kbImageUploads');

const EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp', 'image/svg+xml': 'svg', 'image/bmp': 'bmp' };

/** Read image files (paste/drop) into data: URIs; they are uploaded by the plugin. */
export function readImageFiles(files: File[]): Promise<string[]> {
  return Promise.all(files.map(f => new Promise<string>((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result));
    r.onerror = () => reject(r.error);
    r.readAsDataURL(f);
  })));
}

/** Insert image files at `pos` (or the selection) flagged for upload. */
export function insertImageFiles(editor: Editor, files: File[], pos?: number, onError?: (m: string) => void) {
  const ok = files.filter(f => f.type.startsWith('image/') && f.size <= 10 * 1024 * 1024);
  if (ok.length < files.length) onError?.('Some files were skipped: only images under 10 MB can be placed in the message.');
  if (!ok.length) return;
  readImageFiles(ok).then(srcs => {
    const content = srcs.map((src, i) => ({ type: 'image', attrs: { src, alt: ok[i].name.replace(/\.[a-z0-9]+$/i, '').replace(/[-_]+/g, ' '), rehost: true } }));
    const chain = editor.chain().focus();
    if (typeof pos === 'number') chain.insertContentAt(pos, content).run();
    else chain.insertContent(content).run();
  }).catch(() => onError?.('Could not read the image file.'));
}

export const ImageUploads = Extension.create<ImageUploadsOptions>({
  name: 'kbImageUploads',

  addOptions() {
    return {
      uploadFile: () => Promise.reject(new Error('Uploads are not configured')),
      rehostImage: () => Promise.reject(new Error('Uploads are not configured')),
      onError: undefined,
      onBusyChange: undefined,
    };
  },

  addProseMirrorPlugins() {
    const opts = this.options;
    const inflight = new Set<string>();
    let lastBusy = false;
    const setBusy = () => {
      const busy = inflight.size > 0;
      if (busy !== lastBusy) { lastBusy = busy; opts.onBusyChange?.(busy); }
    };

    const finish = (view: EditorView, src: string, ref: UploadRef | null, message?: string) => {
      inflight.delete(src);
      setBusy();
      if (view.isDestroyed) return;
      const tr = view.state.tr;
      let found = false;
      view.state.doc.descendants((node, pos) => {
        if (node.type.name === 'image' && node.attrs.rehost && node.attrs.src === src) {
          found = true;
          const p = tr.mapping.map(pos);
          if (ref?.publicUrl) tr.setNodeMarkup(p, undefined, { ...node.attrs, src: ref.publicUrl, rehost: false });
          else if (src.startsWith('data:')) tr.delete(p, p + node.nodeSize);
          else tr.setNodeMarkup(p, undefined, { ...node.attrs, rehost: false });
        }
      });
      if (found) {
        tr.setMeta('addToHistory', false);
        view.dispatch(tr);
      }
      if (message) opts.onError?.(message);
    };

    const scan = (view: EditorView) => {
      view.state.doc.descendants(node => {
        if (node.type.name !== 'image' || !node.attrs.rehost) return;
        const src: string = node.attrs.src || '';
        if (!src || inflight.has(src)) return;
        inflight.add(src);
        setBusy();
        if (src.startsWith('data:')) {
          let blob: Blob;
          try { blob = dataUriToBlob(src); } catch { finish(view, src, null, 'A pasted image could not be read and was removed.'); return; }
          const name = `pasted-image.${EXT[blob.type] || 'png'}`;
          opts.uploadFile(blob, name)
            .then(ref => finish(view, src, ref))
            .catch(err => finish(view, src, null, `A pasted image could not be uploaded and was removed (${(err as Error).message}).`));
        } else {
          opts.rehostImage(src)
            .then(ref => finish(view, src, ref))
            .catch(err => finish(view, src, null, `A pasted image could not be copied to our storage, so it still points at the original address, which may expire (${(err as Error).message}).`));
        }
      });
    };

    return [new Plugin({
      key: imageUploadsKey,
      view: view => {
        queueMicrotask(() => { if (!view.isDestroyed) scan(view); });
        return {
          update: (v, prev) => { if (prev.doc !== v.state.doc) scan(v); },
          destroy: () => { inflight.clear(); setBusy(); },
        };
      },
    })];
  },
});
