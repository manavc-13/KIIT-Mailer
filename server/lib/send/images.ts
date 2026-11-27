// Hosted-image handling for the send path. When there is no public Blob store,
// in-body images are stored privately and referenced as /api/uploads/<id>/content.
// Recipients cannot load that URL, so at send time every such <img src> becomes a
// CID inline attachment (cid:<id>@kiit-mailer). Preview keeps the URL. Pure.
import type { MailAttachment } from '../mailer.js';
import { appBaseUrl } from '../unsubscribe.js';

const UUID = '[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}';

function baseUrlPattern(): string {
  const base = appBaseUrl();
  const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return base ? `(?:${esc(base)})?` : '';
}

/** Regex matching src="(appBaseUrl())?/api/uploads/<uuid>/content" (group 1 = quote, 2 = id). */
function srcRegex(): RegExp {
  return new RegExp(`(\\ssrc\\s*=\\s*)(["']?)${baseUrlPattern()}/api/uploads/(${UUID})/content\\2`, 'gi');
}

/** Upload ids referenced by <img src> as app-hosted images. */
export function hostedImageIds(html: string): string[] {
  const ids = new Set<string>();
  for (const m of html.matchAll(srcRegex())) ids.add(m[3].toLowerCase());
  return [...ids];
}

export const cidFor = (uploadId: string) => `${uploadId.toLowerCase()}@kiit-mailer`;

export interface HostedImage { id: string; filename: string; contentType: string; content: Buffer }

/**
 * Rewrite app-hosted image URLs to cid: references and return the inline attachments
 * to add. Images not found in `images` are left unchanged.
 */
export function inlineHostedImages(html: string, images: HostedImage[]): { html: string; attachments: MailAttachment[] } {
  const byId = new Map(images.map(i => [i.id.toLowerCase(), i]));
  const used = new Set<string>();
  const out = html.replace(srcRegex(), (m, pre: string, q: string, id: string) => {
    const img = byId.get(id.toLowerCase());
    if (!img) return m;
    used.add(img.id.toLowerCase());
    const quote = q || '"';
    return `${pre}${quote}cid:${cidFor(img.id)}${quote}`;
  });
  const attachments: MailAttachment[] = [...used].map(id => {
    const img = byId.get(id)!;
    return { filename: img.filename, content: img.content, contentType: img.contentType, cid: cidFor(img.id), contentDisposition: 'inline' };
  });
  return { html: out, attachments };
}
