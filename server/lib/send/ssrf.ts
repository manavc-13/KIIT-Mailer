// SSRF-safe image fetch for rehosting pasted images (Google Docs image URLs expire).
// - https only; at most 3 redirects, each re-validated
// - every DNS answer is checked at CONNECT time (custom lookup), so DNS rebinding
//   between "check" and "fetch" cannot reach private, loopback or link-local hosts
// - image/* content types only, size and time limits
import https from 'node:https';
import dns from 'node:dns';
import net from 'node:net';
import type { LookupFunction } from 'node:net';

export class FetchBlockedError extends Error {}

function v4ToInt(ip: string): number {
  return ip.split('.').reduce((a, o) => (a << 8) + (Number(o) & 255), 0) >>> 0;
}
function inV4(ip: string, cidr: string): boolean {
  const [base, bits] = cidr.split('/');
  const mask = Number(bits) === 0 ? 0 : (~0 << (32 - Number(bits))) >>> 0;
  return (v4ToInt(ip) & mask) === (v4ToInt(base) & mask);
}
const V4_BLOCKED = [
  '0.0.0.0/8', '10.0.0.0/8', '100.64.0.0/10', '127.0.0.0/8', '169.254.0.0/16', '172.16.0.0/12',
  '192.0.0.0/24', '192.0.2.0/24', '192.88.99.0/24', '192.168.0.0/16', '198.18.0.0/15', '198.51.100.0/24',
  '203.0.113.0/24', '224.0.0.0/4', '240.0.0.0/4',
];

function expandV6(ip: string): number[] | null {
  let s = ip.toLowerCase().replace(/^\[|\]$/g, '').split('%')[0];
  let tail: number[] = [];
  const v4 = /(\d+\.\d+\.\d+\.\d+)$/.exec(s);
  if (v4) {
    const n = v4ToInt(v4[1]);
    tail = [(n >>> 16) & 0xffff, n & 0xffff];
    s = s.slice(0, -v4[1].length).replace(/([^:]):$/, '$1');
  }
  const parts = s.split('::');
  if (parts.length > 2) return null;
  const head = parts[0] ? parts[0].split(':').filter(Boolean).map(h => parseInt(h, 16)) : [];
  const rest = parts.length === 2 && parts[1] ? parts[1].split(':').filter(Boolean).map(h => parseInt(h, 16)) : [];
  const total = head.length + rest.length + tail.length;
  if (parts.length === 1 && total !== 8) return null;
  const fill = parts.length === 2 ? new Array(Math.max(0, 8 - total)).fill(0) : [];
  const out = [...head, ...fill, ...rest, ...tail];
  return out.length === 8 && out.every(n => Number.isFinite(n)) ? out : null;
}

/** True for any address that must never be fetched from the server. */
export function isBlockedIp(ip: string): boolean {
  const kind = net.isIP(ip.replace(/^\[|\]$/g, '').split('%')[0]);
  if (kind === 4) return V4_BLOCKED.some(c => inV4(ip, c)) || ip === '255.255.255.255';
  if (kind !== 6) return true;
  const w = expandV6(ip);
  if (!w) return true;
  if (w.every(x => x === 0)) return true; // ::
  if (w.slice(0, 7).every(x => x === 0) && w[7] === 1) return true; // ::1
  // IPv4-mapped / -compatible / NAT64: judge the embedded IPv4 address
  const embedded = `${w[6] >>> 8}.${w[6] & 255}.${w[7] >>> 8}.${w[7] & 255}`;
  if (w.slice(0, 5).every(x => x === 0) && (w[5] === 0xffff || w[5] === 0)) return isBlockedIp(embedded);
  if (w[0] === 0x64 && w[1] === 0xff9b) return isBlockedIp(embedded);
  if ((w[0] & 0xfe00) === 0xfc00) return true; // fc00::/7 unique local
  if ((w[0] & 0xffc0) === 0xfe80) return true; // fe80::/10 link-local
  if ((w[0] & 0xffc0) === 0xfec0) return true; // fec0::/10 site-local (deprecated)
  if ((w[0] & 0xff00) === 0xff00) return true; // multicast
  if (w[0] === 0x2001 && w[1] === 0x0db8) return true; // documentation
  if (w[0] === 0x2002) return isBlockedIp(`${w[1] >>> 8}.${w[1] & 255}.${w[2] >>> 8}.${w[2] & 255}`); // 6to4
  return false;
}

/** Validates a URL for fetching: https, no credentials, no IP-literal private hosts. */
export function assertFetchableUrl(raw: string): URL {
  let u: URL;
  try { u = new URL(raw); } catch { throw new FetchBlockedError('That is not a valid URL.'); }
  if (u.protocol !== 'https:') throw new FetchBlockedError('Only https:// image links can be imported.');
  if (u.username || u.password) throw new FetchBlockedError('Image links with credentials are not allowed.');
  if (u.port && u.port !== '443') throw new FetchBlockedError('Image links must use the standard https port.');
  const host = u.hostname.replace(/^\[|\]$/g, '');
  if (!host || host === 'localhost' || host.endsWith('.localhost') || host.endsWith('.local') || host.endsWith('.internal')) {
    throw new FetchBlockedError('That address is not allowed.');
  }
  if (net.isIP(host) && isBlockedIp(host)) throw new FetchBlockedError('That address is not allowed.');
  return u;
}

/** DNS lookup that refuses private answers. Used by the socket itself (connect-time check). */
const guardedLookup: LookupFunction = (hostname, options, callback) => {
  dns.lookup(hostname, { all: true, family: (options as { family?: number }).family ?? 0 }, (err, addrs) => {
    if (err) return (callback as (e: Error | null, a: string, f: number) => void)(err, '', 0);
    const list = addrs as dns.LookupAddress[];
    const bad = list.find(a => isBlockedIp(a.address));
    if (!list.length || bad) {
      return (callback as (e: Error | null, a: string, f: number) => void)(new FetchBlockedError('That address is not allowed.'), '', 0);
    }
    if ((options as { all?: boolean }).all) return (callback as unknown as (e: null, a: dns.LookupAddress[]) => void)(null, list);
    return (callback as (e: null, a: string, f: number) => void)(null, list[0].address, list[0].family);
  });
};

export interface FetchedImage { buffer: Buffer; contentType: string; finalUrl: string; filename: string }

const ALLOWED_IMAGE_TYPES = ['image/png', 'image/jpeg', 'image/jpg', 'image/gif', 'image/webp'];
export const isAllowedImageType = (ct: string) => ALLOWED_IMAGE_TYPES.includes(ct.split(';')[0].trim().toLowerCase());

function requestOnce(u: URL, maxBytes: number, signal: AbortSignal): Promise<{ status: number; location?: string; contentType: string; buffer?: Buffer }> {
  return new Promise((resolve, reject) => {
    const req = https.get(u, {
      lookup: guardedLookup, signal,
      headers: { 'User-Agent': 'KIIT-Mailer-ImageImport/2.0', Accept: 'image/*' },
    }, res => {
      const status = res.statusCode || 0;
      const contentType = String(res.headers['content-type'] || '').toLowerCase();
      if (status >= 300 && status < 400 && res.headers.location) {
        res.resume();
        return resolve({ status, location: res.headers.location, contentType });
      }
      if (status !== 200) { res.resume(); return reject(new FetchBlockedError(`The image server answered HTTP ${status}.`)); }
      if (!isAllowedImageType(contentType)) {
        res.resume();
        return reject(new FetchBlockedError('That link is not a PNG, JPEG, GIF or WebP image.'));
      }
      const declared = Number(res.headers['content-length'] || 0);
      if (declared > maxBytes) { res.destroy(); return reject(new FetchBlockedError('The image is larger than 10 MB.')); }
      const chunks: Buffer[] = []; let total = 0;
      res.on('data', (c: Buffer) => {
        total += c.length;
        if (total > maxBytes) { res.destroy(); reject(new FetchBlockedError('The image is larger than 10 MB.')); return; }
        chunks.push(c);
      });
      res.on('end', () => resolve({ status, contentType, buffer: Buffer.concat(chunks) }));
      res.on('error', reject);
    });
    req.on('error', reject);
  });
}

/** Fetch an image from a public https URL with SSRF protection. */
export async function fetchImageSafely(raw: string, opts: { maxBytes?: number; timeoutMs?: number; maxRedirects?: number } = {}): Promise<FetchedImage> {
  const maxBytes = opts.maxBytes ?? 10 * 1024 * 1024;
  const maxRedirects = opts.maxRedirects ?? 3;
  const signal = AbortSignal.timeout(opts.timeoutMs ?? 10_000);
  let u = assertFetchableUrl(raw);
  for (let hop = 0; hop <= maxRedirects; hop++) {
    let r;
    try { r = await requestOnce(u, maxBytes, signal); } catch (err) {
      if (err instanceof FetchBlockedError) throw err;
      if (signal.aborted) throw new FetchBlockedError('The image took too long to download (over 10 seconds).');
      throw new FetchBlockedError('The image could not be downloaded.');
    }
    if (r.location) {
      if (hop === maxRedirects) break;
      u = assertFetchableUrl(new URL(r.location, u).toString());
      continue;
    }
    const ct = r.contentType.split(';')[0].trim();
    const ext = ct === 'image/png' ? 'png' : ct === 'image/gif' ? 'gif' : ct === 'image/webp' ? 'webp' : 'jpg';
    const base = decodeURIComponent(u.pathname.split('/').pop() || '').replace(/[^\w.\-]+/g, '_').slice(0, 60);
    const filename = base && /\.(png|jpe?g|gif|webp)$/i.test(base) ? base : `image.${ext}`;
    return { buffer: r.buffer!, contentType: ct === 'image/jpg' ? 'image/jpeg' : ct, finalUrl: u.toString(), filename };
  }
  throw new FetchBlockedError('The image link redirected too many times.');
}
