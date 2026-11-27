// Vendors a compact slice of caniemail.com's public support data for the preview's
// compatibility panel: only the features src/preview/features.ts can detect, and only
// the client platforms the preview targets (Gmail desktop/iOS/Android, Outlook for
// Windows, Apple Mail macOS/iOS).
//
//   npx tsx scripts/build-caniemail.ts
//
// Output: src/preview/data/caniemail.json (kept well under 60 KB).
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { FEATURE_SLUGS, CANIEMAIL_PLATFORMS } from '../src/preview/caniemail-spec.js';

const SOURCE = 'https://www.caniemail.com/api/data.json';
const OUT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../src/preview/data/caniemail.json');

interface RawFeature {
  slug: string; title: string; url: string; category: string;
  stats: Record<string, Record<string, Record<string, string>>>;
  notes_by_num: Record<string, string> | null;
}
interface RawData { last_update_date: string; data: RawFeature[] }

/** Markdown links and code ticks -> plain text; one line; capped length. */
function plain(note: string): string {
  return note
    .replace(/\[([^\]]+)\]\([^)]+\)/g, '$1')
    .replace(/`/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 220);
}

async function main() {
  const res = await fetch(SOURCE, { headers: { 'User-Agent': 'kiit-mailer-build-script' } });
  if (!res.ok) throw new Error(`caniemail fetch failed: HTTP ${res.status}`);
  const raw = (await res.json()) as RawData;
  const bySlug = new Map(raw.data.map(f => [f.slug, f]));

  const features: Record<string, { t: string; u: string; c: string; s: Record<string, string>; n?: Record<string, string> }> = {};
  const missing: string[] = [];
  for (const slug of FEATURE_SLUGS) {
    const f = bySlug.get(slug);
    if (!f) { missing.push(slug); continue; }
    const s: Record<string, string> = {};
    const used = new Set<string>();
    for (const [key, [family, platform]] of Object.entries(CANIEMAIL_PLATFORMS)) {
      const versions = f.stats?.[family]?.[platform];
      if (!versions) { s[key] = 'u'; continue; }
      const keys = Object.keys(versions);
      // caniemail lists versions chronologically; the last entry is the current state.
      const v = String(versions[keys[keys.length - 1]] ?? 'u').trim();
      s[key] = v;
      for (const m of v.matchAll(/#(\d+)/g)) used.add(m[1]);
    }
    const n: Record<string, string> = {};
    for (const num of used) {
      const text = f.notes_by_num?.[num];
      if (text) n[num] = plain(text);
    }
    features[slug] = { t: f.title, u: f.url, c: f.category, s, ...(Object.keys(n).length ? { n } : {}) };
  }

  const out = {
    source: SOURCE,
    license: 'caniemail.com data, CC BY 4.0',
    updated: raw.last_update_date,
    features,
  };
  await mkdir(path.dirname(OUT), { recursive: true });
  const json = JSON.stringify(out);
  await writeFile(OUT, json + '\n', 'utf8');
  console.log(`Wrote ${Object.keys(features).length} features (${(json.length / 1024).toFixed(1)} KB) to ${OUT}`);
  if (missing.length) console.warn(`Not found in caniemail data: ${missing.join(', ')}`);
}

main().catch(err => { console.error(err); process.exit(1); });
