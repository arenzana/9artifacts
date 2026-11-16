// Search also looks inside the entry file of the latest version. The text is
// read from R2 the first time a search needs it and kept in D1 (search_text),
// so later searches are one query.
import type { Ctx } from './env';
import { fileKey, type Version } from './store';

const SEARCH_READ = 512 << 10; // bytes read from the entry file
const SEARCH_KEEP = 64 << 10; // characters of text kept per artifact
const SEARCH_FILL = 20; // files read per request, to stay inside the subrequest limit
const SEARCHABLE = new Set(['.html', '.htm', '.md', '.txt', '.csv', '.svg']);
const HIDDEN = /<(script|style)\b[\s\S]*?<\/(script|style)>/gi;
const TAG = /<[^>]*>/g;

const ext = (name: string) => (name.lastIndexOf('.') < 0 ? '' : name.slice(name.lastIndexOf('.')).toLowerCase());

async function readSearchText(c: Ctx, id: string, v: Version): Promise<string> {
  const f = v.files.find((x) => x.name === v.entry);
  const obj = f && (await c.files.get(fileKey(id, v.n, f), { range: { offset: 0, length: SEARCH_READ } }));
  if (!obj) return '';
  let text = new TextDecoder().decode(await obj.arrayBuffer());
  if (['.html', '.htm', '.svg'].includes(ext(v.entry))) text = text.replace(HIDDEN, ' ').replace(TAG, ' ');
  return text.split(/\s+/).filter(Boolean).join(' ').toLowerCase().slice(0, SEARCH_KEEP);
}

/**
 * Ids of the artifacts whose latest entry file contains q, which must
 * already be lowercase. A failure (the table is missing until the migration
 * runs) only means no content matches.
 */
export async function contentMatches(c: Ctx, items: { a: { id: string }; latest: Version }[], q: string): Promise<Set<string>> {
  const hits = new Set<string>();
  try {
    const { results } = await c.db.prepare('SELECT artifact_id, n, instr(text, ?1) > 0 AS hit FROM search_text').bind(q).all<{ artifact_id: string; n: number; hit: number }>();
    const known = new Map(results.map((r) => [r.artifact_id, r]));
    const fill: Promise<void>[] = [];
    for (const { a, latest } of items) {
      if (!SEARCHABLE.has(ext(latest.entry))) continue;
      const row = known.get(a.id);
      if (row?.n === latest.n) {
        if (row.hit) hits.add(a.id);
      } else if (fill.length < SEARCH_FILL) {
        fill.push(
          readSearchText(c, a.id, latest).then(async (text) => {
            await c.db.prepare('INSERT OR REPLACE INTO search_text (artifact_id, n, text) VALUES (?, ?, ?)').bind(a.id, latest.n, text).run();
            if (text.includes(q)) hits.add(a.id);
          }),
        );
      }
    }
    await Promise.all(fill);
  } catch (e) {
    console.error('content search:', e);
  }
  return hits;
}
