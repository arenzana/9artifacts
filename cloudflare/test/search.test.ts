import { env } from 'cloudflare:workers';
import { beforeAll, describe, expect, it } from 'vitest';
import { Client, body, fresh, makeToken } from './helpers';

let c: Client;

beforeAll(async () => {
  await fresh();
  c = new Client(undefined, await makeToken());
});

// The same cases are in TestSearchInsideContent (internal/app/server_test.go).
describe('search inside artifact text', () => {
  it('matches page text, not markup, and follows the latest version', async () => {
    const a = await body(await c.upload('/api/artifacts', { title: 'Weekly' }, {
      'index.html': '<style>.zebra{}</style><h1>Quarterly <b>Revenue</b></h1><script>var walrus=1</script>',
    }));
    const found = async (q: string) => (await body(await c.api('GET', '/api/artifacts?q=' + q))).artifacts.some((x: any) => x.id === a.id);
    expect(await found('revenue')).toBe(true);
    expect(await found('weekly')).toBe(true);
    for (const q of ['zebra', 'walrus', 'h1']) expect(await found(q), q).toBe(false);

    await c.upload(`/api/artifacts/${a.id}/versions`, {}, { 'index.html': '<p>Churn</p>' });
    expect(await found('revenue')).toBe(false);
    expect(await found('churn')).toBe(true);

    const rpc = await body(await c.api('POST', '/mcp', { jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name: 'list_artifacts', arguments: { query: 'Churn' } } }));
    expect(rpc.result.content[0].text).toContain(a.id);

    expect((await c.api('DELETE', '/api/artifacts/' + a.id)).status).toBe(200);
    expect(await env.DB.prepare('SELECT COUNT(*) AS n FROM search_text').first('n')).toBe(0);
  });
});
