import { beforeAll, describe, expect, it } from 'vitest';
import { Client, body, fresh, makeToken } from './helpers';

let c: Client;

beforeAll(async () => {
  await fresh();
  c = new Client(undefined, await makeToken());
});

// The same cases are in TestPPTXExport (internal/app/server_test.go). The archive stores its
// files uncompressed, so the XML can be read straight from the bytes.
describe('PowerPoint export', () => {
  const open = async (files: Record<string, string>) => {
    const a = await body(await c.upload('/api/artifacts', { title: 'Quý 3 & more' }, files));
    const r = await c.api('GET', `/api/artifacts/${a.id}/pptx`);
    return { r, text: new TextDecoder().decode(await r.arrayBuffer()) };
  };

  it('makes one slide per .slide with escaped, editable text', async () => {
    const { r, text } = await open({
      'index.html': '<section class="slide"><h1>Plan &lt;A&gt;</h1><ul><li>one<li>two</ul></section><section class="slide"><h2>Next</h2><p>Ship it</p></section>',
    });
    expect(r.status).toBe(200);
    expect(r.headers.get('Content-Type')).toBe('application/vnd.openxmlformats-officedocument.presentationml.presentation');
    expect(r.headers.get('Content-Disposition')).toContain('filename="Qu-3-more.pptx"');
    expect(text.startsWith('PK\x03\x04')).toBe(true);
    for (const want of ['<a:t>Plan &lt;A&gt;</a:t>', '<a:buChar', '<a:t>two</a:t>', '<a:t>Ship it</a:t>', 'PartName="/ppt/slides/slide2.xml"', 'Id="rId4"']) expect(text, want).toContain(want);
    expect(text).not.toContain('slide3.xml');
  });

  it('cuts any other page at its h1/h2 and refuses a binary entry', async () => {
    expect((await open({ 'notes.md': '# Intro\n\nHello\n\n## Costs\n\n- rent\n' })).text).toContain('ppt/slides/slide2.xml');
    expect((await open({ 'photo.png': 'x' })).r.status).toBe(415);
  });
});
