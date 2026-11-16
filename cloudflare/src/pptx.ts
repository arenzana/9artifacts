// PowerPoint export. The deck is built from the text form of the entry file
// (text.ts): one slide per .slide of a slide deck, or per h1/h2 of any other
// page, with a title and editable text. The page's own design is not carried
// over. The same code is in internal/app/pptx.go.
import type { Ctx } from './env';
import { getArtifact, getVersion } from './store';
import { artifactText } from './text';
import { json } from './util';

interface Para {
  text: string;
  bullet?: boolean;
  bold?: boolean;
}
interface Slide {
  title: string;
  paras: Para[];
}

const SEARCHABLE = new Set(['.html', '.htm', '.md', '.txt', '.csv', '.svg']);

/** Cuts the text form of a page into slides. */
export function pptSlides(text: string): Slide[] {
  const chunks = /^\[Slide \d+\]$/m.test(text) ? text.split(/^\[Slide \d+\]$/m) : text.split(/^(?=#{1,2} )/m);
  const out: Slide[] = [];
  for (const c of chunks) {
    const s: Slide = { title: '', paras: [] };
    for (const raw of c.split('\n')) {
      const line = raw.trim();
      const h = /^#{1,6} +/.exec(line)?.[0] ?? '';
      if (line === '' || line === '---' || line.startsWith('```')) continue;
      if (h && !s.title) s.title = line.slice(h.length);
      else if (h) s.paras.push({ text: line.slice(h.length), bold: true });
      else if (line.startsWith('- ')) s.paras.push({ text: line.slice(2), bullet: true });
      else s.paras.push({ text: line });
    }
    if (!s.title && s.paras.length) s.title = s.paras.shift()!.text;
    if (s.title) out.push(s);
  }
  return out;
}

const esc = (s: string) =>
  s
    .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const TYPE = 'application/vnd.openxmlformats-officedocument.';
const TREE = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

const THEME =
  HEAD +
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="9artifacts"><a:themeElements>' +
  '<a:clrScheme name="9artifacts"><a:dk1><a:srgbClr val="1A1A1A"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44403C"/></a:dk2><a:lt2><a:srgbClr val="F5F5F4"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="C2410C"/></a:accent1><a:accent2><a:srgbClr val="0F766E"/></a:accent2><a:accent3><a:srgbClr val="1D4ED8"/></a:accent3><a:accent4><a:srgbClr val="A16207"/></a:accent4><a:accent5><a:srgbClr val="7E22CE"/></a:accent5><a:accent6><a:srgbClr val="BE123C"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="1D4ED8"/></a:hlink><a:folHlink><a:srgbClr val="7E22CE"/></a:folHlink></a:clrScheme>' +
  '<a:fontScheme name="9artifacts"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="9artifacts"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>' +
  '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>' +
  '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
  '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>' +
  '</a:themeElements></a:theme>';

// Title and body boxes on a 16:9 slide (12192000 × 6858000 EMU).
const TITLE_BOX = '<a:xfrm><a:off x="685800" y="457200"/><a:ext cx="10820400" cy="1143000"/></a:xfrm>';
const BODY_BOX = '<a:xfrm><a:off x="685800" y="1752600"/><a:ext cx="10820400" cy="4572000"/></a:xfrm>';

const MASTER =
  HEAD +
  '<p:sldMaster ' +
  NS +
  '><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>' +
  TREE +
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr>' +
  TITLE_BOX +
  '</p:spPr><p:txBody><a:bodyPr anchor="b"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr>' +
  BODY_BOX +
  '</p:spPr><p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>' +
  '</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
  '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles>' +
  '<p:titleStyle><a:lvl1pPr algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buNone/><a:defRPr sz="3600" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>' +
  '<p:bodyStyle><a:lvl1pPr marL="0" indent="0" algn="l"><a:spcBef><a:spcPts val="800"/></a:spcBef><a:buNone/><a:defRPr sz="2000"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:bodyStyle>' +
  '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>';

const LAYOUT =
  HEAD +
  '<p:sldLayout ' +
  NS +
  ' type="obj" preserve="1"><p:cSld name="Title and Content"><p:spTree>' +
  TREE +
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>' +
  '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';

function rels(...targets: string[]): string {
  let b = HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
  for (let i = 0; i + 1 < targets.length; i += 2) b += `<Relationship Id="rId${i / 2 + 1}" Type="${REL}${targets[i]}" Target="${targets[i + 1]}"/>`;
  return b + '</Relationships>';
}

function slideXML(s: Slide): string {
  // Long slides get smaller text so they still fit the body box.
  const size = s.paras.length > 16 ? 1100 : s.paras.length > 9 ? 1400 : 2000;
  let body = '';
  for (const p of s.paras) {
    body += '<a:p>';
    if (p.bullet) body += '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>';
    body += `<a:r><a:rPr lang="en-US" sz="${size}"${p.bold ? ' b="1"' : ''}/><a:t>${esc(p.text)}</a:t></a:r></a:p>`;
  }
  if (!s.paras.length) body = '<a:p><a:endParaRPr lang="en-US"/></a:p>';
  return (
    HEAD +
    '<p:sld ' + NS + '><p:cSld><p:spTree>' + TREE +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>' + esc(s.title) + '</a:t></a:r></a:p></p:txBody></p:sp>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>' + body + '</p:txBody></p:sp>' +
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
  );
}

/** The files of a .pptx package in the order they are written. */
export function pptParts(slides: Slide[]): [string, string][] {
  let types =
    HEAD +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    `<Override PartName="/ppt/presentation.xml" ContentType="${TYPE}presentationml.presentation.main+xml"/>` +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${TYPE}presentationml.slideMaster+xml"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${TYPE}presentationml.slideLayout+xml"/>` +
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="${TYPE}theme+xml"/>`;
  let ids = '';
  const presRels = ['slideMaster', 'slideMasters/slideMaster1.xml', 'theme', 'theme/theme1.xml'];
  const parts: [string, string][] = [];
  slides.forEach((s, i) => {
    const name = `slides/slide${i + 1}.xml`;
    types += `<Override PartName="/ppt/${name}" ContentType="${TYPE}presentationml.slide+xml"/>`;
    ids += `<p:sldId id="${256 + i}" r:id="rId${i + 3}"/>`;
    presRels.push('slide', name);
    parts.push(['ppt/' + name, slideXML(s)], [`ppt/slides/_rels/slide${i + 1}.xml.rels`, rels('slideLayout', '../slideLayouts/slideLayout1.xml')]);
  });
  const pres =
    HEAD +
    '<p:presentation ' + NS + '><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>' + ids +
    '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>';
  return [
    ['[Content_Types].xml', types + '</Types>'],
    ['_rels/.rels', rels('officeDocument', 'ppt/presentation.xml')],
    ['ppt/presentation.xml', pres],
    ['ppt/_rels/presentation.xml.rels', rels(...presRels)],
    ['ppt/slideMasters/slideMaster1.xml', MASTER],
    ['ppt/slideMasters/_rels/slideMaster1.xml.rels', rels('slideLayout', '../slideLayouts/slideLayout1.xml', 'theme', '../theme/theme1.xml')],
    ['ppt/slideLayouts/slideLayout1.xml', LAYOUT],
    ['ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels('slideMaster', '../slideMasters/slideMaster1.xml')],
    ['ppt/theme/theme1.xml', THEME],
    ...parts,
  ];
}

const CRC = Uint32Array.from({ length: 256 }, (_, n) => {
  for (let k = 0; k < 8; k++) n = n & 1 ? 0xedb88320 ^ (n >>> 1) : n >>> 1;
  return n >>> 0;
});

/** A zip archive with the files stored uncompressed: a deck is a few kilobytes of XML. */
export function zipStored(files: [string, string][]): Uint8Array {
  const enc = new TextEncoder();
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  // The part local and central headers share: version, UTF-8 names, stored, time, date 1980-01-01, CRC, sizes, name length.
  const common = (d: DataView, p: number, name: Uint8Array, crc: number, size: number) => {
    for (const x of [20, 0x0800, 0, 0, 0x21]) (d.setUint16(p, x, true), (p += 2));
    for (const x of [crc, size, size]) (d.setUint32(p, x, true), (p += 4));
    d.setUint16(p, name.length, true);
  };
  for (const [path, text] of files) {
    const name = enc.encode(path);
    const data = enc.encode(text);
    let crc = ~0;
    for (const x of data) crc = CRC[(crc ^ x) & 0xff] ^ (crc >>> 8);
    crc = ~crc >>> 0;
    const local = new Uint8Array(30 + name.length);
    let d = new DataView(local.buffer);
    d.setUint32(0, 0x04034b50, true);
    common(d, 4, name, crc, data.length);
    local.set(name, 30);
    const entry = new Uint8Array(46 + name.length);
    d = new DataView(entry.buffer);
    d.setUint32(0, 0x02014b50, true);
    d.setUint16(4, 20, true);
    common(d, 6, name, crc, data.length);
    d.setUint32(42, offset, true);
    entry.set(name, 46);
    central.push(entry);
    chunks.push(local, data);
    offset += local.length + data.length;
  }
  const size = central.reduce((n, b) => n + b.length, 0);
  const end = new Uint8Array(22);
  const d = new DataView(end.buffer);
  d.setUint32(0, 0x06054b50, true);
  d.setUint16(8, files.length, true);
  d.setUint16(10, files.length, true);
  d.setUint32(12, size, true);
  d.setUint32(16, offset, true);
  const out = new Uint8Array(offset + size + 22);
  let p = 0;
  for (const b of [...chunks, ...central, end]) (out.set(b, p), (p += b.length));
  return out;
}

/** GET /api/artifacts/{id}/pptx: the entry file as an editable PowerPoint deck. */
export async function handlePPTX(c: Ctx, id: string): Promise<Response> {
  const n = Number.parseInt(c.url.searchParams.get('v') ?? '', 10) || 0;
  const [art, v] = await Promise.all([getArtifact(c, id), getVersion(c, id, n)]);
  if (!art || !v) return json(404, { error: 'artifact or version not found' });
  const dot = v.entry.lastIndexOf('.');
  if (dot < 0 || !SEARCHABLE.has(v.entry.slice(dot).toLowerCase())) return json(415, { error: v.entry + ' is not text, so it has no slides to export' });
  const text = await artifactText(c, id, art.title, v, []);
  let slides = pptSlides(text.slice(text.indexOf('\n---\n\n') + 6));
  if (!slides.length) slides = [{ title: art.title, paras: [] }];
  const name = art.title.replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^[-.]+|[-.]+$/g, '') || id;
  return new Response(zipStored(pptParts(slides)), {
    headers: {
      'Content-Type': TYPE + 'presentationml.presentation',
      'Content-Disposition': `attachment; filename="${name}.pptx"; filename*=UTF-8''${encodeURIComponent(art.title + '.pptx')}`,
      'Cache-Control': 'no-store',
    },
  });
}
