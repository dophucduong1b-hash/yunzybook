import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipEntries, unzipBytes } from '../js/zip.js';
import { parseDocxBytes, buildDocxBytesFromPages, parseDocumentXml } from '../js/docx.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, '..');

async function testZipRoundTripAndSecurity() {
  const payload = [
    ['hello.txt', 'Xin chào YunzyBook iPhone! '.repeat(20)],
    ['folder/sub.xml', '<root><item id="1">Tiếng Việt</item></root>']
  ];
  const packed = await zipEntries(payload);
  const unpacked = await unzipBytes(packed);
  assert.equal(unpacked.size, 2);
  const dec = new TextDecoder();
  assert.equal(dec.decode(unpacked.get('hello.txt')), payload[0][1]);
  assert.equal(dec.decode(unpacked.get('folder/sub.xml')), payload[1][1]);
  console.log('PASS: zip round-trip & UTF-8 filenames');
}

async function testDocxRoundTrip() {
  const pages = [
    {
      blocks: [
        { tag: 'h1', align: 'center', runs: [{ text: 'Tài liệu PWA', bold: true }] },
        {
          tag: 'p',
          align: '',
          runs: [
            { text: 'Bình thường ' },
            { text: 'đậm', bold: true },
            { text: ' sáng', highlight: true }
          ]
        },
        { tag: 'table', rows: [['Cột 1', 'Cột 2'], ['A', 'B']] }
      ]
    },
    {
      blocks: [
        { tag: 'h2', align: '', runs: [{ text: 'Trang thứ hai' }] },
        { tag: 'ul_li', align: '', runs: [{ text: 'Mục danh sách' }] }
      ]
    }
  ];

  const docxBytes = await buildDocxBytesFromPages(pages);
  const parsed = await parseDocxBytes(docxBytes, 'Fallback');
  assert.equal(parsed.pageCount, 2);
  assert.ok(parsed.html.includes('data-page="1"'));
  assert.ok(parsed.html.includes('data-page="2"'));
  assert.ok(parsed.html.includes('<strong>đậm</strong>'));
  assert.ok(parsed.html.includes('<mark> sáng</mark>'));
  assert.ok(parsed.html.includes('<td>Cột 1</td>'));
  console.log('PASS: docx round-trip formatting, tables, lists, pages');
}

async function testWordRenderedPageBreaks() {
  let body = '';
  for (let i = 1; i <= 25; i++) {
    body += `<w:p><w:r><w:t>Dòng ${i}: Kiểm tra ngắt trang chính xác trên iPhone. </w:t></w:r></w:p>`;
  }
  body += `<w:p><w:r><w:t>Cuối trang 1. </w:t></w:r><w:r><w:lastRenderedPageBreak/><w:t>Đầu trang 2.</w:t></w:r></w:p>`;
  body += `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>`;
  const docXml = `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;

  const parsed = parseDocumentXml(docXml, {}, new Set(), 'WordTest');
  assert.equal(parsed.pages.length, 2, 'Expected exact 2 pages when lastRenderedPageBreak is present');
  const p1Text = parsed.pages[0].blocks.flatMap((b) => (b.runs || []).map((r) => r.text)).join('');
  const p2Text = parsed.pages[1].blocks.flatMap((b) => (b.runs || []).map((r) => r.text)).join('');
  assert.ok(p1Text.includes('Cuối trang 1.'), 'Pre-break segment must stay on page 1');
  assert.ok(!p1Text.includes('Đầu trang 2.'), 'Post-break segment must not bleed onto page 1');
  assert.ok(p2Text.includes('Đầu trang 2.'), 'Post-break segment must start page 2');
  console.log('PASS: Word lastRenderedPageBreak mid-paragraph pagination');
}

function testServiceWorkerPrecacheFilesExist() {
  const swSrc = fs.readFileSync(path.join(rootDir, 'sw.js'), 'utf8');
  const m = /PRECACHE_URLS\s*=\s*\[([\s\S]*?)\]/.exec(swSrc);
  assert.ok(m, 'PRECACHE_URLS must exist in sw.js');
  const urls = [...m[1].matchAll(/'([^']+)'/g)].map((x) => x[1]);
  assert.ok(urls.length >= 10, 'Expected full app shell precache list');
  for (const rel of urls) {
    const resolved = rel === './' ? path.join(rootDir, 'index.html') : path.join(rootDir, rel.replace(/^\.\//, ''));
    assert.ok(fs.existsSync(resolved), `Missing precache file on disk: ${rel}`);
  }
  console.log(`PASS: sw.js precache manifest verified (${urls.length} files present)`);
}

function testHtmlDomIdsMatchJs() {
  const html = fs.readFileSync(path.join(rootDir, 'index.html'), 'utf8');
  const htmlIds = new Set([...html.matchAll(/\bid="([^"]+)"/g)].map((m) => m[1]));
  const missing = [];
  for (const rel of ['js/app.js', 'js/reader.js']) {
    const src = fs.readFileSync(path.join(rootDir, rel), 'utf8');
    for (const m of src.matchAll(/\bdocument\.getElementById\('([^']+)'\)/g)) {
      if (!htmlIds.has(m[1])) missing.push(`${rel}:${m[1]}`);
    }
  }
  assert.deepEqual(missing, [], `Missing DOM IDs in index.html: ${missing.join(', ')}`);
  console.log('PASS: all getElementById references exist in index.html');
}

await testZipRoundTripAndSecurity();
await testDocxRoundTrip();
await testWordRenderedPageBreaks();
testServiceWorkerPrecacheFilesExist();
testHtmlDomIdsMatchJs();
