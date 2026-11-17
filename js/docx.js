import { unzipBytes, zipEntries } from './zip.js';

const DEFAULT_PAGE_WIDTH_TWIPS = 11906;
const DEFAULT_PAGE_HEIGHT_TWIPS = 16838;
const DEFAULT_MARGIN_TWIPS = 1440;

const utf8Decoder = new TextDecoder('utf-8');

export function decodeXmlEntities(str) {
  return String(str)
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#([0-9]+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

export function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export function escapeXml(str) {
  return escapeHtml(str).replace(/'/g, '&apos;');
}

/**
 * Pure-JS XML tree parser (works identically in iOS Safari, Chrome, and Node).
 */
export function parseXmlTree(xmlInput) {
  const xml = String(xmlInput)
    .replace(/^\uFEFF/, '')
    .replace(/<!DOCTYPE[^>]*>/gi, '')
    .replace(/<\?[\s\S]*?\?>/g, '')
    .replace(/<!--[\s\S]*?-->/g, '');

  const root = { tag: '#root', attrs: {}, children: [], text: '' };
  const stack = [root];
  const tokenRegex = /<!\[CDATA\[([\s\S]*?)\]\]>|<(\/?)([A-Za-z0-9_:\-.]+)([^>]*)>|([^<]+)/g;

  let m;
  while ((m = tokenRegex.exec(xml)) !== null) {
    if (m[1] !== undefined) {
      stack[stack.length - 1].children.push({ tag: '#text', attrs: {}, children: [], text: m[1] });
      continue;
    }
    if (m[5] !== undefined) {
      const dec = decodeXmlEntities(m[5]);
      stack[stack.length - 1].children.push({ tag: '#text', attrs: {}, children: [], text: dec });
      continue;
    }
    const isClose = m[2] === '/';
    const tag = m[3];
    const attrRaw = m[4] || '';
    const selfClosing = /\/\s*$/.test(attrRaw);

    if (isClose) {
      for (let idx = stack.length - 1; idx > 0; idx--) {
        if (stack[idx].tag === tag) {
          stack.length = idx;
          break;
        }
      }
      continue;
    }

    const attrs = {};
    const attrRegex = /([A-Za-z0-9_:\-.]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    let am;
    while ((am = attrRegex.exec(attrRaw)) !== null) {
      attrs[am[1]] = decodeXmlEntities(am[2] !== undefined ? am[2] : am[3]);
    }
    const node = { tag, attrs, children: [], text: '' };
    stack[stack.length - 1].children.push(node);
    if (!selfClosing) {
      stack.push(node);
    }
  }
  return root.children.find((c) => c.tag !== '#text') || root;
}

function findFirstChild(parent, tagName) {
  if (!parent || !parent.children) return null;
  for (const ch of parent.children) {
    if (ch.tag === tagName) return ch;
  }
  return null;
}

function findDescendants(root, tagName, out = []) {
  if (!root || !root.children) return out;
  for (const ch of root.children) {
    if (ch.tag === tagName) out.push(ch);
    findDescendants(ch, tagName, out);
  }
  return out;
}

function collectText(node) {
  if (!node) return '';
  if (node.tag === '#text') return node.text || '';
  let s = '';
  for (const ch of node.children || []) s += collectText(ch);
  return s;
}

function isToggleOn(el) {
  if (!el) return false;
  const v = el.attrs['w:val'];
  return !v || (v !== '0' && v.toLowerCase() !== 'false');
}

function bytesToBase64(bytes) {
  if (typeof Buffer !== 'undefined' && typeof window === 'undefined') {
    return Buffer.from(bytes).toString('base64');
  }
  let bin = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    bin += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(bin);
}

function base64ToBytes(b64) {
  if (typeof Buffer !== 'undefined' && typeof window === 'undefined') {
    return new Uint8Array(Buffer.from(b64, 'base64'));
  }
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

function mimeForExt(ext) {
  switch (ext.toLowerCase()) {
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'webp':
      return 'image/webp';
    default:
      return 'image/png';
  }
}

function extFromDataUrl(dataUrl) {
  const m = /^data:image\/([a-zA-Z0-9+]+);base64,/.exec(dataUrl || '');
  if (!m) return 'png';
  const sub = m[1].toLowerCase();
  if (sub === 'jpeg') return 'jpg';
  return sub;
}

function parseRelationships(relsXml, mediaDataUrls) {
  if (!relsXml) return {};
  const doc = parseXmlTree(relsXml);
  const rels = findDescendants(doc, 'Relationship');
  const map = {};
  for (const el of rels) {
    const id = el.attrs.Id || '';
    const target = (el.attrs.Target || '').replace(/\\/g, '/');
    if (!id || !target) continue;
    const base = target.split('/').pop();
    const resolved =
      mediaDataUrls[target] ||
      mediaDataUrls[target.replace(/^\.\//, '')] ||
      mediaDataUrls[`word/${target.replace(/^\.\//, '')}`] ||
      mediaDataUrls[base];
    if (resolved) map[id] = resolved;
  }
  return map;
}

function parseOrderedNumIds(numberingXml) {
  const orderedAbstractIds = new Set();
  const orderedNumIds = new Set();
  if (!numberingXml) return orderedNumIds;
  const doc = parseXmlTree(numberingXml);
  for (const abs of findDescendants(doc, 'w:abstractNum')) {
    const absId = abs.attrs['w:abstractNumId'];
    for (const fmt of findDescendants(abs, 'w:numFmt')) {
      const val = fmt.attrs['w:val'] || '';
      if (val === 'decimal' || val === 'lowerLetter' || val === 'upperLetter' || val === 'lowerRoman') {
        orderedAbstractIds.add(absId);
      }
    }
  }
  for (const num of findDescendants(doc, 'w:num')) {
    const numId = num.attrs['w:numId'];
    const absRef = findFirstChild(num, 'w:abstractNumId')?.attrs['w:val'] || '';
    if (orderedAbstractIds.has(absRef)) orderedNumIds.add(numId);
  }
  return orderedNumIds;
}

function extractPageMetrics(bodyNode) {
  const sectPr = findFirstChild(bodyNode, 'w:sectPr');
  const pgSz = findFirstChild(sectPr, 'w:pgSz');
  const pgMar = findFirstChild(sectPr, 'w:pgMar');

  let rawW = parseFloat(pgSz?.attrs['w:w']) || DEFAULT_PAGE_WIDTH_TWIPS;
  let rawH = parseFloat(pgSz?.attrs['w:h']) || DEFAULT_PAGE_HEIGHT_TWIPS;
  if (pgSz?.attrs['w:orient']?.toLowerCase() === 'landscape' && rawW < rawH) {
    [rawW, rawH] = [rawH, rawW];
  }
  const marTop = parseFloat(pgMar?.attrs['w:top']) || DEFAULT_MARGIN_TWIPS;
  const marBottom = parseFloat(pgMar?.attrs['w:bottom']) || DEFAULT_MARGIN_TWIPS;
  const marLeft = parseFloat(pgMar?.attrs['w:left']) || DEFAULT_MARGIN_TWIPS;
  const marRight = parseFloat(pgMar?.attrs['w:right']) || DEFAULT_MARGIN_TWIPS;

  const usableWidthPt = Math.min(950, Math.max(220, (rawW - marLeft - marRight) / 20));
  const usableHeightPt = Math.min(1200, Math.max(280, (rawH - marTop - marBottom) / 20));
  const charsPerLine = Math.min(130, Math.max(45, Math.floor(usableWidthPt / 5.45)));
  return { usableWidthPt, usableHeightPt, charsPerLine };
}

function estimateParagraphHeightPt(tag, plainText, pPr, metrics) {
  const fontSizePt = tag === 'h1' ? 20 : tag === 'h2' ? 16 : tag === 'h3' ? 13.5 : 11.5;
  const indentReduce = (tag === 'ul_li' || tag === 'ol_li') ? 4 : 0;
  const effCharsPerLine = Math.max(24, metrics.charsPerLine * (11.5 / fontSizePt) - indentReduce);

  const spacingEl = findFirstChild(pPr, 'w:spacing');
  const beforePt = Math.min(48, Math.max(0, (parseFloat(spacingEl?.attrs['w:before']) || 0) / 20));
  const afterRaw = parseFloat(spacingEl?.attrs['w:after']);
  const afterPt = Number.isFinite(afterRaw)
    ? Math.min(48, Math.max(0, afterRaw / 20))
    : (tag.startsWith('h') ? 8 : (tag === 'ul_li' || tag === 'ol_li') ? 4 : 6);

  const lineRaw = parseFloat(spacingEl?.attrs['w:line']);
  const lineRule = spacingEl?.attrs['w:lineRule'] || '';
  let linePt = fontSizePt * 1.24;
  if (Number.isFinite(lineRaw) && lineRule === 'exact') {
    linePt = Math.min(48, Math.max(10, lineRaw / 20));
  } else if (Number.isFinite(lineRaw) && lineRaw > 0) {
    linePt = fontSizePt * Math.min(2.5, Math.max(0.95, lineRaw / 240));
  }

  let lines = 1;
  if (plainText) {
    lines = plainText.split('\n').reduce((sum, seg) => (
      sum + Math.max(1, Math.ceil(Math.max(1, seg.length) / effCharsPerLine))
    ), 0);
  }
  return beforePt + lines * linePt + afterPt;
}

function containsPageBreakElement(el) {
  if (findDescendants(el, 'w:lastRenderedPageBreak').length > 0) return true;
  for (const br of findDescendants(el, 'w:br')) {
    if ((br.attrs['w:type'] || '').toLowerCase() === 'page') return true;
  }
  return false;
}

function extractImageSrc(el, relMap) {
  for (const blip of findDescendants(el, 'a:blip')) {
    const rId = blip.attrs['r:embed'];
    if (rId && relMap[rId]) return relMap[rId];
  }
  for (const vImg of findDescendants(el, 'v:imagedata')) {
    const rId = vImg.attrs['r:id'];
    if (rId && relMap[rId]) return relMap[rId];
  }
  return null;
}

function collectCellText(tc) {
  const paras = findDescendants(tc, 'w:p');
  return paras.map((p) => findDescendants(p, 'w:t').map(collectText).join('')).join('\n').trim();
}

export function parseDocumentXml(xml, relMap, orderedNumIds, fallbackTitle) {
  const doc = parseXmlTree(xml);
  const bodyNode = doc.tag === 'w:body' ? doc : findFirstChild(doc, 'w:body');
  if (!bodyNode) {
    return { title: fallbackTitle, firstHeading: '', pages: [{ blocks: [] }] };
  }

  const hasRenderedPageBreaks = xml.includes('lastRenderedPageBreak');
  const metrics = extractPageMetrics(bodyNode);

  const pages = [];
  let currentBlocks = [];
  let currentPageHeightPt = 0;
  let docTitle = '';
  let firstHeading = '';

  function flushPage(allowBlank = false) {
    if (currentBlocks.length > 0) {
      pages.push({ blocks: currentBlocks });
      currentBlocks = [];
      currentPageHeightPt = 0;
    } else if (allowBlank) {
      pages.push({ blocks: [] });
      currentPageHeightPt = 0;
    }
  }

  for (const child of bodyNode.children) {
    if (child.tag === 'w:p') {
      const pPr = findFirstChild(child, 'w:pPr');
      const pStyle = (findFirstChild(pPr, 'w:pStyle')?.attrs['w:val'] || '').toLowerCase();
      const hasPageBreakBefore = findFirstChild(pPr, 'w:pageBreakBefore') !== null;
      const jcVal = (findFirstChild(pPr, 'w:jc')?.attrs['w:val'] || '').toLowerCase();
      const align = jcVal === 'center' ? 'center'
        : jcVal === 'right' ? 'right'
        : (jcVal === 'both' || jcVal === 'distribute') ? 'justify' : '';

      const numPr = findFirstChild(pPr, 'w:numPr');
      const numId = findFirstChild(numPr, 'w:numId')?.attrs['w:val'] || '';
      const isListItem = numPr !== null;
      const isOrdered = orderedNumIds.has(numId);

      const sectPr = findFirstChild(pPr, 'w:sectPr');
      const sectType = (findFirstChild(sectPr, 'w:type')?.attrs['w:val'] || '').toLowerCase();
      const breakAfterSection = sectPr !== null && sectType !== 'continuous';

      if (hasPageBreakBefore && currentBlocks.length > 0) {
        flushPage();
      }

      let activeTag = 'p';
      if (pStyle === 'title' || pStyle === 'heading1' || pStyle === 'heading 1') activeTag = 'h1';
      else if (pStyle === 'subtitle' || pStyle === 'heading2' || pStyle === 'heading 2') activeTag = 'h2';
      else if (pStyle === 'heading3' || pStyle === 'heading 3') activeTag = 'h3';
      else if (isListItem && isOrdered) activeTag = 'ol_li';
      else if (isListItem) activeTag = 'ul_li';

      const runs = [];
      let plainText = '';
      let sawExplicitBreakInPara = false;
      let emittedDrawingInPara = false;

      const handlePageBreak = (explicitManual) => {
        if (runs.length > 0) {
          const seg = plainText.trim();
          if (activeTag === 'h1' && !docTitle && seg) docTitle = seg.slice(0, 80);
          if ((activeTag === 'h1' || activeTag === 'h2') && !firstHeading && seg) firstHeading = seg.slice(0, 60);
          currentBlocks.push({ tag: activeTag, align, runs: runs.slice() });
          runs.length = 0;
          plainText = '';
          if (activeTag === 'ul_li' || activeTag === 'ol_li') activeTag = 'p';
          flushPage();
        } else if (currentBlocks.length > 0) {
          flushPage();
        } else if (explicitManual) {
          flushPage(true);
        }
        sawExplicitBreakInPara = true;
      };

      const visitParaNode = (el) => {
        if (!el || el.tag === '#text') return;
        if (el.tag === 'w:lastRenderedPageBreak') {
          handlePageBreak(false);
        } else if (el.tag === 'w:br') {
          if ((el.attrs['w:type'] || '').toLowerCase() === 'page') handlePageBreak(true);
          else runs.push({ text: '\n' });
        } else if (el.tag === 'w:r') {
          const rPr = findFirstChild(el, 'w:rPr');
          const bold = isToggleOn(findFirstChild(rPr, 'w:b'));
          const italic = isToggleOn(findFirstChild(rPr, 'w:i'));
          const uEl = findFirstChild(rPr, 'w:u');
          const underline = uEl !== null && uEl.attrs['w:val'] !== 'none';
          const strike = isToggleOn(findFirstChild(rPr, 'w:strike'));
          const highlight = findFirstChild(rPr, 'w:highlight') !== null;

          for (const rc of el.children) {
            if (rc.tag === 'w:lastRenderedPageBreak') {
              handlePageBreak(false);
            } else if (rc.tag === 'w:br') {
              if ((rc.attrs['w:type'] || '').toLowerCase() === 'page') handlePageBreak(true);
              else runs.push({ text: '\n' });
            } else if (rc.tag === 'w:tab') {
              runs.push({ text: '    ' });
            } else if (rc.tag === 'w:t') {
              const t = collectText(rc);
              if (t.length > 0) {
                plainText += t;
                runs.push({ text: t, bold, italic, underline, strike, highlight });
              }
            } else if (rc.tag === 'w:drawing' || rc.tag === 'w:pict') {
              const imgSrc = extractImageSrc(rc, relMap);
              if (imgSrc) {
                if (runs.length > 0) {
                  currentBlocks.push({ tag: activeTag, align, runs: runs.slice() });
                  currentPageHeightPt += estimateParagraphHeightPt(activeTag, plainText, pPr, metrics);
                  runs.length = 0;
                  plainText = '';
                }
                if (!hasRenderedPageBreaks && currentBlocks.length > 0 && currentPageHeightPt + 210 > metrics.usableHeightPt) {
                  flushPage();
                }
                currentBlocks.push({ tag: 'img', align, src: imgSrc });
                currentPageHeightPt += 210;
                emittedDrawingInPara = true;
              }
            }
          }
        } else if (el.tag === 'w:hyperlink' || el.tag === 'w:sdt' || el.tag === 'w:sdtContent' || el.tag === 'w:ins') {
          for (const ch of el.children) visitParaNode(ch);
        }
      };

      for (const ch of child.children) visitParaNode(ch);

      const trimmedPlain = plainText.trim();
      if (activeTag === 'h1' && !docTitle && trimmedPlain) docTitle = trimmedPlain.slice(0, 80);
      if ((activeTag === 'h1' || activeTag === 'h2') && !firstHeading && trimmedPlain) firstHeading = trimmedPlain.slice(0, 60);

      const shouldKeepBlock = runs.length > 0 || (!sawExplicitBreakInPara && !emittedDrawingInPara);
      if (shouldKeepBlock) {
        const estHeight = estimateParagraphHeightPt(activeTag, plainText, pPr, metrics);
        if (!hasRenderedPageBreaks && currentBlocks.length > 0) {
          const overflow = currentPageHeightPt + estHeight > metrics.usableHeightPt;
          const orphanHeading = activeTag.startsWith('h') && currentPageHeightPt > metrics.usableHeightPt * 0.86;
          if (overflow || orphanHeading) flushPage();
        }
        currentBlocks.push({ tag: activeTag, align, runs });
        currentPageHeightPt += estHeight;
      }

      if (breakAfterSection && currentBlocks.length > 0) {
        flushPage();
      }
    } else if (child.tag === 'w:tbl') {
      const rows = [];
      for (const tr of child.children) {
        if (tr.tag !== 'w:tr') continue;
        if (containsPageBreakElement(tr)) {
          if (rows.length > 0) {
            currentBlocks.push({ tag: 'table', rows: rows.slice() });
            rows.length = 0;
          }
          if (currentBlocks.length > 0) flushPage();
        }
        const rowCells = [];
        for (const tc of tr.children) {
          if (tc.tag === 'w:tc') rowCells.push(collectCellText(tc));
        }
        if (rowCells.length > 0) {
          const maxCellLen = Math.max(0, ...rowCells.map((c) => c.length));
          const rowEstPt = 18 + Math.ceil(Math.max(1, maxCellLen) / 28) * 14;
          if (!hasRenderedPageBreaks && (currentBlocks.length > 0 || rows.length > 0) && currentPageHeightPt + rowEstPt > metrics.usableHeightPt) {
            if (rows.length > 0) {
              currentBlocks.push({ tag: 'table', rows: rows.slice() });
              rows.length = 0;
            }
            flushPage();
          }
          rows.push(rowCells);
          currentPageHeightPt += rowEstPt;
        }
      }
      if (rows.length > 0) {
        currentBlocks.push({ tag: 'table', rows });
      }
    }
  }

  flushPage();
  if (pages.length === 0) pages.push({ blocks: [] });
  if (!docTitle) docTitle = fallbackTitle;
  return { title: docTitle, firstHeading, pages };
}

function renderRunsHtml(runs) {
  if (!runs || runs.length === 0) return '';
  let out = '';
  for (const r of runs) {
    const esc = escapeHtml(r.text || '').replace(/\n/g, '<br/>');
    if (!esc) continue;
    if (r.highlight) out += '<mark>';
    if (r.bold) out += '<strong>';
    if (r.italic) out += '<em>';
    if (r.underline) out += '<u>';
    if (r.strike) out += '<s>';
    out += esc;
    if (r.strike) out += '</s>';
    if (r.underline) out += '</u>';
    if (r.italic) out += '</em>';
    if (r.bold) out += '</strong>';
    if (r.highlight) out += '</mark>';
  }
  return out;
}

function renderBlocksHtml(blocks) {
  let out = '';
  let inUl = false;
  let inOl = false;

  const closeLists = () => {
    if (inUl) { out += '</ul>\n'; inUl = false; }
    if (inOl) { out += '</ol>\n'; inOl = false; }
  };

  for (const b of blocks || []) {
    if (b.tag === 'ul_li') {
      if (inOl) { out += '</ol>\n'; inOl = false; }
      if (!inUl) { out += '<ul>\n'; inUl = true; }
      const st = b.align ? ` style="text-align:${b.align}"` : '';
      out += `<li${st}>${renderRunsHtml(b.runs) || '<br/>'}</li>\n`;
      continue;
    }
    if (b.tag === 'ol_li') {
      if (inUl) { out += '</ul>\n'; inUl = false; }
      if (!inOl) { out += '<ol>\n'; inOl = true; }
      const st = b.align ? ` style="text-align:${b.align}"` : '';
      out += `<li${st}>${renderRunsHtml(b.runs) || '<br/>'}</li>\n`;
      continue;
    }

    closeLists();
    if (b.tag === 'table') {
      out += '<table>\n';
      for (const row of b.rows || []) {
        out += '<tr>';
        for (const cell of row) {
          out += `<td>${escapeHtml(cell).replace(/\n/g, '<br/>')}</td>`;
        }
        out += '</tr>\n';
      }
      out += '</table>\n';
    } else if (b.tag === 'img') {
      const cls = b.align ? ` style="text-align:${b.align}"` : '';
      out += `<p${cls}><img src="${b.src}" alt="Hình ảnh" /></p>\n`;
    } else {
      const safeTag = (b.tag === 'h1' || b.tag === 'h2' || b.tag === 'h3') ? b.tag : 'p';
      const st = b.align ? ` style="text-align:${b.align}"` : '';
      const inner = renderRunsHtml(b.runs);
      out += `<${safeTag}${st}>${inner || '<br/>'}</${safeTag}>\n`;
    }
  }
  closeLists();
  return out;
}

export function renderDocxPagesToHtml(title, rawPages) {
  const pages = (rawPages && rawPages.length > 0)
    ? rawPages
    : [{ blocks: [{ tag: 'p', align: '', runs: [{ text: '' }] }] }];
  const totalPages = pages.length;

  const sections = pages.map((page, idx) => {
    const num = idx + 1;
    const active = idx === 0 ? ' active-page' : '';
    return `<section class="docx-page${active}" data-page="${num}">\n<div class="docx-page-body">\n${renderBlocksHtml(page.blocks)}</div>\n<div class="docx-page-footer">Trang ${num} / ${totalPages}</div>\n</section>`;
  }).join('\n');

  return `<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8" />
<meta name="yunzy-docx-engine" content="2" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>${escapeHtml(title)}</title>
<style>
:root {
  --page-bg: #FFFFFF;
  --page-fg: #1E2433;
  --page-muted: #6B748F;
  --page-border: rgba(54, 84, 230, 0.12);
  --page-shadow: 0 6px 24px rgba(24, 39, 75, 0.08);
  --doc-backdrop: #EAEEF7;
}
body {
  font-family: -apple-system, BlinkMacSystemFont, system-ui, sans-serif;
  line-height: 1.68;
  margin: 0;
  padding: 18px 12px 92px;
  background: var(--doc-backdrop);
  color: var(--page-fg);
  word-wrap: break-word;
  -webkit-text-size-adjust: 100%;
}
.docx-pages { display: flex; flex-direction: column; gap: 16px; max-width: 780px; margin: 0 auto; }
.docx-page {
  background: var(--page-bg);
  color: var(--page-fg);
  border: 1px solid var(--page-border);
  border-radius: 14px;
  box-shadow: var(--page-shadow);
  padding: 22px 18px 16px;
  min-height: 60vh;
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  box-sizing: border-box;
}
.docx-page-body { outline: none; flex: 1; }
.docx-page-body[contenteditable="true"]:focus { caret-color: #3654E6; }
.docx-page-footer {
  margin-top: 18px; padding-top: 10px;
  border-top: 1px dashed rgba(120, 134, 172, 0.22);
  text-align: center; font-size: 11.5px; font-weight: 600;
  color: var(--page-muted); user-select: none;
}
body[data-page-mode="paginated"] { padding-bottom: 104px; }
body[data-page-mode="paginated"] .docx-page { display: none; min-height: calc(100vh - 190px); }
body[data-page-mode="paginated"] .docx-page.active-page { display: flex; }
body[data-editing="true"] .docx-page { border-color: rgba(54, 84, 230, 0.34); }
h1 { font-size: 1.52em; line-height: 1.3; margin: 0.7em 0 0.45em; }
h2 { font-size: 1.28em; line-height: 1.35; margin: 0.65em 0 0.4em; }
h3 { font-size: 1.12em; line-height: 1.4; margin: 0.55em 0 0.35em; }
p { margin: 0.48em 0; min-height: 1.1em; }
ul, ol { margin: 0.45em 0; padding-left: 1.45em; }
li { margin: 0.25em 0; }
mark { background: #FEF08A; color: #1E2433; padding: 0 2px; border-radius: 3px; }
img { max-width: 100%; height: auto; border-radius: 8px; margin: 0.55em 0; display: block; }
table { border-collapse: collapse; width: 100%; margin: 0.75em 0; font-size: 0.94em; }
td, th { border: 1px solid rgba(128,128,128,0.35); padding: 6px 8px; vertical-align: top; }
</style>
</head>
<body data-doc-kind="docx" data-page-count="${totalPages}">
<div class="docx-pages">
${sections}
</div>
</body>
</html>`;
}

export async function parseDocxBytes(bytes, fallbackTitle = 'Tài liệu Word') {
  const entries = await unzipBytes(bytes);
  const docBytes = entries.get('word/document.xml');
  if (!docBytes) {
    throw new Error('File DOCX không hợp lệ (thiếu word/document.xml).');
  }
  const documentXml = utf8Decoder.decode(docBytes);
  const relsBytes = entries.get('word/_rels/document.xml.rels');
  const relsXml = relsBytes ? utf8Decoder.decode(relsBytes) : '';
  const numBytes = entries.get('word/numbering.xml');
  const numberingXml = numBytes ? utf8Decoder.decode(numBytes) : '';

  const mediaDataUrls = {};
  for (const [name, data] of entries.entries()) {
    if (name.startsWith('word/media/')) {
      const ext = name.split('.').pop() || 'png';
      const dataUrl = `data:${mimeForExt(ext)};base64,${bytesToBase64(data)}`;
      mediaDataUrls[name] = dataUrl;
      mediaDataUrls[name.replace(/^word\//, '')] = dataUrl;
      mediaDataUrls[name.split('/').pop()] = dataUrl;
    }
  }

  const relMap = parseRelationships(relsXml, mediaDataUrls);
  const orderedNumIds = parseOrderedNumIds(numberingXml);
  const parsed = parseDocumentXml(documentXml, relMap, orderedNumIds, fallbackTitle);
  const html = renderDocxPagesToHtml(parsed.title, parsed.pages);

  return {
    title: parsed.title,
    firstHeading: parsed.firstHeading,
    pages: parsed.pages,
    pageCount: parsed.pages.length,
    html
  };
}

export async function buildDocxBytesFromPages(rawPages) {
  const pages = (rawPages && rawPages.length > 0)
    ? rawPages
    : [{ blocks: [{ tag: 'p', align: '', runs: [{ text: '' }] }] }];

  const imageMap = new Map(); // dataUrl -> { rId, name, bytes, ext }
  let relIdx = 1;
  for (const p of pages) {
    for (const b of p.blocks || []) {
      if (b.tag === 'img' && b.src && b.src.startsWith('data:image/') && !imageMap.has(b.src)) {
        const b64 = b.src.split(',')[1] || '';
        const ext = extFromDataUrl(b.src);
        const rId = `rId${relIdx}`;
        const name = `image_${relIdx}.${ext}`;
        imageMap.set(b.src, { rId, name, ext, bytes: base64ToBytes(b64) });
        relIdx++;
      }
    }
  }

  const usedExts = new Set([...imageMap.values()].map((i) => i.ext));
  let contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>\n`;
  if (usedExts.has('png')) contentTypes += `  <Default Extension="png" ContentType="image/png"/>\n`;
  if (usedExts.has('jpg') || usedExts.has('jpeg')) {
    contentTypes += `  <Default Extension="jpg" ContentType="image/jpeg"/>\n  <Default Extension="jpeg" ContentType="image/jpeg"/>\n`;
  }
  if (usedExts.has('gif')) contentTypes += `  <Default Extension="gif" ContentType="image/gif"/>\n`;
  if (usedExts.has('webp')) contentTypes += `  <Default Extension="webp" ContentType="image/webp"/>\n`;
  contentTypes += `  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  let docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n`;
  for (const img of imageMap.values()) {
    docRels += `  <Relationship Id="${img.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${escapeXml(img.name)}"/>\n`;
  }
  docRels += `</Relationships>`;

  let docPrId = 1;
  let bodyXml = '';
  const jcMap = { center: 'center', right: 'right', justify: 'both' };

  pages.forEach((page, pageIdx) => {
    if (pageIdx > 0) bodyXml += `    <w:p><w:r><w:br w:type="page"/></w:r></w:p>\n`;
    const blocks = (page.blocks && page.blocks.length > 0) ? page.blocks : [{ tag: 'p', align: '', runs: [{ text: '' }] }];
    for (const block of blocks) {
      if (block.tag === 'table' && block.rows?.length) {
        bodyXml += `    <w:tbl>\n      <w:tblPr><w:tblBorders>`;
        for (const s of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
          bodyXml += `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>`;
        }
        bodyXml += `</w:tblBorders></w:tblPr>\n`;
        for (const row of block.rows) {
          bodyXml += `      <w:tr>`;
          for (const cell of row) {
            bodyXml += `<w:tc><w:p><w:r><w:t xml:space="preserve">${escapeXml(cell)}</w:t></w:r></w:p></w:tc>`;
          }
          bodyXml += `</w:tr>\n`;
        }
        bodyXml += `    </w:tbl>\n`;
      } else if (block.tag === 'img' && imageMap.has(block.src)) {
        const img = imageMap.get(block.src);
        const idNum = docPrId++;
        const emuW = 4572000;
        const emuH = 3048000;
        const pPr = jcMap[block.align] ? `<w:pPr><w:jc w:val="${jcMap[block.align]}"/></w:pPr>` : '';
        bodyXml += `    <w:p>${pPr}<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${emuW}" cy="${emuH}"/><wp:docPr id="${idNum}" name="Image ${idNum}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${idNum}" name="Image ${idNum}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${img.rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emuW}" cy="${emuH}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></wp:drawing></w:r></w:p>\n`;
      } else {
        let pPr = '';
        if (block.tag === 'h1') pPr += '<w:pStyle w:val="Heading1"/>';
        else if (block.tag === 'h2') pPr += '<w:pStyle w:val="Heading2"/>';
        else if (block.tag === 'h3') pPr += '<w:pStyle w:val="Heading3"/>';
        else if (block.tag === 'ul_li') pPr += '<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
        else if (block.tag === 'ol_li') pPr += '<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>';
        if (jcMap[block.align]) pPr += `<w:jc w:val="${jcMap[block.align]}"/>`;

        bodyXml += `    <w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}`;
        if (block.tag === 'ul_li') bodyXml += `<w:r><w:t xml:space="preserve">• </w:t></w:r>`;

        const runs = (block.runs && block.runs.length > 0) ? block.runs : [{ text: '' }];
        for (const run of runs) {
          let rPr = '';
          if (block.tag === 'h1') rPr += '<w:b/><w:sz w:val="36"/>';
          else if (block.tag === 'h2') rPr += '<w:b/><w:sz w:val="30"/>';
          else if (block.tag === 'h3') rPr += '<w:b/><w:sz w:val="26"/>';
          if (run.bold && !String(block.tag).startsWith('h')) rPr += '<w:b/>';
          if (run.italic) rPr += '<w:i/>';
          if (run.underline) rPr += '<w:u w:val="single"/>';
          if (run.strike) rPr += '<w:strike/>';
          if (run.highlight) rPr += '<w:highlight w:val="yellow"/>';

          bodyXml += `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}`;
          const lines = String(run.text || '').split('\n');
          lines.forEach((seg, sIdx) => {
            if (sIdx > 0) bodyXml += '<w:br/>';
            bodyXml += `<w:t xml:space="preserve">${escapeXml(seg)}</w:t>`;
          });
          bodyXml += `</w:r>`;
        }
        bodyXml += `</w:p>\n`;
      }
    }
  });

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
${bodyXml}    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>
  </w:body>
</w:document>`;

  const entries = [
    ['[Content_Types].xml', contentTypes],
    ['_rels/.rels', rootRels],
    ['word/_rels/document.xml.rels', docRels],
    ['word/document.xml', documentXml]
  ];
  for (const img of imageMap.values()) {
    entries.push([`word/media/${img.name}`, img.bytes]);
  }
  return zipEntries(entries);
}
