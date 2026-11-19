/**
 * UI, cover artwork & formatting utilities matching YunzyBook Android.
 */

export const COVER_PALETTES = [
  { id: 0, name: 'Xanh dương · vòng tròn', bg: '#2F4FD0', top: '#2F4FD0', bottom: '#253FA8', pattern: 'circles-right' },
  { id: 1, name: 'Tím than · sọc chéo', bg: '#1E2A4A', top: '#1E2A4A', bottom: '#151E36', pattern: 'stripes-135' },
  { id: 2, name: 'Rượu vang · chấm', bg: '#7A2438', top: '#7A2438', bottom: '#5D1A29', pattern: 'dots' },
  { id: 3, name: 'Xanh ngọc · cung tròn', bg: '#0F6B5E', top: '#0F6B5E', bottom: '#0B5248', pattern: 'concentric-tr' },
  { id: 4, name: 'Tím · vòng tròn', bg: '#5B3FA8', top: '#5B3FA8', bottom: '#473086', pattern: 'circles-left' },
  { id: 5, name: 'Xanh biển · sọc chéo', bg: '#0E6A8A', top: '#0E6A8A', bottom: '#0A516A', pattern: 'stripes-45' },
  { id: 6, name: 'Cam đất · chấm', bg: '#B4470F', top: '#B4470F', bottom: '#8C360A', pattern: 'dots' },
  { id: 7, name: 'Hồng tím · cung tròn', bg: '#9A3B7A', top: '#9A3B7A', bottom: '#792C5F', pattern: 'concentric-tr' }
];

export const GROUP_ACCENTS = [
  '#2F4FD0', '#0F6B5E', '#7A2438', '#B4470F', '#0E6A8A', '#5B3FA8'
];

export function hashColorIndex(text) {
  let h = 0;
  for (let i = 0; i < (text || '').length; i++) {
    h = ((h << 5) - h + text.charCodeAt(i)) | 0;
  }
  return Math.abs(h) % COVER_PALETTES.length;
}

export function buildHeaderPatternMarkup(pattern) {
  switch (pattern) {
    case 'circles-right':
      return '<span class="pat-c pat-cr-1"></span><span class="pat-c pat-cr-2"></span>';
    case 'circles-left':
      return '<span class="pat-c pat-cl-1"></span><span class="pat-c pat-cl-2"></span>';
    case 'concentric-tr':
      return '<span class="pat-ring pat-ring-outer"></span><span class="pat-ring pat-ring-inner"></span>';
    default:
      return '';
  }
}

export function renderCoverMarkup(book, badgeText = '', overrideIndex = null, options = {}) {
  if (overrideIndex === null && book?.customCoverDataUrl) {
    return `<div class="book-cover-visual"><img src="${book.customCoverDataUrl}" alt="" />${badgeText ? `<span class="cover-badge">${escapeAttr(badgeText)}</span>` : ''}</div>`;
  }
  const rawIdx = overrideIndex !== null ? overrideIndex : (book?.coverColorIndex || 0);
  const idx = ((rawIdx % COVER_PALETTES.length) + COVER_PALETTES.length) % COVER_PALETTES.length;
  const pal = COVER_PALETTES[idx];
  const fullTitle = (book?.title || 'YunzyBook').trim();
  const lenCls = fullTitle.length <= 18 ? 'len-sm' : fullTitle.length <= 36 ? 'len-md' : 'len-lg';
  const kindTag = book?.kind === 'pdf' ? 'PDF' : book?.kind === 'docx' ? 'DOCX' : 'SÁCH';
  const rawSub = (book?.groupName || book?.subtitle || '').trim();
  const subText = rawSub && rawSub.toLowerCase() !== fullTitle.toLowerCase() ? rawSub : '';
  const miniCls = options.isMiniature ? ' is-mini' : '';

  return `
    <div class="gen-cover${miniCls}">
      <div class="gen-cover-top pat-${pal.pattern}" style="background-color:${pal.bg};">
        ${buildHeaderPatternMarkup(pal.pattern)}
        <span class="gen-cover-tag">${kindTag}</span>
      </div>
      <div class="gen-cover-body">
        <span class="gen-cover-bar" style="background:${pal.bg};"></span>
        <div class="gen-cover-title ${lenCls}">${escapeAttr(fullTitle)}</div>
        ${(subText || badgeText) ? `
          <div class="gen-cover-footer">
            ${subText ? `<span class="gen-cover-sub">${escapeAttr(subText)}</span>` : ''}
            ${badgeText ? `<span class="gen-cover-status">${escapeAttr(badgeText)}</span>` : ''}
          </div>` : ''}
      </div>
      <span class="gen-cover-spine" aria-hidden="true"></span>
    </div>
  `;
}

export function escapeAttr(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export function formatBytesMb(bytes) {
  const mb = Math.max(0, Number(bytes || 0)) / (1024 * 1024);
  return `${mb.toFixed(1)} MB`;
}

export function extractHtmlTitle(html, fallback) {
  const m = /<title[^>]*>([\s\S]*?)<\/title>/i.exec(html);
  if (m && m[1]) {
    const clean = m[1].replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
    if (clean) return clean;
  }
  return extractFirstHeading(html) || fallback;
}

export function extractFirstHeading(html) {
  const m = /<h[1-3][^>]*>([\s\S]*?)<\/h[1-3]>/i.exec(html);
  if (m && m[1]) {
    const clean = m[1].replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
    if (clean) return clean.slice(0, 60);
  }
  return '';
}

export function isIosSafari() {
  const ua = navigator.userAgent || '';
  const isIos = /iPad|iPhone|iPod/.test(ua) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  const isStandalone = window.navigator.standalone === true || window.matchMedia('(display-mode: standalone)').matches;
  return { isIos, isStandalone };
}

let hapticEnabledState = true;
let lastHapticAt = 0;
let hapticSwitchLabel = null;
let hapticTimers = [];

function ensureIosHapticHost() {
  if (typeof document === 'undefined' || !document.body) return null;
  if (hapticSwitchLabel && hapticSwitchLabel.isConnected) return hapticSwitchLabel;

  let host = document.getElementById('yunzy-haptic-host');
  if (!host) {
    host = document.createElement('div');
    host.id = 'yunzy-haptic-host';
    host.setAttribute('aria-hidden', 'true');
    // Must stay in render tree (no display:none / visibility:hidden) and fixed at 0,0 to prevent scroll jumps
    host.style.cssText =
      'position:fixed;top:0;left:0;width:1px;height:1px;opacity:0;overflow:hidden;pointer-events:none;z-index:-1;';
    host.addEventListener('click', (e) => e.stopPropagation());

    const input = document.createElement('input');
    input.id = 'yunzy-haptic-switch';
    input.type = 'checkbox';
    input.tabIndex = -1;
    input.setAttribute('switch', '');
    input.style.cssText =
      'position:fixed;top:0;left:0;width:1px;height:1px;opacity:0;pointer-events:none;margin:0;padding:0;';

    const label = document.createElement('label');
    label.id = 'yunzy-haptic-label';
    label.htmlFor = 'yunzy-haptic-switch';
    label.style.cssText =
      'position:fixed;top:0;left:0;width:1px;height:1px;opacity:0;pointer-events:none;margin:0;padding:0;';

    host.appendChild(input);
    host.appendChild(label);
    document.body.appendChild(host);
    hapticSwitchLabel = label;
  } else {
    hapticSwitchLabel = document.getElementById('yunzy-haptic-label');
  }
  return hapticSwitchLabel;
}

function pulseIosSwitch() {
  try {
    const label = ensureIosHapticHost();
    if (label) label.click();
  } catch (_) {}
}

function clearPendingHapticTimers() {
  for (const id of hapticTimers) clearTimeout(id);
  hapticTimers = [];
}

export function setHapticEnabled(enabled) {
  hapticEnabledState = Boolean(enabled);
  if (!hapticEnabledState) clearPendingHapticTimers();
}

export function initHaptics(enabled = true) {
  hapticEnabledState = Boolean(enabled);
  if (typeof window === 'undefined') return;
  if (isIosSafari().isIos) {
    ensureIosHapticHost();
  }
}

export function triggerHaptic(style = 'light') {
  if (!hapticEnabledState || typeof window === 'undefined') return;
  const now = Date.now();
  if (now - lastHapticAt < 28) return;
  lastHapticAt = now;

  // Priority 1: Native Android WebView bridge
  if (typeof window.AndroidBridge?.haptic === 'function') {
    try {
      window.AndroidBridge.haptic(style);
    } catch (_) {}
    return;
  }

  // Priority 2: iOS 17.4+ / 18+ WebKit native switch Taptic Engine
  if (isIosSafari().isIos) {
    clearPendingHapticTimers();
    pulseIosSwitch();
    if (style === 'success') {
      hapticTimers.push(setTimeout(pulseIosSwitch, 85));
    } else if (style === 'warning' || style === 'error') {
      hapticTimers.push(setTimeout(pulseIosSwitch, 85));
      hapticTimers.push(setTimeout(pulseIosSwitch, 170));
    }
    return;
  }

  // Priority 3: Standard Vibration API fallback (Android Chrome / mobile web)
  if (typeof navigator?.vibrate === 'function') {
    const patternMap = {
      selection: 6,
      light: 9,
      medium: 14,
      heavy: 24,
      success: [10, 35, 16],
      warning: [14, 30, 12],
      error: [18, 30, 18, 30, 22]
    };
    try {
      navigator.vibrate(patternMap[style] ?? 9);
    } catch (_) {}
  }
}

export function normalizeAppTheme(appThemeMode) {
  const sysDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  if (appThemeMode === 'Tối') return 'dark';
  if (appThemeMode === 'Sepia') return 'sepia';
  if (appThemeMode === 'Sáng') return 'light';
  return sysDark ? 'dark' : 'light';
}

export function downloadBlob(blobOrBytes, filename, mimeType) {
  const blob = blobOrBytes instanceof Blob ? blobOrBytes : new Blob([blobOrBytes], { type: mimeType });
  if (typeof window !== 'undefined' && window.AndroidBridge && typeof window.AndroidBridge.saveFileBase64 === 'function') {
    const fr = new FileReader();
    fr.onload = () => {
      const raw = String(fr.result || '');
      const commaIdx = raw.indexOf(',');
      const b64 = commaIdx >= 0 ? raw.slice(commaIdx + 1) : raw;
      window.AndroidBridge.saveFileBase64(b64, filename, mimeType || blob.type || 'application/octet-stream');
    };
    fr.readAsDataURL(blob);
    return;
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    a.remove();
    URL.revokeObjectURL(url);
  }, 1500);
}

function readFileAsDataUrlFallback(file) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.onerror = () => reject(fr.error || new Error('Không đọc được ảnh bìa'));
    fr.readAsDataURL(file);
  });
}

export async function compressCoverImage(file, maxW = 360, maxH = 486, quality = 0.82) {
  if (!file) throw new Error('Thiếu tệp ảnh');
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((resolve, reject) => {
      const el = new Image();
      el.onload = () => resolve(el);
      el.onerror = () => reject(new Error('Không giải mã được ảnh'));
      el.src = url;
    });
    const srcW = Math.max(1, img.naturalWidth || img.width || maxW);
    const srcH = Math.max(1, img.naturalHeight || img.height || maxH);
    const scale = Math.min(1, maxW / srcW, maxH / srcH);
    const outW = Math.max(1, Math.round(srcW * scale));
    const outH = Math.max(1, Math.round(srcH * scale));

    const canvas = document.createElement('canvas');
    canvas.width = outW;
    canvas.height = outH;
    const ctx = canvas.getContext('2d');
    if (!ctx) return await readFileAsDataUrlFallback(file);
    ctx.drawImage(img, 0, 0, outW, outH);
    return canvas.toDataURL('image/jpeg', quality);
  } catch (_) {
    return await readFileAsDataUrlFallback(file);
  } finally {
    URL.revokeObjectURL(url);
  }
}
