/**
 * UI, cover artwork & formatting utilities matching YunzyBook Android.
 */

export const COVER_PALETTES = [
  { id: 0, name: 'Cobalt Cổ điển', bg: '#2E4EC8', top: '#2E4EC8', bottom: '#2E4EC8', accent: 'rgba(255,255,255,0.26)', accentSoft: 'rgba(255,255,255,0.10)', ink: '#FFFFFF', pattern: 'frame' },
  { id: 1, name: 'Rừng Thông', bg: '#216E57', top: '#216E57', bottom: '#216E57', accent: 'rgba(255,255,255,0.28)', accentSoft: 'rgba(255,255,255,0.12)', ink: '#FFFFFF', pattern: 'arch' },
  { id: 2, name: 'Vải Đỏ Đô', bg: '#7D2940', top: '#7D2940', bottom: '#7D2940', accent: 'rgba(246,232,216,0.38)', accentSoft: 'rgba(255,255,255,0.10)', ink: '#FAF5EE', pattern: 'monograph' },
  { id: 3, name: 'Xám Đá Swiss', bg: '#323E5B', top: '#323E5B', bottom: '#323E5B', accent: 'rgba(255,255,255,0.24)', accentSoft: 'rgba(255,255,255,0.10)', ink: '#FFFFFF', pattern: 'bauhaus' },
  { id: 4, name: 'Giấy Vải Lanh', bg: '#EAE3D2', top: '#EAE3D2', bottom: '#EAE3D2', accent: 'rgba(28,35,56,0.26)', accentSoft: 'rgba(28,35,56,0.06)', ink: '#1E253B', pattern: 'linen' },
  { id: 5, name: 'Xanh Chàm Sâu', bg: '#1D2951', top: '#1D2951', bottom: '#1D2951', accent: 'rgba(255,255,255,0.24)', accentSoft: 'rgba(255,255,255,0.08)', ink: '#FFFFFF', pattern: 'rings' },
  { id: 6, name: 'Lá Xô Thơm', bg: '#3D6B5E', top: '#3D6B5E', bottom: '#3D6B5E', accent: 'rgba(255,255,255,0.28)', accentSoft: 'rgba(255,255,255,0.10)', ink: '#FFFFFF', pattern: 'botanical' },
  { id: 7, name: 'Hổ Phách Trầm', bg: '#A95C29', top: '#A95C29', bottom: '#A95C29', accent: 'rgba(255,248,236,0.30)', accentSoft: 'rgba(255,255,255,0.10)', ink: '#FFFFFF', pattern: 'grid' },
  { id: 8, name: 'Tím Mận Đậm', bg: '#533A7B', top: '#533A7B', bottom: '#533A7B', accent: 'rgba(255,255,255,0.25)', accentSoft: 'rgba(255,255,255,0.10)', ink: '#FFFFFF', pattern: 'diagonal' },
  { id: 9, name: 'Biển Nhật Bản', bg: '#1E5B82', top: '#1E5B82', bottom: '#1E5B82', accent: 'rgba(255,255,255,0.26)', accentSoft: 'rgba(255,255,255,0.10)', ink: '#FFFFFF', pattern: 'waves' },
  { id: 10, name: 'Đất Nung Gạch', bg: '#9F463B', top: '#9F463B', bottom: '#9F463B', accent: 'rgba(255,255,255,0.28)', accentSoft: 'rgba(255,255,255,0.11)', ink: '#FFFFFF', pattern: 'emblem' },
  { id: 11, name: 'Than Chì Tối Giản', bg: '#232836', top: '#232836', bottom: '#232836', accent: 'rgba(255,255,255,0.24)', accentSoft: 'rgba(255,255,255,0.08)', ink: '#FFFFFF', pattern: 'minimal' }
];

export const GROUP_ACCENTS = [
  '#3654E6', '#216E57', '#7D2940', '#A95C29', '#1E5B82', '#533A7B'
];

export function hashColorIndex(text) {
  let h = 0;
  for (let i = 0; i < (text || '').length; i++) {
    h = ((h << 5) - h + text.charCodeAt(i)) | 0;
  }
  return Math.abs(h) % COVER_PALETTES.length;
}

export function buildCoverPatternSvg(pal) {
  const s = pal.accent;
  const f = pal.accentSoft;
  switch (pal.pattern) {
    case 'frame':
      return `<svg viewBox="0 0 100 136" preserveAspectRatio="none" aria-hidden="true">
        <rect x="13" y="10" width="79" height="116" rx="4" fill="none" stroke="${s}" stroke-width="1.2"/>
        <rect x="17" y="14" width="71" height="108" rx="2" fill="none" stroke="${s}" stroke-width="0.6"/>
        <rect x="26" y="28" width="53" height="22" rx="3" fill="${f}"/>
        <line x1="34" y1="68" x2="71" y2="68" stroke="${s}" stroke-width="1.4"/>
        <circle cx="52.5" cy="92" r="9" fill="none" stroke="${s}" stroke-width="1.2"/>
      </svg>`;
    case 'arch':
      return `<svg viewBox="0 0 100 136" preserveAspectRatio="none" aria-hidden="true">
        <path d="M18,126 L18,54 A34,34 0 0,1 86,54 L86,126" fill="${f}" stroke="${s}" stroke-width="1.2"/>
        <circle cx="52" cy="54" r="14" fill="none" stroke="${s}" stroke-width="1.2"/>
        <line x1="28" y1="96" x2="76" y2="96" stroke="${s}" stroke-width="1"/>
      </svg>`;
    case 'monograph':
      return `<svg viewBox="0 0 100 136" preserveAspectRatio="none" aria-hidden="true">
        <line x1="14" y1="16" x2="90" y2="16" stroke="${s}" stroke-width="1.4"/>
        <line x1="14" y1="20" x2="90" y2="20" stroke="${s}" stroke-width="0.6"/>
        <polygon points="52,44 66,64 52,84 38,64" fill="${f}" stroke="${s}" stroke-width="1.2"/>
        <line x1="14" y1="116" x2="90" y2="116" stroke="${s}" stroke-width="0.6"/>
        <line x1="14" y1="120" x2="90" y2="120" stroke="${s}" stroke-width="1.4"/>
      </svg>`;
    case 'bauhaus':
      return `<svg viewBox="0 0 100 136" preserveAspectRatio="none" aria-hidden="true">
        <rect x="10" y="14" width="84" height="44" rx="2" fill="${f}"/>
        <circle cx="64" cy="86" r="24" fill="${f}" stroke="${s}" stroke-width="1.2"/>
        <rect x="18" y="70" width="22" height="46" fill="none" stroke="${s}" stroke-width="1.2"/>
      </svg>`;
    case 'linen':
      return `<svg viewBox="0 0 100 136" preserveAspectRatio="none" aria-hidden="true">
        <rect x="14" y="12" width="76" height="112" rx="3" fill="none" stroke="${s}" stroke-width="1"/>
        <rect x="24" y="26" width="56" height="32" rx="2" fill="${f}" stroke="${s}" stroke-width="0.8"/>
        <line x1="26" y1="80" x2="78" y2="80" stroke="${s}" stroke-width="1"/>
        <line x1="36" y1="88" x2="68" y2="88" stroke="${s}" stroke-width="0.8"/>
      </svg>`;
    case 'rings':
      return `<svg viewBox="0 0 100 136" preserveAspectRatio="none" aria-hidden="true">
        <circle cx="54" cy="64" r="36" fill="none" stroke="${s}" stroke-width="1"/>
        <circle cx="54" cy="64" r="25" fill="${f}" stroke="${s}" stroke-width="1"/>
        <circle cx="54" cy="64" r="13" fill="none" stroke="${s}" stroke-width="1.2"/>
        <line x1="18" y1="114" x2="90" y2="114" stroke="${s}" stroke-width="1"/>
      </svg>`;
    case 'botanical':
      return `<svg viewBox="0 0 100 136" preserveAspectRatio="none" aria-hidden="true">
        <rect x="14" y="12" width="76" height="112" rx="3" fill="none" stroke="${s}" stroke-width="1"/>
        <path d="M52,102 C52,72 52,48 52,32" fill="none" stroke="${s}" stroke-width="1.4"/>
        <path d="M52,76 C36,70 32,56 44,52 C50,50 52,62 52,76 Z" fill="${f}" stroke="${s}" stroke-width="1"/>
        <path d="M52,62 C68,56 72,42 60,38 C54,36 52,48 52,62 Z" fill="${f}" stroke="${s}" stroke-width="1"/>
        <path d="M52,90 C38,86 35,76 45,73" fill="none" stroke="${s}" stroke-width="1"/>
      </svg>`;
    case 'grid':
      return `<svg viewBox="0 0 100 136" preserveAspectRatio="none" aria-hidden="true">
        <path d="M26,10 V126 M50,10 V126 M74,10 V126 M10,34 H94 M10,64 H94 M10,94 H94" stroke="${f}" stroke-width="1"/>
        <rect x="22" y="26" width="60" height="40" rx="4" fill="${f}" stroke="${s}" stroke-width="1.2"/>
      </svg>`;
    case 'diagonal':
      return `<svg viewBox="0 0 100 136" preserveAspectRatio="none" aria-hidden="true">
        <polygon points="10,104 94,42 94,126 10,126" fill="${f}"/>
        <line x1="12" y1="90" x2="92" y2="30" stroke="${s}" stroke-width="1.3"/>
        <circle cx="34" cy="34" r="11" fill="none" stroke="${s}" stroke-width="1.2"/>
      </svg>`;
    case 'waves':
      return `<svg viewBox="0 0 100 136" preserveAspectRatio="none" aria-hidden="true">
        <path d="M0,78 C32,64 62,92 100,74 L100,136 L0,136 Z" fill="${f}"/>
        <path d="M0,98 C35,88 68,112 100,96" fill="none" stroke="${s}" stroke-width="1.3"/>
        <path d="M0,112 C35,102 68,124 100,110" fill="none" stroke="${s}" stroke-width="1"/>
      </svg>`;
    case 'emblem':
      return `<svg viewBox="0 0 100 136" preserveAspectRatio="none" aria-hidden="true">
        <rect x="8" y="0" width="88" height="32" fill="${f}"/>
        <circle cx="54" cy="70" r="21" fill="none" stroke="${s}" stroke-width="1.4"/>
        <circle cx="54" cy="70" r="5" fill="${s}"/>
        <line x1="20" y1="112" x2="88" y2="112" stroke="${s}" stroke-width="1"/>
      </svg>`;
    default: // 'minimal'
      return `<svg viewBox="0 0 100 136" preserveAspectRatio="none" aria-hidden="true">
        <rect x="16" y="22" width="10" height="86" rx="2" fill="${s}"/>
        <line x1="36" y1="34" x2="86" y2="34" stroke="${s}" stroke-width="2"/>
        <line x1="36" y1="44" x2="72" y2="44" stroke="${s}" stroke-width="1.2"/>
        <rect x="36" y="72" width="48" height="32" rx="3" fill="${f}"/>
      </svg>`;
  }
}

export function renderCoverMarkup(book, badgeText = '', overrideIndex = null) {
  if (overrideIndex === null && book?.customCoverDataUrl) {
    return `<div class="book-cover-visual"><img src="${book.customCoverDataUrl}" alt="" />${badgeText ? `<span class="cover-badge">${escapeAttr(badgeText)}</span>` : ''}</div>`;
  }
  const rawIdx = overrideIndex !== null ? overrideIndex : (book?.coverColorIndex || 0);
  const idx = ((rawIdx % COVER_PALETTES.length) + COVER_PALETTES.length) % COVER_PALETTES.length;
  const pal = COVER_PALETTES[idx];
  const shortLabel = escapeAttr((book?.title || 'YunzyBook').trim().slice(0, 26));
  const kindTag = book?.kind === 'pdf' ? 'PDF' : book?.kind === 'docx' ? 'DOCX' : 'SÁCH';
  return `
    <div class="book-cover-visual" style="background:${pal.bg};color:${pal.ink};">
      ${buildCoverPatternSvg(pal)}
      <span class="cover-spine" aria-hidden="true"></span>
      <span class="cover-caption">
        <small>${kindTag}</small>
        <strong>${shortLabel}</strong>
      </span>
      ${badgeText ? `<span class="cover-badge">${escapeAttr(badgeText)}</span>` : ''}
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
    const clean = m[1].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
    if (clean) return clean;
  }
  return extractFirstHeading(html) || fallback;
}

export function extractFirstHeading(html) {
  const m = /<h[1-3][^>]*>([\s\S]*?)<\/h[1-3]>/i.exec(html);
  if (m && m[1]) {
    const clean = m[1].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
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

export function normalizeAppTheme(appThemeMode) {
  const sysDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  if (appThemeMode === 'Tối') return 'dark';
  if (appThemeMode === 'Sepia') return 'sepia';
  if (appThemeMode === 'Sáng') return 'light';
  return sysDark ? 'dark' : 'light';
}

export function downloadBlob(blobOrBytes, filename, mimeType) {
  const blob = blobOrBytes instanceof Blob ? blobOrBytes : new Blob([blobOrBytes], { type: mimeType });
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
