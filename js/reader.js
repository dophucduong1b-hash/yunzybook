import { buildDocxBytesFromPages, renderDocxPagesToHtml } from './docx.js';
import { downloadBlob, escapeAttr, normalizeAppTheme } from './ui-helpers.js';

export const SESSION_STORAGE_KEY = 'yunzy_reading_session';

let pdfjsPromise = null;
async function loadPdfJs() {
  if (!pdfjsPromise) {
    pdfjsPromise = import('../vendor/pdfjs/pdf.min.mjs').then((mod) => {
      mod.GlobalWorkerOptions.workerSrc = './vendor/pdfjs/pdf.worker.min.mjs';
      return mod;
    });
  }
  return pdfjsPromise;
}

const FONT_CSS_MAP = {
  ORIGINAL: '',
  MODERN: "-apple-system, BlinkMacSystemFont, 'SF Pro Text', system-ui, sans-serif",
  CLASSIC: "Georgia, 'Times New Roman', serif",
  SYSTEM: "'SF Mono', ui-monospace, Menlo, monospace"
};

const LINE_SPACING_MAP = {
  COMPACT: 1.38,
  NORMAL: 1.65,
  RELAXED: 1.98
};

let safeProbeEl = null;

export function syncSafeAreaMetrics() {
  if (typeof document === 'undefined') return 0;
  if (!safeProbeEl) {
    safeProbeEl = document.createElement('div');
    safeProbeEl.id = '__yunzy_safe_probe';
    safeProbeEl.style.cssText =
      'position:fixed;top:0;left:0;width:0;height:0;padding-top:var(--safe-top, env(safe-area-inset-top, 0px));visibility:hidden;pointer-events:none;z-index:-1;';
    (document.body || document.documentElement).appendChild(safeProbeEl);
  }
  return parseFloat(getComputedStyle(safeProbeEl).paddingTop) || 0;
}

export function createReaderController({
  getSettings,
  onUpdateSettings,
  getBookContent,
  onUpdateBookProgress,
  onUpdateBookMeta,
  onSaveBookContent,
  onCreateCopyBook,
  onShowToast,
  onCloseReader
}) {
  const overlay = document.getElementById('reader-overlay');
  const stageEl = document.getElementById('reader-stage');
  const topBarRead = document.getElementById('reader-topbar-read');
  const topBarEdit = document.getElementById('reader-topbar-edit');
  const titleEl = document.getElementById('reader-book-title');
  const subEl = document.getElementById('reader-book-sub');
  const editTitleEl = document.getElementById('editor-book-title');
  const editStatusEl = document.getElementById('editor-status-text');

  const btnBack = document.getElementById('reader-btn-back');
  const btnFloatingBack = document.getElementById('reader-floating-back');
  const btnTts = document.getElementById('reader-btn-tts');
  const ttsLabel = document.getElementById('reader-tts-label');
  const btnEdit = document.getElementById('reader-btn-edit');
  const btnAa = document.getElementById('reader-btn-aa');
  const btnSearch = document.getElementById('reader-btn-search');
  const btnMore = document.getElementById('reader-btn-more');
  const moreDropdown = document.getElementById('reader-more-dropdown');
  const btnBookmark = document.getElementById('reader-btn-bookmark');
  const bookmarkLabel = document.getElementById('reader-bookmark-label');

  const aaBackdrop = document.getElementById('aa-sheet-backdrop');
  const aaSheet = document.getElementById('aa-sheet');
  const searchOverlay = document.getElementById('reader-search-overlay');
  const searchInput = document.getElementById('reader-search-input');
  const searchList = document.getElementById('reader-search-results');
  const searchCountEl = document.getElementById('reader-search-count');

  const pageNavBar = document.getElementById('reader-page-nav');
  const pageNavText = document.getElementById('reader-page-nav-text');
  const pageNavPrev = document.getElementById('reader-page-prev');
  const pageNavNext = document.getElementById('reader-page-next');

  const edgeLeft = document.getElementById('reader-edge-left');
  const edgeRight = document.getElementById('reader-edge-right');

  const editorRibbon = document.getElementById('word-editor-ribbon');
  const editMenuEl = document.getElementById('editor-save-dropdown');

  let currentBook = null;
  let activeIframe = null;
  let activeBlobUrl = null;
  let dockHidden = false;
  let pdfDoc = null;
  let pdfObserver = null;
  let pdfSlots = [];
  let pdfZoomScale = 1.0;
  let currentPage = 1;
  let totalPages = 0;
  let isEditing = false;
  let isDirty = false;
  let savedRange = null;
  let searchMatches = [];
  let activeMatchIdx = 0;
  let isSpeaking = false;
  let wakeLock = null;
  let scrollProgressTimer = null;
  let pdfProgressTimer = null;
  let iframeLoadFallbackTimer = null;
  let openSequence = 0;
  let lastResizeAt = 0;

  function writeSessionSnapshot(bookId, scrollY, scrollRatio, page) {
    if (!bookId) return;
    try {
      localStorage.setItem(
        SESSION_STORAGE_KEY,
        JSON.stringify({
          bookId,
          scrollY: Math.max(0, Math.round(scrollY || 0)),
          scrollRatio: Math.max(0, Math.min(1, Number(scrollRatio || 0))),
          page: Math.max(1, Math.round(page || 1)),
          ts: Date.now()
        })
      );
    } catch (_) {}
  }

  function clearSessionSnapshot() {
    try {
      localStorage.removeItem(SESSION_STORAGE_KEY);
    } catch (_) {}
  }

  function setDockHidden(hidden) {
    const canHide =
      !isEditing &&
      aaBackdrop.classList.contains('hidden') &&
      searchOverlay.classList.contains('hidden') &&
      moreDropdown.classList.contains('hidden');
    dockHidden = canHide ? Boolean(hidden) : false;
    topBarRead?.classList.toggle('dock-hidden', dockHidden);
    overlay?.classList.toggle('reader-immersive', dockHidden);
    btnFloatingBack?.classList.toggle('visible', dockHidden && !isEditing);
  }

  function trimPdfCanvases(centerPage, keepRadius = 3) {
    for (const slot of pdfSlots) {
      if (slot.rendered && Math.abs(slot.pageNum - centerPage) > keepRadius) {
        const h = slot.card.offsetHeight;
        if (h > 0) slot.card.style.minHeight = `${h}px`;
        slot.canvas.width = 1;
        slot.canvas.height = 1;
        slot.canvas.style.width = '';
        slot.canvas.style.height = '';
        slot.rendered = false;
      }
    }
  }

  function cleanupStage() {
    clearTimeout(scrollProgressTimer);
    scrollProgressTimer = null;
    clearTimeout(pdfProgressTimer);
    pdfProgressTimer = null;
    clearTimeout(iframeLoadFallbackTimer);
    iframeLoadFallbackTimer = null;

    if (pdfObserver) {
      pdfObserver.disconnect();
      pdfObserver = null;
    }
    pdfSlots = [];
    if (pdfDoc) {
      pdfDoc.destroy().catch(() => {});
      pdfDoc = null;
    }
    if (activeBlobUrl) {
      URL.revokeObjectURL(activeBlobUrl);
      activeBlobUrl = null;
    }
    stageEl.innerHTML = '';
    activeIframe = null;
  }

  async function requestWakeLock() {
    const keepOn = Boolean(getSettings().keepScreenOn);
    try { window.AndroidBridge?.setKeepScreenOn?.(keepOn); } catch (_) {}
    if (!keepOn || !navigator.wakeLock) return;
    try {
      wakeLock = await navigator.wakeLock.request('screen');
    } catch (_) {}
  }

  function releaseWakeLock() {
    try { window.AndroidBridge?.setKeepScreenOn?.(false); } catch (_) {}
    if (wakeLock) {
      wakeLock.release().catch(() => {});
      wakeLock = null;
    }
  }

  function stopTts() {
    if ('speechSynthesis' in window) {
      try { window.speechSynthesis.cancel(); } catch (_) {}
    }
    isSpeaking = false;
    updateTtsButton();
  }

  function closeMoreMenu() {
    moreDropdown?.classList.add('hidden');
    btnMore?.classList.remove('active');
  }

  function updateReaderSubtitle() {
    if (!currentBook || !subEl) return;
    const pct = Math.max(0, Math.min(100, Math.round(currentBook.progressPercent || 0)));
    const titleTrim = (currentBook.title || '').trim().toLowerCase();
    const g = (currentBook.groupName || '').trim();
    const s = (currentBook.subtitle || '').trim();
    const prefix = (g && g.toLowerCase() !== titleTrim)
      ? g
      : (s && s.toLowerCase() !== titleTrim ? s : '');
    subEl.textContent = prefix ? `${pct}% · ${prefix}` : `${pct}% đã đọc`;
  }

  function updateBookmarkUi() {
    const marked = Boolean(currentBook?.isFavorite);
    btnBookmark.classList.toggle('is-bookmarked', marked);
    const svg = btnBookmark.querySelector('svg');
    if (svg) svg.setAttribute('fill', marked ? 'currentColor' : 'none');
    if (bookmarkLabel) bookmarkLabel.textContent = marked ? 'Đã đánh dấu trang' : 'Đánh dấu trang';
  }

  function updateTtsButton() {
    const s = getSettings();
    const show = Boolean(s.readAloudEnabled && currentBook && currentBook.kind !== 'pdf' && !isEditing);
    btnTts.classList.toggle('hidden', !show);
    btnTts.classList.toggle('active', isSpeaking);
    btnTts.setAttribute('aria-label', isSpeaking ? 'Dừng đọc to' : 'Đọc to tài liệu');
    if (ttsLabel) ttsLabel.textContent = isSpeaking ? 'Dừng đọc to' : 'Đọc to tài liệu';
  }

  function toggleTts() {
    if (!('speechSynthesis' in window)) {
      onShowToast('Trình duyệt không hỗ trợ đọc to');
      return;
    }
    if (isSpeaking) {
      stopTts();
      return;
    }
    const doc = activeIframe?.contentDocument;
    if (!doc) return;
    const activeDocx = doc.querySelector('.docx-page.active-page .docx-page-body');
    const target = (doc.body.getAttribute('data-page-mode') === 'paginated' && activeDocx) ? activeDocx : doc.body;
    const text = (target?.innerText || target?.textContent || '').replace(/\s+/g, ' ').trim();
    if (!text) {
      onShowToast('Trang hiện tại chưa có chữ để đọc');
      return;
    }
    window.speechSynthesis.cancel();
    const utter = new SpeechSynthesisUtterance(text.slice(0, 6000));
    utter.lang = 'vi-VN';
    const voices = window.speechSynthesis.getVoices();
    const viVoice = voices.find((v) => (v.lang || '').toLowerCase().startsWith('vi'));
    if (viVoice) utter.voice = viVoice;
    utter.onend = () => { isSpeaking = false; updateTtsButton(); };
    utter.onerror = () => { isSpeaking = false; updateTtsButton(); };
    isSpeaking = true;
    updateTtsButton();
    window.speechSynthesis.speak(utter);
  }

  function resolveEffectiveTheme(settings) {
    if (settings.themeMode === 'AUTO') {
      const appT = normalizeAppTheme(settings.appThemeMode);
      if (appT === 'dark') return 'DARK';
      if (appT === 'sepia') return 'SEPIA';
      return 'LIGHT';
    }
    return settings.themeMode;
  }

  function applyAppearance() {
    if (!currentBook) return;
    const s = getSettings();
    const effTheme = resolveEffectiveTheme(s);

    const bgMap = {
      ORIGINAL: '#FFFFFF',
      AUTO: '#F7F8FC',
      LIGHT: '#F7F8FC',
      SEPIA: '#F4ECD8',
      DARK: '#121622'
    };
    const readerBg = bgMap[effTheme] || '#F7F8FC';
    overlay.style.background = readerBg;
    overlay.dataset.readerTheme = effTheme;

    const themeMeta = document.querySelector('meta[name="theme-color"]');
    if (themeMeta) {
      themeMeta.setAttribute('content', readerBg);
    }
    try { window.AndroidBridge?.setSystemBarStyle?.(effTheme === 'DARK'); } catch (_) {}

    updateTtsButton();
    updatePageNavUi();

    if (currentBook.kind === 'pdf') {
      const filterMap = {
        DARK: 'invert(0.92) hue-rotate(180deg) contrast(0.95)',
        SEPIA: 'sepia(0.32) brightness(0.96) contrast(0.96)',
        LIGHT: 'none',
        ORIGINAL: 'none'
      };
      stageEl.style.filter = filterMap[effTheme] || 'none';
      return;
    }

    stageEl.style.filter = 'none';
    const doc = activeIframe?.contentDocument;
    if (!doc || !doc.head || !doc.body) return;

    // 1. Safe padding (stable whether dock is visible or auto-hidden) + responsive text size
    let padStyle = doc.getElementById('__yunzy_pad_style');
    if (!padStyle) {
      padStyle = doc.createElement('style');
      padStyle.id = '__yunzy_pad_style';
      doc.head.appendChild(padStyle);
    }
    const zoomScale = (s.fontSizePt / 17).toFixed(3);
    const safeTop = syncSafeAreaMetrics();
    const topPad = Math.max(84, Math.ceil(safeTop + 80));
    const botPad = isEditing ? 108 : 88;
    padStyle.textContent = `
      html, body { height: auto !important; min-height: 100% !important; overflow-y: auto !important; overflow-x: hidden !important; -webkit-overflow-scrolling: touch !important; }
      body { padding-top: ${topPad}px !important; padding-bottom: ${botPad}px !important; font-size: calc(16px * ${zoomScale}) !important; }
    `;

    // 2. Theme, font, spacing
    const isPureOriginal = effTheme === 'ORIGINAL' && s.fontFamily === 'ORIGINAL' && s.lineSpacing === 'NORMAL';
    let appStyle = doc.getElementById('__yunzy_appearance_css');
    if (isPureOriginal) {
      if (appStyle) appStyle.remove();
    } else {
      if (!appStyle) {
        appStyle = doc.createElement('style');
        appStyle.id = '__yunzy_appearance_css';
        doc.head.appendChild(appStyle);
      }
      const themeCss = effTheme === 'DARK'
        ? ':root { --doc-backdrop: #0B0F1A; --page-bg: #151B2B; --page-fg: #E2E8F0; --page-muted: #94A3B8; --page-border: rgba(148,163,184,0.22); } html, body { background: #121622 !important; color: #E2E8F0 !important; } body[data-doc-kind="docx"] { background: #0B0F1A !important; } h1,h2,h3,h4,p,li,td,th,summary { color: #E2E8F0 !important; } table, .card-box { border-color: rgba(148,163,184,0.25) !important; }'
        : effTheme === 'SEPIA'
        ? ':root { --doc-backdrop: #E7DCC4; --page-bg: #F6EFE0; --page-fg: #3B2E1E; --page-muted: #7C6A52; --page-border: rgba(124,106,82,0.20); } html, body { background: #F4ECD8 !important; color: #3B2E1E !important; } body[data-doc-kind="docx"] { background: #E7DCC4 !important; } h1,h2,h3,h4,p,li,td,th,summary { color: #3B2E1E !important; }'
        : effTheme === 'LIGHT'
        ? ':root { --doc-backdrop: #EAEEF7; --page-bg: #FFFFFF; --page-fg: #151C34; --page-muted: #6B748F; } html, body { background: #F7F8FC !important; color: #151C34 !important; } body[data-doc-kind="docx"] { background: #EAEEF7 !important; }'
        : '';
      const fontCss = FONT_CSS_MAP[s.fontFamily] ? `font-family: ${FONT_CSS_MAP[s.fontFamily]} !important;` : '';
      const lineCss = `line-height: ${LINE_SPACING_MAP[s.lineSpacing] || 1.65} !important;`;
      appStyle.textContent = `${themeCss}\nhtml, body { ${fontCss} ${lineCss} }\np, li, div, span { ${fontCss} ${lineCss} }`;
    }

    // 3. Paginated DOCX handling
    const docxPages = doc.querySelectorAll('.docx-page');
    if (docxPages.length > 0) {
      totalPages = docxPages.length;
      const usePag = s.pageMode === 'PAGINATED' && !isEditing;
      doc.body.setAttribute('data-page-mode', usePag ? 'paginated' : 'scroll');
      currentPage = Math.max(1, Math.min(currentPage, totalPages));
      docxPages.forEach((sec, idx) => {
        sec.classList.toggle('active-page', idx + 1 === currentPage);
        const ft = sec.querySelector('.docx-page-footer');
        if (ft) ft.textContent = `Trang ${idx + 1} / ${totalPages}`;
      });
      if (usePag) activeIframe.contentWindow?.scrollTo(0, 0);
    } else if (doc.body) {
      const winH = Math.max(activeIframe.clientHeight - 110, 260);
      const docH = Math.max(doc.body.scrollHeight, doc.documentElement.scrollHeight, 1);
      totalPages = Math.max(1, Math.ceil(docH / winH));
    }
    updatePageNavUi();
  }

  function updatePageNavUi() {
    if (!currentBook) return;
    const s = getSettings();
    const showNav = !isEditing && s.pageMode === 'PAGINATED' && totalPages > 0 && aaBackdrop.classList.contains('hidden') && searchOverlay.classList.contains('hidden');
    pageNavBar.classList.toggle('hidden', !showNav);
    pageNavText.textContent = `Trang ${currentPage} / ${totalPages}`;
    pageNavPrev.disabled = currentPage <= 1;
    pageNavNext.disabled = currentPage >= totalPages;

    const showEdges = showNav && s.tapEdgeToTurnPage && totalPages > 1;
    edgeLeft.classList.toggle('hidden', !showEdges);
    edgeRight.classList.toggle('hidden', !showEdges);
  }

  async function goToPage(target) {
    if (!currentBook || totalPages <= 0) return;
    const clamped = Math.max(1, Math.min(target, totalPages));
    currentPage = clamped;

    let scrollY = 0;
    if (currentBook.kind === 'pdf') {
      await renderPdfPaginatedPage(clamped);
    } else if (activeIframe?.contentDocument) {
      const doc = activeIframe.contentDocument;
      if (doc.querySelectorAll('.docx-page').length > 0) {
        applyAppearance();
      } else {
        const winH = Math.max(activeIframe.clientHeight - 110, 260);
        scrollY = Math.max(0, (clamped - 1) * winH);
        activeIframe.contentWindow?.scrollTo({ top: scrollY, behavior: 'smooth' });
      }
    }

    const ratio = totalPages <= 1 ? 1 : (clamped - 1) / (totalPages - 1);
    const pct = Math.max(0, Math.min(100, Math.round((clamped / totalPages) * 100)));
    writeSessionSnapshot(currentBook.id, scrollY, ratio, clamped);
    await persistProgress(pct, ratio, `Trang ${clamped} / ${totalPages}`, scrollY);
    updatePageNavUi();
  }

  async function persistProgress(pct, ratio, chapter, scrollY = null) {
    if (!currentBook) return;
    const bookId = currentBook.id;
    const nextScrollY = scrollY !== null ? Math.max(0, Math.round(scrollY)) : (currentBook.scrollY || 0);
    const patch = {
      progressPercent: pct,
      scrollRatio: ratio,
      scrollY: nextScrollY,
      currentChapterTitle: chapter || currentBook.currentChapterTitle,
      lastOpenedAt: Date.now()
    };
    currentBook = { ...currentBook, ...patch };
    updateReaderSubtitle();
    await onUpdateBookProgress(bookId, patch);
  }

  function flushCurrentProgressSync() {
    if (!currentBook) return;
    if (currentBook.kind === 'pdf') {
      const ratio = totalPages <= 1 ? 1 : (currentPage - 1) / (totalPages - 1);
      const pct = totalPages > 0
        ? Math.max(0, Math.min(100, Math.round((currentPage / totalPages) * 100)))
        : (currentBook.progressPercent || 0);
      const chap = totalPages > 0 ? `Trang ${currentPage} / ${totalPages}` : currentBook.currentChapterTitle;
      writeSessionSnapshot(currentBook.id, 0, ratio, currentPage);
      persistProgress(pct, ratio, chap, 0).catch(() => {});
      return;
    }

    const doc = activeIframe?.contentDocument;
    const win = activeIframe?.contentWindow;
    if (!doc || !win) return;
    const y = Math.max(0, win.scrollY || 0);
    const docH = Math.max(doc.body?.scrollHeight || 0, doc.documentElement?.scrollHeight || 0, 1);
    const maxScroll = Math.max(1, docH - activeIframe.clientHeight);
    const ratio = Math.max(0, Math.min(1, y / maxScroll));
    const pct = Math.max(0, Math.min(100, Math.round(ratio * 100)));
    const approxPage = totalPages > 0
      ? Math.max(1, Math.min(totalPages, Math.round(ratio * Math.max(0, totalPages - 1)) + 1))
      : currentPage;
    writeSessionSnapshot(currentBook.id, y, ratio, approxPage);
    persistProgress(pct, ratio, currentBook.currentChapterTitle, y).catch(() => {});
  }

  async function renderPdfBook(book, rawBlob, restoreHint = null) {
    cleanupStage();
    if (!rawBlob) {
      throw new Error('Không tìm thấy dữ liệu PDF.');
    }
    const pdfjs = await loadPdfJs();
    const dataCopy = rawBlob instanceof ArrayBuffer
      ? rawBlob.slice(0)
      : rawBlob instanceof Uint8Array
      ? rawBlob.slice().buffer
      : await rawBlob.arrayBuffer();

    pdfDoc = await pdfjs.getDocument({ data: new Uint8Array(dataCopy) }).promise;
    totalPages = pdfDoc.numPages || 1;
    const hintRatio = restoreHint?.scrollRatio ?? book.scrollRatio ?? 0;
    const restored = restoreHint?.page
      ? Number(restoreHint.page)
      : Math.round(hintRatio * Math.max(0, totalPages - 1)) + 1;
    currentPage = Math.max(1, Math.min(restored, totalPages));
    pdfZoomScale = 1.0;

    if (getSettings().pageMode === 'PAGINATED') {
      await renderPdfPaginatedPage(currentPage);
    } else {
      await renderPdfScrollMode();
    }
    updatePageNavUi();
  }

  async function renderPdfPageCanvas(pageNum, canvas, maxCssWidth) {
    if (!pdfDoc || !canvas) return;
    const page = await pdfDoc.getPage(pageNum);
    const dpr = Math.min(window.devicePixelRatio || 1, 2);
    const baseVp = page.getViewport({ scale: 1 });
    const cssScale = (maxCssWidth / baseVp.width) * pdfZoomScale;
    const vp = page.getViewport({ scale: cssScale * dpr });

    canvas.width = Math.max(1, Math.floor(vp.width));
    canvas.height = Math.max(1, Math.floor(vp.height));
    canvas.style.width = `${Math.floor(vp.width / dpr)}px`;
    canvas.style.height = `${Math.floor(vp.height / dpr)}px`;

    const ctx = canvas.getContext('2d');
    await page.render({ canvasContext: ctx, viewport: vp }).promise;
    try { page.cleanup(); } catch (_) {}
  }

  async function renderPdfPaginatedPage(pageNum) {
    if (!pdfDoc) return;
    if (pdfObserver) {
      pdfObserver.disconnect();
      pdfObserver = null;
    }
    pdfSlots = [];
    stageEl.innerHTML = '';
    const wrap = document.createElement('div');
    wrap.className = 'pdf-paginated-wrap';
    const canvas = document.createElement('canvas');
    canvas.className = 'pdf-page-canvas';
    wrap.appendChild(canvas);
    stageEl.appendChild(wrap);

    const maxW = Math.max(260, stageEl.clientWidth - 28);
    await renderPdfPageCanvas(pageNum, canvas, maxW);
  }

  async function renderPdfScrollMode() {
    if (!pdfDoc) return;
    if (pdfObserver) {
      pdfObserver.disconnect();
      pdfObserver = null;
    }
    stageEl.innerHTML = '';

    const wrap = document.createElement('div');
    wrap.className = 'pdf-scroll-wrap';
    const maxW = Math.max(260, Math.min(760, stageEl.clientWidth - 24));

    const placeholders = [];
    for (let i = 1; i <= totalPages; i++) {
      const card = document.createElement('div');
      card.className = 'pdf-page-card';
      card.dataset.page = String(i);
      const canvas = document.createElement('canvas');
      canvas.className = 'pdf-page-canvas';
      const label = document.createElement('div');
      label.className = 'pdf-page-tag';
      label.textContent = `Trang ${i} / ${totalPages}`;
      card.appendChild(canvas);
      card.appendChild(label);
      wrap.appendChild(card);
      placeholders.push({ pageNum: i, card, canvas, rendered: false });
    }
    pdfSlots = placeholders;
    stageEl.appendChild(wrap);

    const mountedAt = Date.now();
    let lastWrapY = 0;
    let tapStartX = 0;
    let tapStartY = 0;
    let tapStartTime = 0;

    wrap.addEventListener('pointerdown', (e) => {
      closeMoreMenu();
      tapStartX = e.clientX;
      tapStartY = e.clientY;
      tapStartTime = Date.now();
    }, { passive: true });

    wrap.addEventListener('click', (e) => {
      if (Date.now() - tapStartTime > 320) return;
      if (Math.hypot(e.clientX - tapStartX, e.clientY - tapStartY) >= 10) return;
      const target = e.target?.nodeType === 1 ? e.target : e.target?.parentElement;
      if (target?.closest('a, button, input, select, textarea, [role="button"]')) return;
      setDockHidden(!dockHidden);
    });

    wrap.addEventListener('scroll', () => {
      const y = Math.max(0, wrap.scrollTop || 0);
      if (Date.now() - mountedAt > 250 && Date.now() - lastResizeAt > 350) {
        if (y > 56 && y - lastWrapY > 18) setDockHidden(true);
        else if (lastWrapY - y > 14 || y <= 32) setDockHidden(false);
      }
      lastWrapY = y;
    }, { passive: true });

    pdfObserver = new IntersectionObserver((entries) => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        const pNum = parseInt(entry.target.dataset.page, 10);
        const slot = placeholders[pNum - 1];
        if (slot && !slot.rendered) {
          slot.rendered = true;
          renderPdfPageCanvas(pNum, slot.canvas, maxW)
            .then(() => {
              slot.card.style.minHeight = '';
            })
            .catch(() => {});
        }
        currentPage = pNum;
        trimPdfCanvases(pNum, 3);
        const ratio = totalPages <= 1 ? 1 : (pNum - 1) / (totalPages - 1);
        const pct = Math.round((pNum / totalPages) * 100);
        if (currentBook) {
          writeSessionSnapshot(currentBook.id, 0, ratio, pNum);
          clearTimeout(pdfProgressTimer);
          pdfProgressTimer = setTimeout(() => {
            persistProgress(pct, ratio, `Trang ${pNum} / ${totalPages}`, 0).catch(() => {});
          }, 240);
        }
      }
    }, { root: wrap, threshold: 0.25 });

    placeholders.forEach((s) => pdfObserver.observe(s.card));

    if (currentPage > 1 && placeholders[currentPage - 1]) {
      setTimeout(() => {
        placeholders[currentPage - 1]?.card.scrollIntoView({ behavior: 'instant', block: 'start' });
        lastWrapY = wrap.scrollTop || 0;
      }, 60);
    }
  }

  function renderHtmlOrDocxBook(book, htmlContent, restoreHint = null) {
    cleanupStage();
    const blob = new Blob(
      [htmlContent || '<!DOCTYPE html><html><body><p></p></body></html>'],
      { type: 'text/html;charset=utf-8' }
    );
    activeBlobUrl = URL.createObjectURL(blob);

    const iframe = document.createElement('iframe');
    iframe.className = 'reader-iframe is-restoring';
    iframe.setAttribute('sandbox', 'allow-same-origin allow-scripts');
    iframe.setAttribute('title', book.title);
    iframe.src = activeBlobUrl;
    stageEl.appendChild(iframe);
    activeIframe = iframe;

    iframeLoadFallbackTimer = setTimeout(() => {
      iframe.classList.remove('is-restoring');
    }, 400);

    iframe.addEventListener('load', () => {
      if (iframe !== activeIframe) return;
      const doc = iframe.contentDocument;
      const win = iframe.contentWindow;
      if (!doc || !win) return;

      const loadedAt = Date.now();
      let lastScrollY = 0;
      let restoreComplete = false;
      let tapStartX = 0;
      let tapStartY = 0;
      let tapStartTime = 0;

      doc.addEventListener('pointerdown', (e) => {
        closeMoreMenu();
        tapStartX = e.clientX;
        tapStartY = e.clientY;
        tapStartTime = Date.now();
      }, { passive: true });

      doc.addEventListener('click', (e) => {
        if (isEditing) return;
        if (!aaBackdrop.classList.contains('hidden') || !searchOverlay.classList.contains('hidden')) return;
        const selText = win.getSelection()?.toString().trim() || '';
        if (selText.length > 0) return;
        if (Date.now() - tapStartTime > 320) return;
        if (Math.hypot(e.clientX - tapStartX, e.clientY - tapStartY) >= 10) return;
        const target = e.target?.nodeType === 1 ? e.target : e.target?.parentElement;
        if (!target) return;
        if (target.closest('a, button, input, select, textarea, label, summary, details, [role="button"], [onclick], [contenteditable="true"]')) {
          return;
        }
        if (win.getComputedStyle(target).cursor === 'pointer') {
          return;
        }
        setDockHidden(!dockHidden);
      });

      const docxCount = doc.querySelectorAll('.docx-page').length;
      const hintRatio = Number(restoreHint?.scrollRatio ?? book.scrollRatio ?? 0);
      if (docxCount > 0) {
        totalPages = docxCount;
        const targetPage = restoreHint?.page
          ? Number(restoreHint.page)
          : Math.round(hintRatio * Math.max(0, docxCount - 1)) + 1;
        currentPage = Math.max(1, Math.min(targetPage, docxCount));
      }

      applyAppearance();

      const targetY = Number(restoreHint?.scrollY ?? book.scrollY ?? 0);
      const targetRatio = hintRatio;

      win.requestAnimationFrame(() => {
        win.requestAnimationFrame(() => {
          if (iframe !== activeIframe) return;
          if (getSettings().pageMode === 'SCROLL') {
            const maxScroll = Math.max(0, Math.max(doc.body.scrollHeight, doc.documentElement.scrollHeight) - iframe.clientHeight);
            if (targetY > 0 && maxScroll > 0) {
              win.scrollTo(0, Math.min(targetY, maxScroll));
            } else if (targetRatio > 0 && maxScroll > 0) {
              win.scrollTo(0, maxScroll * targetRatio);
            }
          }
          clearTimeout(iframeLoadFallbackTimer);
          iframeLoadFallbackTimer = null;
          iframe.classList.remove('is-restoring');
          lastScrollY = Math.max(0, win.scrollY || 0);
          restoreComplete = true;
        });
      });

      win.addEventListener('scroll', () => {
        if (!restoreComplete) return;
        if (getSettings().pageMode === 'PAGINATED' && doc.querySelectorAll('.docx-page').length > 0) return;
        const y = Math.max(0, win.scrollY || 0);
        if (Date.now() - loadedAt > 250 && Date.now() - lastResizeAt > 350 && !isEditing) {
          if (y > 56 && y - lastScrollY > 18) setDockHidden(true);
          else if (lastScrollY - y > 14 || y <= 32) setDockHidden(false);
        }
        lastScrollY = y;

        clearTimeout(scrollProgressTimer);
        scrollProgressTimer = setTimeout(() => {
          if (!currentBook || iframe !== activeIframe) return;
          const docH = Math.max(doc.body.scrollHeight, doc.documentElement.scrollHeight, 1);
          const maxScroll = Math.max(1, docH - iframe.clientHeight);
          const ratio = Math.max(0, Math.min(1, y / maxScroll));
          const pct = Math.max(0, Math.min(100, Math.round(ratio * 100)));
          const approxPage = totalPages > 0 ? Math.max(1, Math.min(totalPages, Math.round(ratio * Math.max(0, totalPages - 1)) + 1)) : 1;
          currentPage = approxPage;
          const chap = doc.querySelectorAll('.docx-page').length > 0 ? `Trang ${approxPage} / ${totalPages}` : currentBook.currentChapterTitle;
          writeSessionSnapshot(currentBook.id, y, ratio, approxPage);
          persistProgress(pct, ratio, chap, y);
        }, 240);
      }, { passive: true });
    });
  }

  function wireIframeEditing(enabled) {
    const doc = activeIframe?.contentDocument;
    const win = activeIframe?.contentWindow;
    if (!doc || !win) return;

    doc.body.setAttribute('data-editing', enabled ? 'true' : 'false');
    const pages = doc.querySelectorAll('.docx-page-body');
    if (pages.length > 0) {
      pages.forEach((p) => p.setAttribute('contenteditable', enabled ? 'true' : 'false'));
      if (enabled) {
        (doc.querySelector('.docx-page.active-page .docx-page-body') || pages[0])?.focus();
      }
    } else {
      doc.body.setAttribute('contenteditable', enabled ? 'true' : 'false');
      if (enabled) doc.body.focus();
    }

    if (enabled && !doc.__yunzyEditWired) {
      doc.__yunzyEditWired = true;
      const saveRangeAndNotify = () => {
        const sel = win.getSelection();
        if (sel && sel.rangeCount > 0) {
          savedRange = sel.getRangeAt(0).cloneRange();
        }
        syncRibbonActiveStates();
      };
      doc.addEventListener('input', () => {
        isDirty = true;
        updateEditStatus();
        saveRangeAndNotify();
      });
      doc.addEventListener('selectionchange', saveRangeAndNotify);
      doc.addEventListener('keyup', saveRangeAndNotify);
      doc.addEventListener('mouseup', saveRangeAndNotify);
      doc.addEventListener('touchend', saveRangeAndNotify);
    }

    if (!enabled) {
      isDirty = false;
      savedRange = null;
    }
    updateEditStatus();
    syncRibbonActiveStates();
  }

  function updateEditStatus() {
    editStatusEl.textContent = isDirty ? '• Chưa lưu thay đổi' : 'Chế độ chỉnh sửa';
    editStatusEl.style.color = isDirty ? 'var(--primary-blue)' : 'var(--text-secondary)';
  }

  function syncRibbonActiveStates() {
    const doc = activeIframe?.contentDocument;
    if (!doc || !isEditing) return;
    const cmdMap = {
      bold: doc.queryCommandState('bold'),
      italic: doc.queryCommandState('italic'),
      underline: doc.queryCommandState('underline'),
      strikeThrough: doc.queryCommandState('strikeThrough'),
      insertUnorderedList: doc.queryCommandState('insertUnorderedList'),
      insertOrderedList: doc.queryCommandState('insertOrderedList'),
      justifyLeft: doc.queryCommandState('justifyLeft'),
      justifyCenter: doc.queryCommandState('justifyCenter'),
      justifyRight: doc.queryCommandState('justifyRight')
    };
    editorRibbon.querySelectorAll('[data-cmd]').forEach((btn) => {
      btn.classList.toggle('active', Boolean(cmdMap[btn.dataset.cmd]));
    });
    const curBlock = (doc.queryCommandValue('formatBlock') || 'p').toLowerCase().replace(/[<>]/g, '');
    editorRibbon.querySelectorAll('[data-block]').forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.block === curBlock);
    });
    const sel = activeIframe.contentWindow.getSelection();
    let inMark = false;
    if (sel && sel.anchorNode) {
      const el = sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentElement;
      inMark = Boolean(el?.closest('mark'));
    }
    editorRibbon.querySelector('[data-action="highlight"]')?.classList.toggle('active-highlight', inMark);
  }

  function restoreRangeAndExec(cmd, val = null) {
    const doc = activeIframe?.contentDocument;
    const win = activeIframe?.contentWindow;
    if (!doc || !win) return;

    const sel = win.getSelection();
    if ((!sel || sel.rangeCount === 0) && savedRange) {
      try {
        const container = savedRange.commonAncestorContainer;
        const el = container.nodeType === 1 ? container : container.parentElement;
        (el?.closest('[contenteditable="true"]') || doc.body).focus();
        sel.removeAllRanges();
        sel.addRange(savedRange);
      } catch (_) {}
    }

    if (cmd === 'highlightToggle') {
      const anchor = sel?.anchorNode;
      const anchorEl = anchor ? (anchor.nodeType === 1 ? anchor : anchor.parentElement) : null;
      const markEl = anchorEl?.closest('mark');
      if (markEl) {
        const p = markEl.parentNode;
        while (markEl.firstChild) p.insertBefore(markEl.firstChild, markEl);
        p.removeChild(markEl);
      } else if (sel && sel.rangeCount > 0 && !sel.getRangeAt(0).collapsed) {
        const r = sel.getRangeAt(0);
        const mark = doc.createElement('mark');
        try {
          r.surroundContents(mark);
        } catch (_) {
          doc.execCommand('hiliteColor', false, '#FEF08A');
        }
      }
    } else {
      doc.execCommand(cmd, false, val);
    }

    isDirty = true;
    updateEditStatus();
    syncRibbonActiveStates();
  }

  function insertDocxPageBreak() {
    const doc = activeIframe?.contentDocument;
    const win = activeIframe?.contentWindow;
    if (!doc || !win) return;
    const wrap = doc.querySelector('.docx-pages');
    if (!wrap) return;

    const sel = win.getSelection();
    const anchor = sel?.anchorNode || savedRange?.startContainer;
    let curSec = anchor ? (anchor.nodeType === 1 ? anchor : anchor.parentElement)?.closest('.docx-page') : null;
    if (!curSec) curSec = doc.querySelector('.docx-page.active-page') || wrap.lastElementChild;

    const newSec = doc.createElement('section');
    newSec.className = 'docx-page';
    newSec.innerHTML = '<div class="docx-page-body" contenteditable="true"><p><br/></p></div><div class="docx-page-footer"></div>';
    if (curSec && curSec.nextSibling) wrap.insertBefore(newSec, curSec.nextSibling);
    else wrap.appendChild(newSec);

    const all = wrap.querySelectorAll('.docx-page');
    totalPages = all.length;
    all.forEach((s, idx) => {
      s.setAttribute('data-page', String(idx + 1));
      s.classList.remove('active-page');
      if (s === newSec) currentPage = idx + 1;
      const ft = s.querySelector('.docx-page-footer');
      if (ft) ft.textContent = `Trang ${idx + 1} / ${totalPages}`;
    });
    newSec.classList.add('active-page');
    const nb = newSec.querySelector('.docx-page-body');
    if (nb) {
      nb.focus();
      const range = doc.createRange();
      range.selectNodeContents(nb.firstChild || nb);
      range.collapse(true);
      const s = win.getSelection();
      s.removeAllRanges();
      s.addRange(range);
      savedRange = range.cloneRange();
    }
    newSec.scrollIntoView({ behavior: 'smooth', block: 'start' });
    isDirty = true;
    updateEditStatus();
  }

  function extractEditedDocxPages() {
    const doc = activeIframe?.contentDocument;
    if (!doc) return { pages: [], bodyHtml: '' };

    function readRuns(node, fmt, out) {
      if (node.nodeType === 3) {
        const txt = node.nodeValue || '';
        if (txt.length > 0) out.push({ text: txt, bold: fmt.b, italic: fmt.i, underline: fmt.u, strike: fmt.s, highlight: fmt.h });
        return;
      }
      if (node.nodeType !== 1) return;
      const t = node.tagName.toUpperCase();
      if (t === 'BR') { out.push({ text: '\n', ...fmt }); return; }
      const nf = { ...fmt };
      if (t === 'B' || t === 'STRONG') nf.b = true;
      if (t === 'I' || t === 'EM') nf.i = true;
      if (t === 'U') nf.u = true;
      if (t === 'S' || t === 'STRIKE' || t === 'DEL') nf.s = true;
      if (t === 'MARK' || (node.style?.backgroundColor && node.style.backgroundColor !== 'transparent')) nf.h = true;
      if (node.style) {
        if (node.style.fontWeight === 'bold' || parseInt(node.style.fontWeight, 10) >= 600) nf.b = true;
        if (node.style.fontStyle === 'italic') nf.i = true;
      }
      for (const ch of node.childNodes) readRuns(ch, nf, out);
    }

    function getAlign(el) {
      const a = (el.style?.textAlign || '').toLowerCase();
      return (a === 'center' || a === 'right' || a === 'justify') ? a : '';
    }

    function parseContainer(container, blocks) {
      for (const ch of container.childNodes) {
        if (ch.nodeType === 3) {
          const txt = (ch.nodeValue || '').trim();
          if (txt) blocks.push({ tag: 'p', align: '', runs: [{ text: txt }] });
          continue;
        }
        if (ch.nodeType !== 1) continue;
        const tag = ch.tagName.toUpperCase();
        if (tag === 'TABLE') {
          const rows = [...ch.querySelectorAll('tr')].map((tr) =>
            [...tr.querySelectorAll('td,th')].map((c) => (c.innerText || '').trim())
          ).filter((r) => r.length > 0);
          if (rows.length) blocks.push({ tag: 'table', rows });
        } else if (tag === 'UL' || tag === 'OL') {
          const liTag = tag === 'UL' ? 'ul_li' : 'ol_li';
          for (const li of ch.querySelectorAll(':scope > li')) {
            const runs = [];
            readRuns(li, { b: false, i: false, u: false, s: false, h: false }, runs);
            blocks.push({ tag: liTag, align: getAlign(li), runs });
          }
        } else if (tag === 'IMG') {
          const src = ch.getAttribute('src') || '';
          if (src) blocks.push({ tag: 'img', align: '', src });
        } else {
          const img = ch.querySelector('img');
          if (img && (ch.innerText || '').trim() === '') {
            blocks.push({ tag: 'img', align: getAlign(ch), src: img.getAttribute('src') || '' });
          } else {
            const runs = [];
            readRuns(ch, { b: false, i: false, u: false, s: false, h: false }, runs);
            const bTag = tag === 'H1' ? 'h1' : tag === 'H2' ? 'h2' : tag === 'H3' ? 'h3' : 'p';
            blocks.push({ tag: bTag, align: getAlign(ch), runs });
          }
        }
      }
    }

    const secEls = doc.querySelectorAll('.docx-page-body');
    const pages = [];
    if (secEls.length > 0) {
      secEls.forEach((sec) => {
        const blocks = [];
        parseContainer(sec, blocks);
        pages.push({ blocks });
      });
    } else {
      const blocks = [];
      parseContainer(doc.body, blocks);
      if (blocks.length) pages.push({ blocks });
    }
    return { pages, bodyHtml: doc.body.innerHTML };
  }

  async function saveDirect() {
    if (!currentBook || !activeIframe?.contentDocument) return;
    const { pages, bodyHtml } = extractEditedDocxPages();
    if (currentBook.kind === 'docx') {
      const updatedHtml = renderDocxPagesToHtml(currentBook.title, pages);
      const metaPatch = {
        subtitle: `DOCX · ${Math.max(1, pages.length)} trang`,
        fileSizeBytes: updatedHtml.length
      };
      await onSaveBookContent(currentBook.id, { htmlContent: updatedHtml, docxPages: pages }, metaPatch);
      currentBook = { ...currentBook, ...metaPatch };
    } else {
      const content = await getBookContent(currentBook.id);
      const baseHtml = content?.htmlContent || '';
      const updatedHtml = /<body[^>]*>[\s\S]*<\/body>/i.test(baseHtml)
        ? baseHtml.replace(/(<body[^>]*>)[\s\S]*(<\/body>)/i, (_, open, close) => `${open}\n${bodyHtml}\n${close}`)
        : `<!DOCTYPE html><html lang="vi"><head><meta charset="utf-8"/><title>${escapeAttr(currentBook.title)}</title></head><body>\n${bodyHtml}\n</body></html>`;
      const metaPatch = { fileSizeBytes: updatedHtml.length };
      await onSaveBookContent(currentBook.id, { htmlContent: updatedHtml }, metaPatch);
      currentBook = { ...currentBook, ...metaPatch };
    }
    isDirty = false;
    updateEditStatus();
    onShowToast('Đã lưu thay đổi trực tiếp');
  }

  async function saveCopyToLibrary() {
    if (!currentBook || !activeIframe?.contentDocument) return;
    const copyTitle = window.prompt('Nhập tên bản sao vào thư viện:', `${currentBook.title} (Bản sao)`);
    if (!copyTitle || !copyTitle.trim()) return;
    const { pages, bodyHtml } = extractEditedDocxPages();
    await onCreateCopyBook(currentBook, copyTitle.trim(), pages, bodyHtml);
    isDirty = false;
    updateEditStatus();
  }

  async function exportDocxCopy() {
    if (!currentBook || !activeIframe?.contentDocument) return;
    const { pages } = extractEditedDocxPages();
    const bytes = await buildDocxBytesFromPages(pages);
    const safeName = (currentBook.title || 'YunzyBook').replace(/[^\w\s\p{L}.-]/gu, '_').trim();
    downloadBlob(
      bytes,
      `${safeName}_ban_sao.docx`,
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    );
    onShowToast('Đã xuất bản sao .docx');
  }

  function enterEditMode() {
    if (!currentBook || currentBook.kind === 'pdf') return;
    isEditing = true;
    setDockHidden(false);
    topBarRead.classList.add('hidden');
    topBarEdit.classList.remove('hidden');
    editorRibbon.classList.remove('hidden');
    editTitleEl.textContent = currentBook.title;
    applyAppearance();
    wireIframeEditing(true);
  }

  function exitEditMode() {
    if (isDirty && !window.confirm('Tài liệu có nội dung chưa lưu. Bạn chắc chắn muốn thoát chỉnh sửa?')) {
      return;
    }
    isEditing = false;
    topBarEdit.classList.add('hidden');
    editorRibbon.classList.add('hidden');
    topBarRead.classList.remove('hidden');
    editMenuEl.classList.add('hidden');
    setDockHidden(false);
    wireIframeEditing(false);
    applyAppearance();
  }

  function runSearch(query) {
    const clean = (query || '').trim().toLowerCase();
    searchMatches = [];
    activeMatchIdx = 0;
    searchList.innerHTML = '';

    if (!clean) {
      searchCountEl.classList.add('hidden');
      searchList.innerHTML = '<div class="empty-hint">Nhập từ khóa để tìm nhanh trong tài liệu</div>';
      return;
    }

    const doc = activeIframe?.contentDocument;
    if (!doc) return;
    const blocks = doc.querySelectorAll('p, li, td, h1, h2, h3');
    const docH = Math.max(doc.body.scrollHeight, doc.documentElement.scrollHeight, 1);

    for (const el of blocks) {
      const txt = (el.innerText || '').trim();
      const idx = txt.toLowerCase().indexOf(clean);
      if (idx >= 0 && txt.length > 0) {
        const sec = el.closest?.('.docx-page');
        const pageNum = sec ? (parseInt(sec.getAttribute('data-page') || '0', 10) || 0) : 0;
        const absTop = el.getBoundingClientRect().top + activeIframe.contentWindow.scrollY;
        const ratio = Math.max(0, Math.min(1, absTop / docH));
        searchMatches.push({
          snippet: txt.length > 130 ? `${txt.slice(0, 130)}…` : txt,
          pageNumber: pageNum,
          ratio
        });
        if (searchMatches.length >= 30) break;
      }
    }

    searchCountEl.textContent = `${searchMatches.length}`;
    searchCountEl.classList.toggle('hidden', searchMatches.length === 0);

    if (searchMatches.length === 0) {
      searchList.innerHTML = '<div class="empty-hint">Không tìm thấy kết quả phù hợp</div>';
      return;
    }

    searchMatches.forEach((m, idx) => {
      const item = document.createElement('button');
      item.type = 'button';
      item.className = 'search-match-card';
      item.innerHTML = `<div class="match-meta">${m.pageNumber > 0 ? `Trang ${m.pageNumber}` : `Vị trí ${idx + 1}`}</div><div class="match-text">${escapeAttr(m.snippet)}</div>`;
      item.addEventListener('click', () => {
        if (m.pageNumber > 0) {
          currentPage = m.pageNumber;
          applyAppearance();
        }
        const maxS = Math.max(0, Math.max(doc.body.scrollHeight, doc.documentElement.scrollHeight) - activeIframe.clientHeight);
        activeIframe.contentWindow?.scrollTo({ top: maxS * m.ratio, behavior: 'smooth' });
        searchOverlay.classList.add('hidden');
        updatePageNavUi();
      });
      searchList.appendChild(item);
    });
  }

  function syncAaSheetControls() {
    const s = getSettings();
    aaSheet.querySelectorAll('[data-theme]').forEach((b) => b.classList.toggle('selected', b.dataset.theme === s.themeMode));
    aaSheet.querySelectorAll('[data-font]').forEach((b) => b.classList.toggle('selected', b.dataset.font === s.fontFamily));
    aaSheet.querySelectorAll('[data-spacing]').forEach((b) => b.classList.toggle('selected', b.dataset.spacing === s.lineSpacing));
    aaSheet.querySelectorAll('[data-pagemode]').forEach((b) => b.classList.toggle('selected', b.dataset.pagemode === s.pageMode));
    const sizeVal = document.getElementById('aa-font-size-val');
    if (sizeVal) sizeVal.textContent = `${s.fontSizePt}pt`;
  }

  // Lifecycle listeners for iOS PWA suspend / resume & Safe Area updates
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden') {
      if (currentBook) {
        flushCurrentProgressSync();
        if (currentBook.kind === 'pdf') {
          trimPdfCanvases(currentPage, 1);
        }
      }
    } else if (document.visibilityState === 'visible' && currentBook) {
      requestWakeLock();
    }
  });
  window.addEventListener('pagehide', () => {
    if (currentBook) flushCurrentProgressSync();
  });
  window.addEventListener('pageshow', (e) => {
    if (e.persisted && currentBook) requestWakeLock();
  });
  window.addEventListener('resize', () => {
    lastResizeAt = Date.now();
    syncSafeAreaMetrics();
    if (currentBook) applyAppearance();
  });
  window.addEventListener('orientationchange', () => {
    lastResizeAt = Date.now();
    setTimeout(() => {
      syncSafeAreaMetrics();
      if (currentBook) applyAppearance();
    }, 120);
  });

  // Event bindings
  btnBack.addEventListener('click', close);
  btnFloatingBack?.addEventListener('click', close);
  btnMore?.addEventListener('click', (e) => {
    e.stopPropagation();
    const willOpen = moreDropdown.classList.contains('hidden');
    if (willOpen) {
      aaBackdrop.classList.add('hidden');
      btnAa.classList.remove('active');
      searchOverlay.classList.add('hidden');
      moreDropdown.classList.remove('hidden');
      btnMore.classList.add('active');
      setDockHidden(false);
    } else {
      closeMoreMenu();
    }
  });
  document.addEventListener('pointerdown', (e) => {
    if (!moreDropdown || moreDropdown.classList.contains('hidden')) return;
    if (!e.target.closest('#reader-more-dropdown') && !e.target.closest('#reader-btn-more')) {
      closeMoreMenu();
    }
  });

  btnTts.addEventListener('click', () => {
    closeMoreMenu();
    toggleTts();
  });
  btnEdit.addEventListener('click', () => {
    closeMoreMenu();
    enterEditMode();
  });
  btnAa.addEventListener('click', () => {
    closeMoreMenu();
    syncAaSheetControls();
    aaBackdrop.classList.toggle('hidden');
    btnAa.classList.toggle('active', !aaBackdrop.classList.contains('hidden'));
    setDockHidden(false);
    updatePageNavUi();
  });
  aaBackdrop.addEventListener('click', (e) => {
    if (e.target === aaBackdrop || e.target.closest('#aa-btn-close')) {
      aaBackdrop.classList.add('hidden');
      btnAa.classList.remove('active');
      updatePageNavUi();
    }
  });
  btnSearch.addEventListener('click', () => {
    closeMoreMenu();
    aaBackdrop.classList.add('hidden');
    btnAa.classList.remove('active');
    searchOverlay.classList.toggle('hidden');
    setDockHidden(false);
    if (!searchOverlay.classList.contains('hidden')) {
      runSearch(searchInput.value);
      setTimeout(() => searchInput.focus(), 50);
    }
    updatePageNavUi();
  });
  document.getElementById('reader-search-close').addEventListener('click', () => {
    searchOverlay.classList.add('hidden');
    updatePageNavUi();
  });
  searchInput.addEventListener('input', () => runSearch(searchInput.value));

  btnBookmark.addEventListener('click', async () => {
    if (!currentBook) return;
    const nextFav = !currentBook.isFavorite;
    currentBook = { ...currentBook, isFavorite: nextFav };
    updateBookmarkUi();
    closeMoreMenu();
    await onUpdateBookMeta(currentBook.id, { isFavorite: nextFav });
  });

  pageNavPrev.addEventListener('click', () => goToPage(currentPage - 1));
  pageNavNext.addEventListener('click', () => goToPage(currentPage + 1));
  edgeLeft.addEventListener('click', () => goToPage(currentPage - 1));
  edgeRight.addEventListener('click', () => goToPage(currentPage + 1));

  // Aa Sheet controls
  aaSheet.addEventListener('click', async (e) => {
    const themeBtn = e.target.closest('[data-theme]');
    const fontBtn = e.target.closest('[data-font]');
    const spBtn = e.target.closest('[data-spacing]');
    const pmBtn = e.target.closest('[data-pagemode]');
    const cur = getSettings();

    if (themeBtn) {
      await onUpdateSettings({ ...cur, themeMode: themeBtn.dataset.theme });
    } else if (fontBtn) {
      await onUpdateSettings({ ...cur, fontFamily: fontBtn.dataset.font });
    } else if (spBtn) {
      await onUpdateSettings({ ...cur, lineSpacing: spBtn.dataset.spacing });
    } else if (pmBtn) {
      const nextMode = pmBtn.dataset.pagemode;
      await onUpdateSettings({ ...cur, pageMode: nextMode });
      if (currentBook?.kind === 'pdf') {
        if (nextMode === 'PAGINATED') await renderPdfPaginatedPage(currentPage);
        else await renderPdfScrollMode();
      }
    } else if (e.target.closest('#aa-font-minus')) {
      await onUpdateSettings({ ...cur, fontSizePt: Math.max(12, cur.fontSizePt - 1) });
    } else if (e.target.closest('#aa-font-plus')) {
      await onUpdateSettings({ ...cur, fontSizePt: Math.min(28, cur.fontSizePt + 1) });
    } else if (e.target.closest('#aa-btn-reset')) {
      await onUpdateSettings({
        ...cur,
        themeMode: 'ORIGINAL',
        fontSizePt: 17,
        fontFamily: 'ORIGINAL',
        lineSpacing: 'NORMAL',
        pageMode: 'SCROLL'
      });
      if (currentBook?.kind === 'pdf') await renderPdfScrollMode();
    } else {
      return;
    }
    syncAaSheetControls();
    applyAppearance();
  });

  // Editor Top bar & Formatting Ribbon
  document.getElementById('editor-btn-close').addEventListener('click', exitEditMode);
  document.getElementById('editor-btn-undo').addEventListener('click', () => restoreRangeAndExec('undo'));
  document.getElementById('editor-btn-redo').addEventListener('click', () => restoreRangeAndExec('redo'));
  document.getElementById('editor-btn-save').addEventListener('click', saveDirect);
  document.getElementById('editor-btn-more').addEventListener('click', () => {
    editMenuEl.classList.toggle('hidden');
  });
  document.getElementById('editor-opt-save-direct').addEventListener('click', () => {
    editMenuEl.classList.add('hidden');
    saveDirect();
  });
  document.getElementById('editor-opt-save-copy').addEventListener('click', () => {
    editMenuEl.classList.add('hidden');
    saveCopyToLibrary();
  });
  document.getElementById('editor-opt-export-docx').addEventListener('click', () => {
    editMenuEl.classList.add('hidden');
    exportDocxCopy();
  });

  editorRibbon.addEventListener('pointerdown', (e) => {
    // Keep WebView/iframe text selection focused on iOS Safari when tapping formatting pills
    if (e.target.closest('button')) e.preventDefault();
  });
  editorRibbon.addEventListener('click', (e) => {
    const cmdBtn = e.target.closest('[data-cmd]');
    const blockBtn = e.target.closest('[data-block]');
    const actionBtn = e.target.closest('[data-action]');
    if (cmdBtn) {
      restoreRangeAndExec(cmdBtn.dataset.cmd);
    } else if (blockBtn) {
      restoreRangeAndExec('formatBlock', `<${blockBtn.dataset.block}>`);
    } else if (actionBtn?.dataset.action === 'highlight') {
      restoreRangeAndExec('highlightToggle');
    } else if (actionBtn?.dataset.action === 'pagebreak') {
      insertDocxPageBreak();
    } else if (actionBtn?.dataset.action === 'blur') {
      activeIframe?.contentDocument?.activeElement?.blur();
    }
  });

  async function open(book, restoreHint = null) {
    const seq = ++openSequence;
    cleanupStage();

    currentBook = { ...book };
    isEditing = false;
    isDirty = false;
    currentPage = Math.max(1, Math.round(restoreHint?.page || 1));
    totalPages = 0;

    syncSafeAreaMetrics();
    titleEl.textContent = book.title;
    updateReaderSubtitle();
    updateBookmarkUi();
    btnEdit.classList.toggle('hidden', book.kind === 'pdf');
    btnSearch.classList.toggle('hidden', book.kind === 'pdf');
    topBarRead.classList.remove('hidden');
    topBarEdit.classList.add('hidden');
    editorRibbon.classList.add('hidden');
    aaBackdrop.classList.add('hidden');
    btnAa.classList.remove('active');
    searchOverlay.classList.add('hidden');
    closeMoreMenu();
    setDockHidden(false);

    document.body.classList.add('reader-open');
    overlay.classList.remove('hidden');
    applyAppearance();
    requestWakeLock();

    writeSessionSnapshot(
      book.id,
      restoreHint?.scrollY ?? book.scrollY ?? 0,
      restoreHint?.scrollRatio ?? book.scrollRatio ?? 0,
      currentPage
    );

    const content = await getBookContent(book.id);
    if (seq !== openSequence || !currentBook) return;

    if (book.kind === 'pdf') {
      await renderPdfBook(book, content?.rawBlob || book.rawBlob || null, restoreHint);
    } else {
      const html = content?.htmlContent ?? book.htmlContent ?? '';
      renderHtmlOrDocxBook(book, html, restoreHint);
    }
    if (seq === openSequence && currentBook) {
      applyAppearance();
    }
  }

  function close() {
    if (isEditing && isDirty) {
      if (!window.confirm('Bạn có thay đổi chưa lưu. Thoát trình đọc và hủy thay đổi?')) return;
    }
    openSequence++;
    flushCurrentProgressSync();
    clearSessionSnapshot();
    stopTts();
    releaseWakeLock();
    cleanupStage();
    currentBook = null;
    isEditing = false;
    setDockHidden(false);
    delete overlay.dataset.readerTheme;
    document.body.classList.remove('reader-open');
    overlay.classList.add('hidden');
    onCloseReader();
  }

  function handleHardwareBack() {
    if (!currentBook || overlay.classList.contains('hidden')) return false;
    if (!moreDropdown.classList.contains('hidden')) {
      closeMoreMenu();
      return true;
    }
    if (!editMenuEl.classList.contains('hidden')) {
      editMenuEl.classList.add('hidden');
      return true;
    }
    if (!searchOverlay.classList.contains('hidden')) {
      searchOverlay.classList.add('hidden');
      updatePageNavUi();
      return true;
    }
    if (!aaBackdrop.classList.contains('hidden')) {
      aaBackdrop.classList.add('hidden');
      btnAa.classList.remove('active');
      updatePageNavUi();
      return true;
    }
    if (isEditing) {
      exitEditMode();
      return true;
    }
    close();
    return true;
  }

  return {
    open,
    close,
    applyAppearance,
    handleHardwareBack
  };
}
