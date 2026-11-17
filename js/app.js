import { parseDocxBytes, renderDocxPagesToHtml, buildDocxBytesFromPages } from './docx.js';
import { createReaderController } from './reader.js';
import {
  DEFAULT_SETTINGS,
  computeStorageStats,
  deleteBook,
  loadAllBooks,
  loadGroups,
  loadSettings,
  saveGroups,
  saveSettings,
  upsertBook
} from './storage.js';
import {
  COVER_PALETTES,
  GROUP_ACCENTS,
  downloadBlob,
  escapeAttr,
  extractFirstHeading,
  extractHtmlTitle,
  formatBytesMb,
  hashColorIndex,
  isIosSafari,
  normalizeAppTheme,
  renderCoverMarkup
} from './ui-helpers.js';
import { unzipBytes, zipEntries } from './zip.js';

const TAB_ORDER = ['library', 'groups', 'settings'];
const APP_THEMES = ['Tự động', 'Sáng', 'Sepia', 'Tối'];
const FONT_FAMILIES = [
  { key: 'ORIGINAL', label: 'Mặc định' },
  { key: 'MODERN', label: 'Hiện đại' },
  { key: 'CLASSIC', label: 'Cổ điển' },
  { key: 'SYSTEM', label: 'Hệ thống' }
];
const LINE_SPACING_LABELS = {
  COMPACT: 'Gọn',
  NORMAL: 'Chuẩn',
  RELAXED: 'Thoáng'
};

let books = [];
let groups = [];
let settings = { ...DEFAULT_SETTINGS };
let activeTab = 'library';
let searchQuery = '';
let selectedFilter = 'ALL';
let selectedGroup = null;
let toastTimer = null;
let pendingCoverBookId = null;
let lastLongPressAt = 0;

function showToast(msg) {
  const el = document.getElementById('toast-pill');
  if (!el) return;
  el.textContent = msg;
  el.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.add('hidden'), 2600);
}

function applyGlobalVisualTheme() {
  const mode = normalizeAppTheme(settings.appThemeMode);
  document.documentElement.dataset.appTheme = mode;
  const op = Math.max(0.25, Math.min(1.0, Number(settings.glassOpacity ?? 0.78)));
  document.documentElement.style.setProperty('--glass-opacity', op.toFixed(2));
  const themeMeta = document.querySelector('meta[name="theme-color"]');
  if (themeMeta) {
    themeMeta.setAttribute('content', mode === 'dark' ? '#10141F' : mode === 'sepia' ? '#F2EBE0' : '#F4F6FB');
  }
}

async function updateSettings(next) {
  settings = await saveSettings(next);
  applyGlobalVisualTheme();
  renderSettingsTab();
}

// Reusable in-app modal helpers (replaces browser prompt/confirm)
function openInputModal({ title, value = '', placeholder = '' }) {
  return new Promise((resolve) => {
    const modal = document.getElementById('input-dialog-modal');
    const titleEl = document.getElementById('input-dialog-title');
    const field = document.getElementById('input-dialog-field');
    const btnOk = document.getElementById('input-dialog-ok');
    const btnCancel = document.getElementById('input-dialog-cancel');

    titleEl.textContent = title;
    field.value = value;
    field.placeholder = placeholder;
    modal.classList.remove('hidden');
    setTimeout(() => { field.focus(); field.select(); }, 50);

    const cleanup = (res) => {
      modal.classList.add('hidden');
      btnOk.removeEventListener('click', onOk);
      btnCancel.removeEventListener('click', onCancel);
      modal.removeEventListener('click', onBackdrop);
      field.removeEventListener('keydown', onKey);
      resolve(res);
    };
    const onOk = () => cleanup(field.value);
    const onCancel = () => cleanup(null);
    const onBackdrop = (e) => { if (e.target === modal) cleanup(null); };
    const onKey = (e) => {
      if (e.key === 'Enter') onOk();
      else if (e.key === 'Escape') onCancel();
    };
    btnOk.addEventListener('click', onOk);
    btnCancel.addEventListener('click', onCancel);
    modal.addEventListener('click', onBackdrop);
    field.addEventListener('keydown', onKey);
  });
}

function openConfirmModal({ title, message, okLabel = 'Xóa', danger = true }) {
  return new Promise((resolve) => {
    const modal = document.getElementById('confirm-dialog-modal');
    const titleEl = document.getElementById('confirm-dialog-title');
    const msgEl = document.getElementById('confirm-dialog-msg');
    const btnOk = document.getElementById('confirm-dialog-ok');
    const btnCancel = document.getElementById('confirm-dialog-cancel');

    titleEl.textContent = title;
    msgEl.textContent = message;
    btnOk.textContent = okLabel;
    btnOk.style.background = danger ? 'var(--danger-red)' : 'var(--primary-blue)';
    modal.classList.remove('hidden');

    const cleanup = (res) => {
      modal.classList.add('hidden');
      btnOk.removeEventListener('click', onOk);
      btnCancel.removeEventListener('click', onCancel);
      modal.removeEventListener('click', onBackdrop);
      resolve(res);
    };
    const onOk = () => cleanup(true);
    const onCancel = () => cleanup(false);
    const onBackdrop = (e) => { if (e.target === modal) cleanup(false); };
    btnOk.addEventListener('click', onOk);
    btnCancel.addEventListener('click', onCancel);
    modal.addEventListener('click', onBackdrop);
  });
}

function openGroupPickerModal(book) {
  const modal = document.getElementById('group-picker-modal');
  const titleEl = document.getElementById('group-picker-title');
  const listEl = document.getElementById('group-picker-list');
  const newInput = document.getElementById('group-picker-new-input');
  const newBtn = document.getElementById('group-picker-new-add');
  const closeBtn = document.getElementById('group-picker-close');

  titleEl.textContent = `Nhóm cho "${book.title}"`;
  newInput.value = '';

  const renderGroupChoices = () => {
    const choices = [
      { key: '', label: 'Không thuộc nhóm nào' },
      ...groups.map((g) => ({ key: g, label: g }))
    ];
    listEl.innerHTML = choices.map((c) => {
      const sel = (book.groupName || '').trim().toLowerCase() === c.key.toLowerCase();
      return `<button type="button" class="group-pick-item bounceable ${sel ? 'selected' : ''}" data-pick-group="${escapeAttr(c.key)}">
        <span>${escapeAttr(c.label)}</span>
        ${sel ? '<span>✓</span>' : ''}
      </button>`;
    }).join('');
  };
  renderGroupChoices();
  modal.classList.remove('hidden');

  const closeModal = () => {
    modal.classList.add('hidden');
    listEl.removeEventListener('click', onPick);
    newBtn.removeEventListener('click', onAddNew);
    closeBtn.removeEventListener('click', closeModal);
    modal.removeEventListener('click', onBackdrop);
  };
  const applyGroup = async (groupName) => {
    await upsertBook({ ...book, groupName });
    books = await loadAllBooks();
    groups = await loadGroups();
    renderAllTabs();
    closeModal();
    showToast(groupName ? `Đã chuyển vào nhóm "${groupName}"` : 'Đã bỏ khỏi nhóm');
  };
  const onPick = (e) => {
    const btn = e.target.closest('[data-pick-group]');
    if (btn) applyGroup(btn.dataset.pickGroup);
  };
  const onAddNew = async () => {
    const clean = newInput.value.trim();
    if (!clean) return;
    if (!groups.some((g) => g.toLowerCase() === clean.toLowerCase())) {
      groups.push(clean);
      await saveGroups(groups);
    }
    await applyGroup(clean);
  };
  const onBackdrop = (e) => { if (e.target === modal) closeModal(); };

  listEl.addEventListener('click', onPick);
  newBtn.addEventListener('click', onAddNew);
  closeBtn.addEventListener('click', closeModal);
  modal.addEventListener('click', onBackdrop);
}

function openCoverEditorModal(book) {
  const modal = document.getElementById('cover-editor-modal');
  const previewEl = document.getElementById('cover-ed-preview');
  const nameEl = document.getElementById('cover-ed-book-name');
  const gridEl = document.getElementById('cover-palette-grid');
  const pickImgBtn = document.getElementById('cover-ed-pick-img');
  const saveBtn = document.getElementById('cover-ed-save');
  const cancelBtn = document.getElementById('cover-ed-cancel');
  const closeBtn = document.getElementById('cover-ed-close');

  let selectedIdx = Number(book.coverColorIndex || 0);
  let useCustom = Boolean(book.customCoverDataUrl);

  nameEl.textContent = book.title;
  const renderPreviewAndGrid = () => {
    if (useCustom && book.customCoverDataUrl) {
      previewEl.innerHTML = renderCoverMarkup(book, '');
    } else {
      previewEl.innerHTML = renderCoverMarkup(book, '', selectedIdx);
    }
    gridEl.innerHTML = COVER_PALETTES.map((pal, idx) => `
      <button type="button" class="cover-style-tile bounceable ${!useCustom && selectedIdx === idx ? 'selected' : ''}" data-cstyle="${idx}" aria-label="${escapeAttr(pal.name)}">
        ${renderCoverMarkup({ ...book, customCoverDataUrl: null, coverColorIndex: idx }, '', idx)}
      </button>
    `).join('');
  };
  renderPreviewAndGrid();
  modal.classList.remove('hidden');

  const closeModal = () => {
    modal.classList.add('hidden');
    gridEl.removeEventListener('click', onSelectStyle);
    pickImgBtn.removeEventListener('click', onPickImage);
    saveBtn.removeEventListener('click', onSave);
    cancelBtn.removeEventListener('click', closeModal);
    closeBtn.removeEventListener('click', closeModal);
    modal.removeEventListener('click', onBackdrop);
  };
  const onSelectStyle = (e) => {
    const tile = e.target.closest('[data-cstyle]');
    if (!tile) return;
    selectedIdx = Number(tile.dataset.cstyle);
    useCustom = false;
    renderPreviewAndGrid();
  };
  const onPickImage = () => {
    pendingCoverBookId = book.id;
    closeModal();
    const coverInput = document.getElementById('cover-pick-input');
    coverInput.value = '';
    coverInput.click();
  };
  const onSave = async () => {
    await upsertBook({
      ...book,
      coverColorIndex: selectedIdx,
      customCoverDataUrl: useCustom ? book.customCoverDataUrl : null
    });
    books = await loadAllBooks();
    renderAllTabs();
    closeModal();
    showToast('Đã cập nhật ảnh bìa');
  };
  const onBackdrop = (e) => { if (e.target === modal) closeModal(); };

  gridEl.addEventListener('click', onSelectStyle);
  pickImgBtn.addEventListener('click', onPickImage);
  saveBtn.addEventListener('click', onSave);
  cancelBtn.addEventListener('click', closeModal);
  closeBtn.addEventListener('click', closeModal);
  modal.addEventListener('click', onBackdrop);
}

const reader = createReaderController({
  getSettings: () => settings,
  onUpdateSettings: updateSettings,
  onUpdateBook: async (updated) => {
    await upsertBook(updated);
    const idx = books.findIndex((b) => b.id === updated.id);
    if (idx >= 0) books[idx] = updated;
    else books.unshift(updated);
  },
  onCreateCopyBook: async (sourceBook, newTitle, pages, bodyHtml) => {
    const newId = crypto.randomUUID ? crypto.randomUUID().slice(0, 8) : Math.random().toString(36).slice(2, 10);
    const isDocx = sourceBook.kind === 'docx';
    const htmlContent = isDocx
      ? renderDocxPagesToHtml(newTitle, pages)
      : sourceBook.htmlContent.replace(/(<body[^>]*>)[\s\S]*(<\/body>)/i, `$1\n${bodyHtml}\n$2`);
    const copyDoc = {
      ...sourceBook,
      id: newId,
      title: newTitle,
      subtitle: isDocx ? `DOCX · ${Math.max(1, pages.length)} trang` : sourceBook.subtitle,
      docxPages: isDocx ? pages : sourceBook.docxPages,
      htmlContent,
      progressPercent: 0,
      scrollRatio: 0,
      addedAt: Date.now(),
      lastOpenedAt: Date.now(),
      fileSizeBytes: htmlContent.length
    };
    await upsertBook(copyDoc);
    books = await loadAllBooks();
    await reader.open(copyDoc);
    showToast(`Đã lưu bản sao "${newTitle}"`);
  },
  onShowToast: showToast,
  onCloseReader: async () => {
    books = await loadAllBooks();
    renderAllTabs();
  }
});

async function openBook(book) {
  const updated = { ...book, lastOpenedAt: Date.now() };
  await upsertBook(updated);
  books = await loadAllBooks();
  await reader.open(updated);
}

function getFilteredBooks() {
  let list = books;
  if (selectedFilter === 'READING') {
    list = list.filter((b) => b.progressPercent > 0 && b.progressPercent < 100);
  } else if (selectedFilter === 'FINISHED') {
    list = list.filter((b) => b.progressPercent >= 100);
  } else if (selectedFilter === 'FAVORITE') {
    list = list.filter((b) => b.isFavorite);
  }
  const q = searchQuery.trim().toLowerCase();
  if (q) {
    list = list.filter((b) =>
      (b.title || '').toLowerCase().includes(q) ||
      (b.subtitle || '').toLowerCase().includes(q) ||
      (b.groupName || '').toLowerCase().includes(q)
    );
  }
  return list;
}

function statusBadgeFor(b) {
  const pct = Number(b.progressPercent || 0);
  if (pct >= 100) return 'Xong';
  if (pct <= 0) return 'Mới';
  return `${pct}%`;
}

function renderBookGrid(container, list) {
  if (!list.length) {
    container.innerHTML = '';
    return;
  }
  container.innerHTML = list.map((b) => `
    <article class="book-tile" data-book-id="${escapeAttr(b.id)}">
      <div class="book-cover-wrap">
        ${renderCoverMarkup(b, statusBadgeFor(b))}
      </div>
      <div class="book-tile-title">${escapeAttr(b.title)}</div>
      ${b.groupName ? `<div class="book-tile-group">${escapeAttr(b.groupName)}</div>` : ''}
    </article>
  `).join('');
}

function renderLibraryTab() {
  const heroWrap = document.getElementById('library-hero-wrap');
  const gridEl = document.getElementById('library-books-grid');
  const emptyEl = document.getElementById('library-empty-state');
  const clearBtn = document.getElementById('library-search-clear');
  if (clearBtn) clearBtn.classList.toggle('hidden', !searchQuery.trim());

  const heroBook = !searchQuery.trim()
    ? (books.find((b) => b.progressPercent > 0 && b.progressPercent < 100) || books[0] || null)
    : null;

  if (heroBook) {
    heroWrap.classList.remove('hidden');
    heroWrap.innerHTML = `
      <div class="hero-glass-card" data-open-hero="${escapeAttr(heroBook.id)}">
        <div class="hero-cover">${renderCoverMarkup(heroBook)}</div>
        <div class="hero-info">
          <span class="hero-eyebrow">ĐANG ĐỌC</span>
          <h3 class="hero-title">${escapeAttr(heroBook.title)}</h3>
          <p class="hero-sub">${escapeAttr(heroBook.currentChapterTitle || heroBook.subtitle || '')}</p>
          <div class="hero-progress-row">
            <div class="progress-track"><div class="progress-fill" style="width:${Math.max(6, heroBook.progressPercent || 0)}%"></div></div>
            <span class="hero-pct">${heroBook.progressPercent || 0}%</span>
            <span class="hero-continue-btn">Đọc tiếp</span>
          </div>
        </div>
      </div>
    `;
  } else {
    heroWrap.classList.add('hidden');
    heroWrap.innerHTML = '';
  }

  document.querySelectorAll('#library-filter-row .filter-pill').forEach((pill) => {
    pill.classList.toggle('selected', pill.dataset.filter === selectedFilter);
  });

  const filtered = getFilteredBooks();
  emptyEl.classList.toggle('hidden', filtered.length > 0);
  renderBookGrid(gridEl, filtered);
}

function renderGroupsTab() {
  const emptyCard = document.getElementById('groups-empty-card');
  const cardsEl = document.getElementById('groups-folder-grid');
  const booksHeader = document.getElementById('groups-books-header');
  const booksGrid = document.getElementById('groups-books-grid');

  const counts = new Map();
  let ungroupedCount = 0;
  for (const b of books) {
    const g = (b.groupName || '').trim();
    if (!g) ungroupedCount++;
    else counts.set(g, (counts.get(g) || 0) + 1);
  }

  if (emptyCard) emptyCard.classList.toggle('hidden', groups.length > 0);

  if (selectedGroup && selectedGroup !== '__UNGROUPED__' && !groups.includes(selectedGroup)) {
    selectedGroup = null;
  }

  cardsEl.innerHTML = groups.map((g, idx) => {
    const isSel = selectedGroup === g;
    const accent = GROUP_ACCENTS[idx % GROUP_ACCENTS.length];
    return `
      <div class="group-folder-card ${isSel ? 'selected' : ''}" data-select-group="${escapeAttr(g)}">
        <div class="group-card-top">
          <div class="group-folder-badge" style="background:${accent};">
            <svg class="svg-icon sm" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/></svg>
          </div>
          <div class="group-card-actions">
            <button type="button" class="group-mini-edit bounceable" data-edit-group="${escapeAttr(g)}">Sửa</button>
            <button type="button" class="group-mini-del bounceable" data-del-group="${escapeAttr(g)}" aria-label="Xóa">✕</button>
          </div>
        </div>
        <div class="group-name">${escapeAttr(g)}</div>
        <div class="group-count">${counts.get(g) || 0} tài liệu</div>
      </div>
    `;
  }).join('');

  const groupBooks = selectedGroup
    ? books.filter((b) => (b.groupName || '').trim() === selectedGroup)
    : books.filter((b) => !(b.groupName || '').trim());
  booksHeader.textContent = selectedGroup
    ? `${selectedGroup} · ${groupBooks.length} sách`
    : `Chưa phân nhóm · ${ungroupedCount} sách`;
  renderBookGrid(booksGrid, groupBooks);
}

function renderSettingsTab() {
  document.getElementById('set-val-app-theme').textContent = settings.appThemeMode;
  const fObj = FONT_FAMILIES.find((f) => f.key === settings.fontFamily);
  document.getElementById('set-val-font').textContent = fObj ? fObj.label : 'Mặc định';

  document.getElementById('set-toggle-tts').checked = Boolean(settings.readAloudEnabled);
  document.getElementById('set-toggle-keep-on').checked = Boolean(settings.keepScreenOn);
  document.getElementById('set-toggle-tap-edge').checked = Boolean(settings.tapEdgeToTurnPage);

  const opSlider = document.getElementById('set-glass-opacity-slider');
  const opVal = document.getElementById('set-glass-opacity-val');
  const pct = Math.round((settings.glassOpacity ?? 0.78) * 100);
  opSlider.value = String(pct);
  opVal.textContent = `${pct}%`;

  const spacingLabel = LINE_SPACING_LABELS[settings.lineSpacing] || 'Chuẩn';
  const fontSummary = document.getElementById('set-font-summary');
  if (fontSummary) fontSummary.textContent = `${settings.fontSizePt} pt · ${spacingLabel}`;
  const fontSlider = document.getElementById('set-font-slider');
  if (fontSlider) fontSlider.value = String(settings.fontSizePt);

  document.querySelectorAll('#set-theme-swatches .swatch-btn').forEach((btn) => {
    btn.classList.toggle('selected', btn.dataset.setTheme === settings.themeMode);
  });

  const stats = computeStorageStats(books);
  const totalBytes = stats.booksBytes + stats.coversBytes;
  document.getElementById('set-storage-summary').textContent = `${formatBytesMb(totalBytes)} · ${stats.bookCount} sách`;
}

function renderAllTabs() {
  renderLibraryTab();
  renderGroupsTab();
  renderSettingsTab();
}

function switchTab(nextTab) {
  if (nextTab === activeTab) return;
  const prevIdx = TAB_ORDER.indexOf(activeTab);
  const nextIdx = TAB_ORDER.indexOf(nextTab);
  const animClass = nextIdx >= prevIdx ? 'anim-right' : 'anim-left';
  activeTab = nextTab;

  document.querySelectorAll('.tab-view').forEach((el) => {
    el.classList.remove('active', 'anim-right', 'anim-left');
    if (el.id === `tab-${nextTab}`) {
      el.classList.add('active');
      void el.offsetWidth;
      el.classList.add(animClass);
    }
  });
  document.querySelectorAll('.taskbar-item').forEach((btn) => {
    btn.classList.toggle('selected', btn.dataset.tab === nextTab);
  });
  document.getElementById('liquid-taskbar')?.classList.remove('taskbar-hidden');
}

async function importUploadedFiles(fileList) {
  let importedCount = 0;
  let firstBook = null;
  for (const file of fileList) {
    const name = file.name || 'Tai_lieu';
    const lower = name.toLowerCase();
    try {
      if (lower.endsWith('.zip')) {
        const bytes = new Uint8Array(await file.arrayBuffer());
        const entries = await unzipBytes(bytes);
        if (entries.has('yunzybook_metadata.json')) {
          const metaText = new TextDecoder('utf-8').decode(entries.get('yunzybook_metadata.json'));
          const parsedMeta = JSON.parse(metaText);
          if (Array.isArray(parsedMeta)) {
            for (const item of parsedMeta) {
              if (!item.id || !item.title) continue;
              const restored = { ...item };
              const htmlEntry = entries.get(`books/${item.id}.html`);
              const binEntry = entries.get(`books/${item.id}.bin`);
              if (htmlEntry) restored.htmlContent = new TextDecoder('utf-8').decode(htmlEntry);
              if (binEntry) restored.rawBlob = binEntry.buffer;
              await upsertBook(restored);
              importedCount++;
            }
            continue;
          }
        }
        const htmlEntries = [...entries.keys()].filter((k) => /\.(html?|xhtml)$/i.test(k) && !k.split('/').pop().startsWith('._'));
        if (!htmlEntries.length) throw new Error('Không tìm thấy file HTML trong gói ZIP.');
        const mainPath = htmlEntries.find((k) => /index\.html?$/i.test(k)) || htmlEntries[0];
        const htmlText = new TextDecoder('utf-8').decode(entries.get(mainPath));
        const id = crypto.randomUUID ? crypto.randomUUID().slice(0, 8) : Math.random().toString(36).slice(2, 10);
        const title = extractHtmlTitle(htmlText, name.replace(/\.zip$/i, ''));
        const doc = {
          id,
          title,
          subtitle: extractFirstHeading(htmlText) || 'HTML Archive (ZIP)',
          kind: 'html',
          htmlContent: htmlText,
          coverColorIndex: hashColorIndex(title),
          progressPercent: 0,
          scrollRatio: 0,
          isFavorite: false,
          groupName: '',
          currentChapterTitle: extractFirstHeading(htmlText) || 'Chương 1',
          fileSizeBytes: file.size,
          addedAt: Date.now(),
          lastOpenedAt: 0
        };
        await upsertBook(doc);
        importedCount++;
        if (!firstBook) firstBook = doc;
      } else if (lower.endsWith('.pdf')) {
        const buf = await file.arrayBuffer();
        const id = crypto.randomUUID ? crypto.randomUUID().slice(0, 8) : Math.random().toString(36).slice(2, 10);
        const title = name.replace(/\.pdf$/i, '') || 'Tài liệu PDF';
        const doc = {
          id,
          title,
          subtitle: 'PDF ngoại tuyến',
          kind: 'pdf',
          rawBlob: buf,
          coverColorIndex: hashColorIndex(title),
          progressPercent: 0,
          scrollRatio: 0,
          isFavorite: false,
          groupName: '',
          currentChapterTitle: 'Trang 1',
          fileSizeBytes: file.size,
          addedAt: Date.now(),
          lastOpenedAt: 0
        };
        await upsertBook(doc);
        importedCount++;
        if (!firstBook) firstBook = doc;
      } else if (lower.endsWith('.docx')) {
        const bytes = new Uint8Array(await file.arrayBuffer());
        const fallbackTitle = name.replace(/\.docx$/i, '') || 'Tài liệu Word';
        const parsed = await parseDocxBytes(bytes, fallbackTitle);
        const id = crypto.randomUUID ? crypto.randomUUID().slice(0, 8) : Math.random().toString(36).slice(2, 10);
        const doc = {
          id,
          title: parsed.title,
          subtitle: `DOCX · ${parsed.pageCount} trang`,
          kind: 'docx',
          htmlContent: parsed.html,
          docxPages: parsed.pages,
          coverColorIndex: hashColorIndex(parsed.title),
          progressPercent: 0,
          scrollRatio: 0,
          isFavorite: false,
          groupName: '',
          currentChapterTitle: parsed.firstHeading || `Trang 1 / ${parsed.pageCount}`,
          fileSizeBytes: parsed.html.length + file.size,
          addedAt: Date.now(),
          lastOpenedAt: 0
        };
        await upsertBook(doc);
        importedCount++;
        if (!firstBook) firstBook = doc;
      } else {
        const htmlText = await file.text();
        const id = crypto.randomUUID ? crypto.randomUUID().slice(0, 8) : Math.random().toString(36).slice(2, 10);
        const title = extractHtmlTitle(htmlText, name.replace(/\.(html?|txt)$/i, ''));
        const wrapped = /<html/i.test(htmlText)
          ? htmlText
          : `<!DOCTYPE html><html lang="vi"><head><meta charset="utf-8"/><title>${escapeAttr(title)}</title></head><body><pre style="white-space:pre-wrap;font-family:inherit;">${escapeAttr(htmlText)}</pre></body></html>`;
        const doc = {
          id,
          title,
          subtitle: extractFirstHeading(wrapped) || 'HTML ngoại tuyến',
          kind: 'html',
          htmlContent: wrapped,
          coverColorIndex: hashColorIndex(title),
          progressPercent: 0,
          scrollRatio: 0,
          isFavorite: false,
          groupName: '',
          currentChapterTitle: extractFirstHeading(wrapped) || 'Chương 1',
          fileSizeBytes: file.size,
          addedAt: Date.now(),
          lastOpenedAt: 0
        };
        await upsertBook(doc);
        importedCount++;
        if (!firstBook) firstBook = doc;
      }
    } catch (err) {
      showToast(`Lỗi nhập ${name}: ${err.message}`);
    }
  }

  if (importedCount > 0) {
    books = await loadAllBooks();
    groups = await loadGroups();
    renderAllTabs();
    showToast(`Đã thêm ${importedCount} tài liệu`);
    if (importedCount === 1 && firstBook) {
      await openBook(firstBook);
    }
  }
}

async function exportLibraryBackupZip() {
  const all = await loadAllBooks();
  const entries = [];
  const metaList = [];
  for (const b of all) {
    const { htmlContent, rawBlob, ...meta } = b;
    metaList.push(meta);
    if (htmlContent) entries.push([`books/${b.id}.html`, htmlContent]);
    if (rawBlob) {
      const bin = rawBlob instanceof Uint8Array ? rawBlob : new Uint8Array(rawBlob);
      entries.push([`books/${b.id}.bin`, bin]);
    }
  }
  entries.unshift(['yunzybook_metadata.json', JSON.stringify(metaList)]);
  const zipBytes = await zipEntries(entries);
  downloadBlob(zipBytes, `YunzyBook_Backup_${Date.now()}.zip`, 'application/zip');
  showToast('Đã xuất file sao lưu thư viện');
}

function openBookActionModal(book) {
  const modal = document.getElementById('book-action-modal');
  const preview = document.getElementById('context-cover-preview');
  const titleEl = document.getElementById('book-action-title');
  const favLabel = document.getElementById('book-act-fav-label');

  modal.dataset.bookId = book.id;
  preview.innerHTML = renderCoverMarkup(book, statusBadgeFor(book));
  titleEl.textContent = book.title;
  favLabel.textContent = book.isFavorite ? 'Bỏ yêu thích' : 'Yêu thích';
  modal.classList.remove('hidden');
}

function initCreateHtmlModal() {
  const modal = document.getElementById('create-html-modal');
  const titleInput = document.getElementById('create-html-title');
  const codeInput = document.getElementById('create-html-code');
  const swatchRow = document.getElementById('create-html-colors');
  let selectedColorIdx = 0;

  swatchRow.innerHTML = COVER_PALETTES.slice(0, 8).map((p, idx) =>
    `<button type="button" class="color-dot bounceable ${idx === 0 ? 'selected' : ''}" data-cidx="${idx}" style="background:${p.bg};"></button>`
  ).join('');

  swatchRow.addEventListener('click', (e) => {
    const dot = e.target.closest('[data-cidx]');
    if (!dot) return;
    selectedColorIdx = Number(dot.dataset.cidx);
    swatchRow.querySelectorAll('.color-dot').forEach((d, idx) => d.classList.toggle('selected', idx === selectedColorIdx));
  });

  document.getElementById('create-html-cancel').addEventListener('click', () => modal.classList.add('hidden'));
  document.getElementById('create-html-save').addEventListener('click', async () => {
    const raw = codeInput.value.trim();
    if (!raw) {
      showToast('Hãy dán mã HTML trước khi lưu');
      return;
    }
    const cleanTitle = titleInput.value.trim() || extractHtmlTitle(raw, 'Tài liệu HTML mới');
    const htmlContent = /<html/i.test(raw)
      ? raw
      : `<!DOCTYPE html><html lang="vi"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/><title>${escapeAttr(cleanTitle)}</title></head><body>${raw}</body></html>`;
    const doc = {
      id: crypto.randomUUID ? crypto.randomUUID().slice(0, 8) : Math.random().toString(36).slice(2, 10),
      title: cleanTitle,
      subtitle: extractFirstHeading(htmlContent) || 'HTML tự tạo',
      kind: 'html',
      htmlContent,
      coverColorIndex: selectedColorIdx,
      progressPercent: 0,
      scrollRatio: 0,
      isFavorite: false,
      groupName: '',
      currentChapterTitle: 'Chương 1',
      fileSizeBytes: htmlContent.length,
      addedAt: Date.now(),
      lastOpenedAt: 0
    };
    await upsertBook(doc);
    books = await loadAllBooks();
    modal.classList.add('hidden');
    titleInput.value = '';
    codeInput.value = '';
    renderAllTabs();
    showToast(`Đã tạo "${cleanTitle}"`);
  });
}

function initTouchAndLongPressGestures() {
  let pressedEl = null;
  let holdTimer = null;
  let holdStartX = 0;
  let holdStartY = 0;

  const clearPressed = () => {
    if (pressedEl) {
      pressedEl.classList.remove('is-pressed');
      pressedEl = null;
    }
  };
  const cancelHold = () => {
    if (holdTimer) {
      clearTimeout(holdTimer);
      holdTimer = null;
    }
  };

  document.addEventListener('pointerdown', (e) => {
    const target = e.target.closest('button, .book-tile, .hero-glass-card, .group-folder-card, .settings-row, .bounceable');
    if (target) {
      clearPressed();
      pressedEl = target;
      pressedEl.classList.add('is-pressed');
    }

    const tile = e.target.closest('.book-tile');
    if (tile) {
      cancelHold();
      holdStartX = e.clientX;
      holdStartY = e.clientY;
      holdTimer = setTimeout(() => {
        holdTimer = null;
        lastLongPressAt = Date.now();
        clearPressed();
        try { navigator.vibrate?.(16); } catch (_) {}
        const b = books.find((x) => x.id === tile.dataset.bookId);
        if (b) openBookActionModal(b);
      }, 430);
    }
  }, { passive: true });

  document.addEventListener('pointermove', (e) => {
    if (holdTimer && Math.hypot(e.clientX - holdStartX, e.clientY - holdStartY) > 10) {
      cancelHold();
      clearPressed();
    }
  }, { passive: true });

  document.addEventListener('pointerup', () => { cancelHold(); clearPressed(); }, { passive: true });
  document.addEventListener('pointercancel', () => { cancelHold(); clearPressed(); }, { passive: true });

  document.getElementById('main-scroll-area').addEventListener('contextmenu', (e) => {
    const tile = e.target.closest('.book-tile');
    if (!tile) return;
    e.preventDefault();
    const b = books.find((x) => x.id === tile.dataset.bookId);
    if (b) openBookActionModal(b);
  });

  // Auto-hide taskbar cleanly on scroll down, reveal on scroll up
  const scrollArea = document.getElementById('main-scroll-area');
  const taskbar = document.getElementById('liquid-taskbar');
  let lastY = 0;
  scrollArea.addEventListener('scroll', () => {
    cancelHold();
    clearPressed();
    const y = scrollArea.scrollTop;
    if (y > 40 && y > lastY + 12) taskbar?.classList.add('taskbar-hidden');
    else if (y < lastY - 8 || y <= 20) taskbar?.classList.remove('taskbar-hidden');
    lastY = y;
  }, { passive: true });
}

function scrubThirdPartyBanners() {
  const removeInjected = () => {
    document.querySelectorAll('a[href*="netlify"], iframe[src*="netlify"], div[class*="netlify"], [id*="netlify"]').forEach((el) => {
      if (!el.closest('.app-shell')) el.remove();
    });
  };
  removeInjected();
  new MutationObserver(removeInjected).observe(document.body, { childList: true });
}

function showIosInstallModalIfFirstVisit() {
  const { isIos, isStandalone } = isIosSafari();
  const dismissed = localStorage.getItem('yunzy_ios_guide_dismissed') === '1';
  if (isIos && !isStandalone && !dismissed) {
    document.getElementById('ios-install-banner').classList.remove('hidden');
  }
}

async function bootstrap() {
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  }

  settings = await loadSettings();
  books = await loadAllBooks();
  groups = await loadGroups();

  applyGlobalVisualTheme();
  renderAllTabs();
  initCreateHtmlModal();
  initTouchAndLongPressGestures();
  scrubThirdPartyBanners();
  showIosInstallModalIfFirstVisit();

  // Taskbar tab clicks
  document.querySelectorAll('.taskbar-item').forEach((btn) => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });

  // Search & filter pills
  const searchInput = document.getElementById('library-search-input');
  searchInput.addEventListener('input', (e) => {
    searchQuery = e.target.value;
    renderLibraryTab();
  });
  document.getElementById('library-search-clear')?.addEventListener('click', () => {
    searchQuery = '';
    searchInput.value = '';
    renderLibraryTab();
  });
  document.getElementById('library-filter-row').addEventListener('click', (e) => {
    const pill = e.target.closest('[data-filter]');
    if (!pill) return;
    selectedFilter = pill.dataset.filter;
    renderLibraryTab();
  });

  // Add popover
  const addPopover = document.getElementById('add-menu-popover');
  const fileInput = document.getElementById('file-import-input');
  const backupInput = document.getElementById('backup-restore-input');
  const coverInput = document.getElementById('cover-pick-input');

  const toggleAddPopover = () => addPopover.classList.toggle('hidden');
  document.getElementById('btn-add-top').addEventListener('click', toggleAddPopover);
  document.getElementById('btn-add-empty').addEventListener('click', toggleAddPopover);
  addPopover.addEventListener('click', (e) => {
    if (e.target === addPopover) addPopover.classList.add('hidden');
  });
  document.getElementById('add-opt-pick-file').addEventListener('click', () => {
    addPopover.classList.add('hidden');
    fileInput.value = '';
    fileInput.click();
  });
  document.getElementById('add-opt-create-html').addEventListener('click', () => {
    addPopover.classList.add('hidden');
    document.getElementById('create-html-modal').classList.remove('hidden');
  });
  fileInput.addEventListener('change', (e) => {
    if (e.target.files?.length) importUploadedFiles(e.target.files);
  });
  backupInput.addEventListener('change', (e) => {
    if (e.target.files?.length) importUploadedFiles(e.target.files);
  });
  coverInput.addEventListener('change', async (e) => {
    const f = e.target.files?.[0];
    if (!f || !pendingCoverBookId) return;
    const target = books.find((b) => b.id === pendingCoverBookId);
    pendingCoverBookId = null;
    if (!target) return;
    const readerFr = new FileReader();
    readerFr.onload = async () => {
      await upsertBook({ ...target, customCoverDataUrl: readerFr.result });
      books = await loadAllBooks();
      renderAllTabs();
      showToast('Đã cập nhật ảnh bìa');
    };
    readerFr.readAsDataURL(f);
  });

  // Grid, hero & group card clicks (delegated)
  document.getElementById('main-scroll-area').addEventListener('click', async (e) => {
    if (Date.now() - lastLongPressAt < 380) return;

    const editGroupBtn = e.target.closest('[data-edit-group]');
    if (editGroupBtn) {
      const targetGroup = editGroupBtn.dataset.editGroup;
      const next = await openInputModal({ title: 'Đổi tên nhóm', value: targetGroup, placeholder: 'Tên nhóm mới' });
      const clean = next?.trim();
      if (clean && clean !== targetGroup) {
        groups = groups.map((g) => (g === targetGroup ? clean : g));
        await saveGroups(groups);
        for (const b of books.filter((x) => x.groupName === targetGroup)) {
          await upsertBook({ ...b, groupName: clean });
        }
        if (selectedGroup === targetGroup) selectedGroup = clean;
        books = await loadAllBooks();
        renderAllTabs();
      }
      return;
    }

    const delGroupBtn = e.target.closest('[data-del-group]');
    if (delGroupBtn) {
      const targetGroup = delGroupBtn.dataset.delGroup;
      const ok = await openConfirmModal({ title: 'Xóa nhóm', message: `Xóa nhóm "${targetGroup}"? Các sách sẽ giữ nguyên và chuyển về Chưa phân nhóm.` });
      if (!ok) return;
      groups = groups.filter((g) => g !== targetGroup);
      await saveGroups(groups);
      for (const b of books.filter((x) => x.groupName === targetGroup)) {
        await upsertBook({ ...b, groupName: '' });
      }
      if (selectedGroup === targetGroup) selectedGroup = null;
      books = await loadAllBooks();
      renderAllTabs();
      showToast(`Đã xóa nhóm "${targetGroup}"`);
      return;
    }

    const groupBtn = e.target.closest('[data-select-group]');
    if (groupBtn) {
      const g = groupBtn.dataset.selectGroup;
      selectedGroup = selectedGroup === g ? null : g;
      renderGroupsTab();
      return;
    }

    const heroCard = e.target.closest('[data-open-hero]');
    if (heroCard) {
      const b = books.find((x) => x.id === heroCard.dataset.openHero);
      if (b) await openBook(b);
      return;
    }

    const tile = e.target.closest('.book-tile');
    if (tile) {
      const b = books.find((x) => x.id === tile.dataset.bookId);
      if (b) await openBook(b);
    }
  });

  // Lifted Book Context Menu handlers
  const actionModal = document.getElementById('book-action-modal');
  actionModal.addEventListener('click', (e) => {
    if (e.target === actionModal) actionModal.classList.add('hidden');
  });
  document.getElementById('book-act-cover').addEventListener('click', () => {
    const b = books.find((x) => x.id === actionModal.dataset.bookId);
    actionModal.classList.add('hidden');
    if (b) openCoverEditorModal(b);
  });
  document.getElementById('book-act-rename').addEventListener('click', async () => {
    const b = books.find((x) => x.id === actionModal.dataset.bookId);
    actionModal.classList.add('hidden');
    if (!b) return;
    const next = await openInputModal({ title: 'Đổi tên tài liệu', value: b.title, placeholder: 'Nhập tên mới' });
    if (next && next.trim()) {
      await upsertBook({ ...b, title: next.trim() });
      books = await loadAllBooks();
      renderAllTabs();
      showToast('Đã đổi tên tài liệu');
    }
  });
  document.getElementById('book-act-group').addEventListener('click', () => {
    const b = books.find((x) => x.id === actionModal.dataset.bookId);
    actionModal.classList.add('hidden');
    if (b) openGroupPickerModal(b);
  });
  document.getElementById('book-act-fav').addEventListener('click', async () => {
    const b = books.find((x) => x.id === actionModal.dataset.bookId);
    actionModal.classList.add('hidden');
    if (!b) return;
    await upsertBook({ ...b, isFavorite: !b.isFavorite });
    books = await loadAllBooks();
    renderAllTabs();
    showToast(b.isFavorite ? 'Đã bỏ yêu thích' : 'Đã thêm vào yêu thích');
  });
  document.getElementById('book-act-share').addEventListener('click', async () => {
    const b = books.find((x) => x.id === actionModal.dataset.bookId);
    actionModal.classList.add('hidden');
    if (!b) return;
    if (b.kind === 'docx' && b.docxPages) {
      const bytes = await buildDocxBytesFromPages(b.docxPages);
      downloadBlob(bytes, `${b.title}.docx`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    } else if (b.kind === 'pdf' && b.rawBlob) {
      downloadBlob(b.rawBlob, `${b.title}.pdf`, 'application/pdf');
    } else {
      downloadBlob(b.htmlContent || '', `${b.title}.html`, 'text/html');
    }
    showToast('Đã tải tệp về thiết bị');
  });
  document.getElementById('book-act-delete').addEventListener('click', async () => {
    const b = books.find((x) => x.id === actionModal.dataset.bookId);
    actionModal.classList.add('hidden');
    if (!b) return;
    const ok = await openConfirmModal({ title: 'Xóa sách', message: `Xóa "${b.title}" khỏi thư viện ngoại tuyến?` });
    if (!ok) return;
    await deleteBook(b.id);
    books = await loadAllBooks();
    renderAllTabs();
    showToast(`Đã xóa ${b.title}`);
  });

  // Groups header buttons
  document.getElementById('groups-btn-new').addEventListener('click', async () => {
    const name = await openInputModal({ title: 'Tạo nhóm mới', placeholder: 'Ví dụ: Nội khoa, Giải phẫu...' });
    const clean = name?.trim();
    if (clean) {
      if (!groups.includes(clean)) {
        groups.push(clean);
        await saveGroups(groups);
      }
      selectedGroup = clean;
      renderGroupsTab();
      showToast(`Đã tạo nhóm "${clean}"`);
    }
  });
  document.getElementById('groups-btn-edit').addEventListener('click', async () => {
    const target = selectedGroup;
    if (!target) {
      showToast('Chạm vào một thẻ nhóm bên dưới để sửa hoặc xóa');
      return;
    }
    const next = await openInputModal({ title: `Đổi tên nhóm "${target}"`, value: target, placeholder: 'Tên nhóm mới' });
    const clean = next?.trim();
    if (clean && clean !== target) {
      groups = groups.map((g) => (g === target ? clean : g));
      await saveGroups(groups);
      for (const b of books.filter((x) => x.groupName === target)) {
        await upsertBook({ ...b, groupName: clean });
      }
      selectedGroup = clean;
      books = await loadAllBooks();
      renderAllTabs();
    }
  });

  // Settings handlers
  document.getElementById('set-row-app-theme').addEventListener('click', () => {
    const idx = (APP_THEMES.indexOf(settings.appThemeMode) + 1) % APP_THEMES.length;
    updateSettings({ ...settings, appThemeMode: APP_THEMES[idx] });
  });
  document.getElementById('set-row-font').addEventListener('click', () => {
    const curIdx = FONT_FAMILIES.findIndex((f) => f.key === settings.fontFamily);
    const next = FONT_FAMILIES[(curIdx + 1) % FONT_FAMILIES.length];
    updateSettings({ ...settings, fontFamily: next.key });
  });
  document.getElementById('set-toggle-tts').addEventListener('change', (e) => {
    updateSettings({ ...settings, readAloudEnabled: e.target.checked });
  });
  document.getElementById('set-toggle-keep-on').addEventListener('change', (e) => {
    updateSettings({ ...settings, keepScreenOn: e.target.checked });
  });
  document.getElementById('set-toggle-tap-edge').addEventListener('change', (e) => {
    updateSettings({ ...settings, tapEdgeToTurnPage: e.target.checked });
  });
  document.getElementById('set-glass-opacity-slider').addEventListener('input', (e) => {
    const op = Math.max(25, Math.min(100, Number(e.target.value))) / 100;
    updateSettings({ ...settings, glassOpacity: op });
  });
  document.getElementById('set-theme-swatches').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-set-theme]');
    if (btn) updateSettings({ ...settings, themeMode: btn.dataset.setTheme });
  });
  document.getElementById('set-font-slider')?.addEventListener('input', (e) => {
    const sz = Math.max(12, Math.min(28, Number(e.target.value)));
    updateSettings({ ...settings, fontSizePt: sz });
  });
  document.getElementById('set-font-minus').addEventListener('click', () => {
    updateSettings({ ...settings, fontSizePt: Math.max(12, settings.fontSizePt - 1) });
  });
  document.getElementById('set-font-plus').addEventListener('click', () => {
    updateSettings({ ...settings, fontSizePt: Math.min(28, settings.fontSizePt + 1) });
  });
  document.getElementById('set-btn-backup').addEventListener('click', exportLibraryBackupZip);
  document.getElementById('set-btn-restore').addEventListener('click', () => {
    backupInput.value = '';
    backupInput.click();
  });

  // iPhone install guide triggers
  const iosModal = document.getElementById('ios-guide-modal');
  document.getElementById('set-row-ios-install').addEventListener('click', () => iosModal.classList.remove('hidden'));
  document.getElementById('ios-banner-open').addEventListener('click', () => iosModal.classList.remove('hidden'));
  document.getElementById('ios-banner-dismiss').addEventListener('click', () => {
    localStorage.setItem('yunzy_ios_guide_dismissed', '1');
    document.getElementById('ios-install-banner').classList.add('hidden');
  });
  document.getElementById('ios-guide-close').addEventListener('click', () => iosModal.classList.add('hidden'));
  iosModal.addEventListener('click', (e) => {
    if (e.target === iosModal) iosModal.classList.add('hidden');
  });
}

bootstrap();
