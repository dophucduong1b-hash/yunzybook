/**
 * Offline IndexedDB storage for YunzyBook PWA (books, groups, appearance settings).
 */

const DB_NAME = 'yunzybook_pwa_db';
const DB_VERSION = 1;
const STORE_BOOKS = 'books';
const STORE_META = 'meta';

export const DEFAULT_SETTINGS = {
  themeMode: 'ORIGINAL',      // ORIGINAL | AUTO | LIGHT | SEPIA | DARK
  fontSizePt: 17,             // 12..28
  fontFamily: 'ORIGINAL',     // ORIGINAL | MODERN | CLASSIC | SYSTEM
  lineSpacing: 'NORMAL',      // COMPACT | NORMAL | RELAXED
  pageMode: 'SCROLL',         // SCROLL | PAGINATED
  brightness: 0.72,           // 0.15..1.0
  appThemeMode: 'Tự động',    // Tự động | Sáng | Sepia | Tối
  readAloudEnabled: false,
  keepScreenOn: true,
  tapEdgeToTurnPage: false,
  glassOpacity: 0.72
};

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_BOOKS)) {
        db.createObjectStore(STORE_BOOKS, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(STORE_META)) {
        db.createObjectStore(STORE_META, { keyPath: 'key' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function txDone(tx) {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error('Transaction aborted'));
  });
}

function reqPromise(req) {
  return new Promise((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export async function loadAllBooks() {
  const db = await openDb();
  const tx = db.transaction(STORE_BOOKS, 'readonly');
  const rows = await reqPromise(tx.objectStore(STORE_BOOKS).getAll());
  db.close();
  return (rows || []).sort((a, b) => (b.lastOpenedAt || b.addedAt || 0) - (a.lastOpenedAt || a.addedAt || 0));
}

export async function getBookById(id) {
  const db = await openDb();
  const tx = db.transaction(STORE_BOOKS, 'readonly');
  const row = await reqPromise(tx.objectStore(STORE_BOOKS).get(id));
  db.close();
  return row || null;
}

export async function upsertBook(book) {
  const db = await openDb();
  const tx = db.transaction(STORE_BOOKS, 'readwrite');
  tx.objectStore(STORE_BOOKS).put(book);
  await txDone(tx);
  db.close();
  return book;
}

export async function deleteBook(id) {
  const db = await openDb();
  const tx = db.transaction(STORE_BOOKS, 'readwrite');
  tx.objectStore(STORE_BOOKS).delete(id);
  await txDone(tx);
  db.close();
}

export async function loadGroups() {
  const db = await openDb();
  const tx = db.transaction([STORE_META, STORE_BOOKS], 'readonly');
  const savedRow = await reqPromise(tx.objectStore(STORE_META).get('user_groups'));
  const books = await reqPromise(tx.objectStore(STORE_BOOKS).getAll());
  db.close();

  const saved = Array.isArray(savedRow?.value) ? [...savedRow.value] : [];
  const seen = new Set(saved.map((g) => g.toLowerCase()));
  let changed = false;
  for (const b of books || []) {
    const g = (b.groupName || '').trim();
    if (g && !seen.has(g.toLowerCase())) {
      saved.push(g);
      seen.add(g.toLowerCase());
      changed = true;
    }
  }
  if (changed) await saveGroups(saved);
  return saved;
}

export async function saveGroups(groups) {
  const db = await openDb();
  const tx = db.transaction(STORE_META, 'readwrite');
  tx.objectStore(STORE_META).put({ key: 'user_groups', value: groups });
  await txDone(tx);
  db.close();
}

export async function loadSettings() {
  const db = await openDb();
  const tx = db.transaction(STORE_META, 'readonly');
  const row = await reqPromise(tx.objectStore(STORE_META).get('appearance_settings'));
  db.close();
  return { ...DEFAULT_SETTINGS, ...(row?.value || {}) };
}

export async function saveSettings(settings) {
  const clean = { ...DEFAULT_SETTINGS, ...settings };
  const db = await openDb();
  const tx = db.transaction(STORE_META, 'readwrite');
  tx.objectStore(STORE_META).put({ key: 'appearance_settings', value: clean });
  await txDone(tx);
  db.close();
  return clean;
}

export function computeStorageStats(books) {
  let booksBytes = 0;
  let coversBytes = 0;
  for (const b of books || []) {
    booksBytes += Number(b.fileSizeBytes || 0);
    if (b.customCoverDataUrl) {
      coversBytes += b.customCoverDataUrl.length;
    }
  }
  return {
    booksBytes,
    coversBytes,
    bookCount: (books || []).length
  };
}
