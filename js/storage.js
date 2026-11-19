/**
 * Offline IndexedDB storage for YunzyBook PWA (split metadata & heavy contents).
 */

const DB_NAME = 'yunzybook_pwa_db';
const DB_VERSION = 2;
const STORE_BOOKS = 'books';
const STORE_CONTENTS = 'book_contents';
const STORE_META = 'meta';
const CONTENT_KEYS = ['htmlContent', 'rawBlob', 'docxPages'];

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
  hapticEnabled: true,
  glassOpacity: 0.72
};

let dbPromise = null;
let migrationPromise = null;

function openDbConnection() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_BOOKS)) {
        db.createObjectStore(STORE_BOOKS, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(STORE_CONTENTS)) {
        db.createObjectStore(STORE_CONTENTS, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(STORE_META)) {
        db.createObjectStore(STORE_META, { keyPath: 'key' });
      }
    };
    req.onsuccess = () => {
      const db = req.result;
      db.onclose = () => {
        dbPromise = null;
      };
      db.onversionchange = () => {
        db.close();
        dbPromise = null;
      };
      resolve(db);
    };
    req.onerror = () => {
      dbPromise = null;
      reject(req.error);
    };
    req.onblocked = () => {
      dbPromise = null;
    };
  });
}

function getDb() {
  if (!dbPromise) {
    dbPromise = openDbConnection().catch((err) => {
      dbPromise = null;
      throw err;
    });
  }
  return dbPromise;
}

async function withDb(fn) {
  const db = await getDb();
  try {
    return await fn(db);
  } catch (err) {
    const name = err?.name || '';
    if (name === 'InvalidStateError' || name === 'UnknownError') {
      try { db.close(); } catch (_) {}
      dbPromise = null;
      const freshDb = await getDb();
      return await fn(freshDb);
    }
    throw err;
  }
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

function stripBookContent(book) {
  if (!book || typeof book !== 'object') return book;
  const copy = { ...book };
  for (const key of CONTENT_KEYS) {
    delete copy[key];
  }
  return copy;
}

function hasContentFields(book) {
  return CONTENT_KEYS.some(
    (k) => book && Object.prototype.hasOwnProperty.call(book, k) && book[k] !== undefined
  );
}

function extractContentRecord(book) {
  return {
    id: book.id,
    htmlContent: book.htmlContent ?? null,
    rawBlob: book.rawBlob ?? null,
    docxPages: book.docxPages ?? null
  };
}

function migrateInlineContentsIfNeeded() {
  if (!migrationPromise) {
    migrationPromise = withDb((db) => new Promise((resolve, reject) => {
      const tx = db.transaction([STORE_BOOKS, STORE_CONTENTS, STORE_META], 'readwrite');
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error || new Error('Transaction aborted'));

      const metaStore = tx.objectStore(STORE_META);
      const booksStore = tx.objectStore(STORE_BOOKS);
      const contentsStore = tx.objectStore(STORE_CONTENTS);

      const flagReq = metaStore.get('contents_migrated_v2');
      flagReq.onsuccess = () => {
        if (flagReq.result?.value === true) return;
        const cursorReq = booksStore.openCursor();
        cursorReq.onsuccess = (e) => {
          const cursor = e.target.result;
          if (!cursor) {
            metaStore.put({ key: 'contents_migrated_v2', value: true });
            return;
          }
          const val = cursor.value;
          if (hasContentFields(val)) {
            contentsStore.put(extractContentRecord(val));
            cursor.update(stripBookContent(val));
          }
          cursor.continue();
        };
      };
    })).catch((err) => {
      migrationPromise = null;
      throw err;
    });
  }
  return migrationPromise;
}

export async function loadAllBooks() {
  await migrateInlineContentsIfNeeded();
  const rows = await withDb((db) => {
    const tx = db.transaction(STORE_BOOKS, 'readonly');
    return reqPromise(tx.objectStore(STORE_BOOKS).getAll());
  });
  return (rows || [])
    .map(stripBookContent)
    .sort((a, b) => (b.lastOpenedAt || b.addedAt || 0) - (a.lastOpenedAt || a.addedAt || 0));
}

export async function getBookById(id) {
  const row = await withDb((db) => {
    const tx = db.transaction(STORE_BOOKS, 'readonly');
    return reqPromise(tx.objectStore(STORE_BOOKS).get(id));
  });
  return row ? stripBookContent(row) : null;
}

export async function getBookContent(id) {
  await migrateInlineContentsIfNeeded();
  return withDb((db) => new Promise((resolve, reject) => {
    const tx = db.transaction([STORE_CONTENTS, STORE_BOOKS], 'readonly');
    let result = null;
    tx.oncomplete = () => resolve(result);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error('Transaction aborted'));

    const cReq = tx.objectStore(STORE_CONTENTS).get(id);
    cReq.onsuccess = () => {
      if (cReq.result) {
        result = cReq.result;
        return;
      }
      const bReq = tx.objectStore(STORE_BOOKS).get(id);
      bReq.onsuccess = () => {
        const legacy = bReq.result;
        if (legacy && hasContentFields(legacy)) {
          result = extractContentRecord(legacy);
        }
      };
    };
  }));
}

export async function upsertBook(book) {
  const meta = stripBookContent(book);
  const includeContent = hasContentFields(book);
  const contentRecord = includeContent ? extractContentRecord(book) : null;

  await withDb((db) => {
    const stores = includeContent ? [STORE_BOOKS, STORE_CONTENTS] : [STORE_BOOKS];
    const tx = db.transaction(stores, 'readwrite');
    const done = txDone(tx);
    tx.objectStore(STORE_BOOKS).put(meta);
    if (includeContent) {
      tx.objectStore(STORE_CONTENTS).put(contentRecord);
    }
    return done;
  });
  return meta;
}

export async function updateBookMeta(id, patch) {
  const cleanPatch = stripBookContent(patch);
  return withDb((db) => new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_BOOKS, 'readwrite');
    let next = null;
    tx.oncomplete = () => resolve(next);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error('Transaction aborted'));

    const store = tx.objectStore(STORE_BOOKS);
    const getReq = store.get(id);
    getReq.onsuccess = () => {
      const existing = getReq.result;
      if (!existing) return;
      next = { ...stripBookContent(existing), ...cleanPatch, id };
      store.put(next);
    };
  }));
}

export async function updateBookProgress(id, progressPatch) {
  return updateBookMeta(id, progressPatch);
}

export async function saveBookContent(id, contentPatch, metaPatch = null) {
  return withDb((db) => new Promise((resolve, reject) => {
    const stores = metaPatch ? [STORE_BOOKS, STORE_CONTENTS] : [STORE_CONTENTS];
    const tx = db.transaction(stores, 'readwrite');
    let updatedMeta = null;
    tx.oncomplete = () => resolve(updatedMeta);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error('Transaction aborted'));

    const contentsStore = tx.objectStore(STORE_CONTENTS);
    const cReq = contentsStore.get(id);
    cReq.onsuccess = () => {
      const existingContent = cReq.result || { id };
      contentsStore.put({ ...existingContent, ...contentPatch, id });
      if (metaPatch) {
        const booksStore = tx.objectStore(STORE_BOOKS);
        const mReq = booksStore.get(id);
        mReq.onsuccess = () => {
          const existingMeta = mReq.result;
          if (existingMeta) {
            updatedMeta = { ...stripBookContent(existingMeta), ...stripBookContent(metaPatch), id };
            booksStore.put(updatedMeta);
          }
        };
      }
    };
  }));
}

export async function deleteBook(id) {
  await withDb((db) => {
    const tx = db.transaction([STORE_BOOKS, STORE_CONTENTS], 'readwrite');
    const done = txDone(tx);
    tx.objectStore(STORE_BOOKS).delete(id);
    tx.objectStore(STORE_CONTENTS).delete(id);
    return done;
  });
}

export async function loadGroups() {
  const { savedRow, books } = await withDb(async (db) => {
    const tx = db.transaction([STORE_META, STORE_BOOKS], 'readonly');
    const [savedRow, books] = await Promise.all([
      reqPromise(tx.objectStore(STORE_META).get('user_groups')),
      reqPromise(tx.objectStore(STORE_BOOKS).getAll())
    ]);
    return { savedRow, books };
  });

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
  await withDb((db) => {
    const tx = db.transaction(STORE_META, 'readwrite');
    const done = txDone(tx);
    tx.objectStore(STORE_META).put({ key: 'user_groups', value: groups });
    return done;
  });
}

export async function loadSettings() {
  const row = await withDb((db) => {
    const tx = db.transaction(STORE_META, 'readonly');
    return reqPromise(tx.objectStore(STORE_META).get('appearance_settings'));
  });
  return { ...DEFAULT_SETTINGS, ...(row?.value || {}) };
}

export async function saveSettings(settings) {
  const clean = { ...DEFAULT_SETTINGS, ...settings };
  await withDb((db) => {
    const tx = db.transaction(STORE_META, 'readwrite');
    const done = txDone(tx);
    tx.objectStore(STORE_META).put({ key: 'appearance_settings', value: clean });
    return done;
  });
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
