// Simple IndexedDB wrapper for workbook snapshots
/* eslint-disable prettier/prettier */
// Store: "snapshots" with keyPath "id"
/* global indexedDB, IDBKeyRange */

const DB_NAME = "cc_snapshots_v1";
const DB_VERSION = 2;
const STORE = "snapshots";
const MAX_SNAPSHOTS_PER_WORKBOOK = 50; // retention cap per workbook

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      let store;
      if (!db.objectStoreNames.contains(STORE)) {
        store = db.createObjectStore(STORE, { keyPath: "id" });
      } else {
        store = req.transaction.objectStore(STORE);
      }
      // Ensure indexes exist
      const existingIndexes = Array.from(store.indexNames || []);
      if (!existingIndexes.includes("ts")) {
        store.createIndex("ts", "ts", { unique: false });
      }
      if (!existingIndexes.includes("workbookId")) {
        store.createIndex("workbookId", "workbookId", { unique: false });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function runTxn(mode, fn) {
  return openDb().then(
    (db) =>
      new Promise((resolve, reject) => {
        const tx = db.transaction(STORE, mode);
        const store = tx.objectStore(STORE);
        Promise.resolve(fn(store))
          .then((res) => {
            tx.oncomplete = () => resolve(res);
            tx.onerror = () => reject(tx.error);
            tx.onabort = () => reject(tx.error);
          })
          .catch(reject);
      })
  );
}

function genId() {
  const rand = Math.random().toString(36).slice(2, 8);
  return `${Date.now()}_${rand}`;
}

export async function saveSnapshot(model, meta = {}) {
  const record = {
    id: genId(),
    name: meta.name || "Snapshot",
    ts: Date.now(),
    sheetCount: Array.isArray(model?.sheets) ? model.sheets.length : 0,
    workbookId: meta.workbookId || null,
    model,
  };
  await runTxn("readwrite", (store) => store.put(record));
  // Best-effort prune for this workbook to avoid unbounded growth
  try {
    if (record.workbookId) {
      await pruneSnapshotsForWorkbook(record.workbookId, MAX_SNAPSHOTS_PER_WORKBOOK);
    }
  } catch (_) {
    // ignore pruning errors
  }
  return record;
}

export async function listSnapshots() {
  return runTxn(
    "readonly",
    (store) =>
      new Promise((resolve, reject) => {
        const results = [];
        const idx = store.index("ts");
        const req = idx.openCursor(null, "prev"); // newest first
        req.onsuccess = () => {
          const cursor = req.result;
          if (cursor) {
            results.push(cursor.value);
            cursor.continue();
          } else {
            resolve(results);
          }
        };
        req.onerror = () => reject(req.error);
      })
  );
}

export async function listSnapshotsByWorkbook(workbookId) {
  if (!workbookId) return [];
  return runTxn(
    "readonly",
    (store) =>
      new Promise((resolve, reject) => {
        const results = [];
        let idx;
        try {
          idx = store.index("workbookId");
        } catch (_) {
          // Index might be missing if DB didn't upgrade; fallback to listSnapshots and filter
          return listSnapshots()
            .then((all) => resolve(all.filter((r) => r.workbookId === workbookId)))
            .catch(reject);
        }
        const range = IDBKeyRange.only(workbookId);
        const req = idx.openCursor(range, "prev");
        req.onsuccess = () => {
          const cursor = req.result;
          if (cursor) {
            results.push(cursor.value);
            cursor.continue();
          } else {
            // Sort by ts desc just in case index order differs
            results.sort((a, b) => (b.ts || 0) - (a.ts || 0));
            resolve(results);
          }
        };
        req.onerror = () => reject(req.error);
      })
  );
}

async function pruneSnapshotsForWorkbook(workbookId, max = MAX_SNAPSHOTS_PER_WORKBOOK) {
  if (!workbookId || max <= 0) return;
  const items = await listSnapshotsByWorkbook(workbookId);
  if (items.length <= max) return;
  const toDelete = items.slice(max); // items are newest-first
  await Promise.allSettled(toDelete.map((r) => deleteSnapshot(r.id)));
}

export async function deleteSnapshot(id) {
  if (!id) return false;
  await runTxn(
    "readwrite",
    (store) =>
      new Promise((resolve, reject) => {
        const req = store.delete(id);
        req.onsuccess = () => resolve(true);
        req.onerror = () => reject(req.error);
      })
  );
  return true;
}

export async function getSnapshot(id) {
  if (!id) return null;
  return runTxn(
    "readonly",
    (store) =>
      new Promise((resolve, reject) => {
        const req = store.get(id);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => reject(req.error);
      })
  );
}
