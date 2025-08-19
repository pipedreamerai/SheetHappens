// Simple IndexedDB wrapper for workbook snapshots
// Store: "snapshots" with keyPath "id"
/* global indexedDB */

const DB_NAME = "cc_snapshots_v1";
const DB_VERSION = 1;
const STORE = "snapshots";

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) {
        const store = db.createObjectStore(STORE, { keyPath: "id" });
        store.createIndex("ts", "ts", { unique: false });
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
    model,
  };
  await runTxn("readwrite", (store) => store.put(record));
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
