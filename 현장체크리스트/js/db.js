// IndexedDB 저장소. 모든 데이터는 이 기기의 브라우저 안에만 저장된다 (외부 전송 없음).
const DB_NAME = "hygiene-checklist";
const DB_VERSION = 2;
// master: 현재 점검 기준정보 (key "current")
// rounds: 점검 회차 (keyPath id). 만들 때 기준정보를 복사해 두어, 나중에 기준정보가 바뀌어도 지난 회차는 그대로 유지
// photos: 지적사항 사진 (keyPath id, roundId 색인). 사진은 휴대폰 갤러리가 아니라 여기에만 저장
const STORES = {
  master: {},
  rounds: { keyPath: "id" },
  photos: { keyPath: "id", indexes: [["roundId", "roundId"]] },
};

let dbPromise;
function open() {
  if (!dbPromise) {
    dbPromise = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        for (const [name, { indexes = [], ...opts }] of Object.entries(STORES)) {
          if (db.objectStoreNames.contains(name)) continue;
          const store = db.createObjectStore(name, opts);
          for (const [idx, path] of indexes) store.createIndex(idx, path);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  return dbPromise;
}

async function run(store, mode, fn) {
  const db = await open();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, mode);
    const req = fn(tx.objectStore(store));
    tx.oncomplete = () => resolve(req ? req.result : undefined);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}

export const db = {
  getMaster: () => run("master", "readonly", (s) => s.get("current")),
  putMaster: (m) => run("master", "readwrite", (s) => s.put(m, "current")),
  listRounds: () => run("rounds", "readonly", (s) => s.getAll()),
  getRound: (id) => run("rounds", "readonly", (s) => s.get(id)),
  // touch=true 이면 회차의 마지막 변경 시각을 기록 (백업·내보내기 기록만 남길 때는 false)
  putRound: (r, touch = true) => {
    if (touch) r.changedAt = new Date().toISOString();
    return run("rounds", "readwrite", (s) => s.put(r));
  },
  deleteRound: (id) => run("rounds", "readwrite", (s) => s.delete(id)),
  putPhoto: (p) => run("photos", "readwrite", (s) => s.put(p)),
  getPhoto: (id) => run("photos", "readonly", (s) => s.get(id)),
  deletePhotos: (ids) => run("photos", "readwrite", (s) => { ids.forEach((id) => s.delete(id)); return null; }),
  photosByRound: (roundId) => run("photos", "readonly", (s) => s.index("roundId").getAll(roundId)),
};

// 브라우저가 저장 공간이 부족할 때 데이터를 임의로 지우지 않도록 요청 (지원하는 브라우저만)
export async function requestPersistentStorage() {
  try { return navigator.storage?.persist ? await navigator.storage.persist() : false; } catch { return false; }
}
