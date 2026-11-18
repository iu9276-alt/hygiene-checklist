// 오프라인 동작 (PRD F10). 화면 파일을 기기에 저장해 두고, 통신이 끊기면 저장해 둔 것으로 연다.
// 입력 데이터·사진은 원래부터 IndexedDB(기기 안)에 저장되므로 통신과 무관.
// 방식: 먼저 네트워크(최대 3초) → 실패하면 저장해 둔 파일. 새 버전을 올리면 다음 접속 때 자동으로 바뀜.
const CACHE = "hygiene-checklist-v8";
const ASSETS = [
  "./", "./index.html", "./manifest.json", "./css/app.css",
  "./js/app.js", "./js/db.js", "./js/xlsx.js", "./js/master.js", "./js/entry.js", "./js/photo.js",
  "./js/summary.js", "./js/export.js", "./js/xlsx-write.js", "./js/zip.js", "./js/followup.js", "./js/backup.js", "./js/notice.js", "./js/camera.js",
  "./icons/icon-192.png", "./icons/icon-512.png",
];
const TIMEOUT_MS = 3000;

self.addEventListener("install", (ev) => {
  ev.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (ev) => {
  ev.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim()),
  );
});

self.addEventListener("fetch", (ev) => {
  const req = ev.request;
  if (req.method !== "GET" || new URL(req.url).origin !== self.location.origin) return;
  ev.respondWith((async () => {
    const cache = await caches.open(CACHE);
    // no-cache: 브라우저가 예전 파일을 그대로 쓰지 않고 서버에 바뀌었는지 꼭 확인 (업데이트 후 새·옛 파일이 섞이는 것 방지)
    const net = fetch(req.url, { cache: "no-cache", credentials: "same-origin" }).then((res) => {
      if (res.ok) cache.put(req, res.clone());
      return res;
    });
    try {
      return await Promise.race([net, new Promise((_, rej) => setTimeout(() => rej(new Error("timeout")), TIMEOUT_MS))]);
    } catch {
      const hit = await cache.match(req, { ignoreSearch: true })
        || (req.mode === "navigate" ? await cache.match("./index.html") : null);
      if (hit) return hit;
      return net; // 저장해 둔 것도 없으면 네트워크 결과를 끝까지 기다림
    }
  })());
});
