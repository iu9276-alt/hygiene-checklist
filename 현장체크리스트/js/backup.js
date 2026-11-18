// 회차 백업·복원 (PRD F11, 시나리오 D: 대체 담당자 인계, 기기 교체 대비)
// 백업 파일 = zip:  backup.json (회차 전체) + photos/{사진id}.jpg
// 내보내기(엑셀+사진, 보고서용)와 다르다: 백업은 웹앱에 다시 넣어 이어서 입력하기 위한 파일.
import { makeZip } from "./zip.js";
import { unzip } from "./xlsx.js";

export const FORMAT = "hygiene-checklist-backup";
export const VERSION = 1;

const stamp = (d = new Date()) => {
  const z = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${z(d.getMonth() + 1)}${z(d.getDate())}-${z(d.getHours())}${z(d.getMinutes())}`;
};

/** @param {Map<string, object>} photos photoId → 사진 레코드 */
export async function buildBackup(round, photos) {
  const meta = [];
  const files = [];
  for (const p of photos.values()) {
    const file = `photos/${p.id}.jpg`;
    meta.push({ id: p.id, key: p.key, width: p.width, height: p.height, createdAt: p.createdAt, file });
    files.push({ name: file, data: new Uint8Array(await p.blob.arrayBuffer()) });
  }
  const json = JSON.stringify({ format: FORMAT, version: VERSION, createdAt: new Date().toISOString(), round, photos: meta });
  const blob = makeZip([{ name: "backup.json", data: json }, ...files]);
  return { blob, fileName: `위생점검백업_${round.year}_${round.quarter}Q_${stamp()}.zip`, photoCount: files.length };
}

/** 백업 파일을 읽고 검사한다. 값을 고치지 않고, 문제가 있으면 오류를 던진다. */
export async function readBackup(buffer) {
  let zip;
  try { zip = await unzip(buffer); } catch { throw new Error("백업 파일(zip)이 아닙니다."); }
  const text = await zip.text("backup.json");
  if (!text) throw new Error("백업 파일이 아닙니다. (backup.json 없음) — 엑셀 내보내기 파일은 백업으로 가져올 수 없습니다.");
  let data;
  try { data = JSON.parse(text); } catch { throw new Error("백업 파일이 손상되었습니다. (backup.json 읽기 실패)"); }
  if (data.format !== FORMAT) throw new Error("이 웹앱의 백업 파일이 아닙니다.");
  if (data.version !== VERSION) throw new Error(`지원하지 않는 백업 버전입니다. (${data.version})`);
  const r = data.round;
  if (!r || !/^\d{4}-[1-4]Q$/.test(r.id || "") || !Array.isArray(r.master?.stops) || !Array.isArray(r.master?.items)) {
    throw new Error("백업 안의 회차 정보가 올바르지 않습니다.");
  }
  const photos = [];
  for (const m of data.photos || []) {
    const bytes = await zip.bytes(m.file);
    if (!bytes) throw new Error(`백업 안에 사진 파일이 없습니다: ${m.file}`);
    photos.push({ id: m.id, roundId: r.id, key: m.key, width: m.width, height: m.height, createdAt: m.createdAt,
      blob: new Blob([bytes], { type: "image/jpeg" }) });
  }
  // 입력 기록이 가리키는 사진이 백업 안에 다 있는지
  const have = new Set(photos.map((p) => p.id));
  const missing = [];
  for (const bucket of [r.entries || {}, r.drafts || {}]) {
    for (const e of Object.values(bucket)) for (const id of e.사진 || []) if (!have.has(id)) missing.push(id);
  }
  return { round: r, photos, createdAt: data.createdAt, missing };
}

/** 회차의 마지막 변경 이후 백업·내보내기를 했는지 */
export function savedSinceChange(round) {
  const last = [...(round.backups || []), ...(round.exports || [])].map((x) => x.at).sort().at(-1) || null;
  const changed = round.changedAt || round.createdAt;
  return { lastSaved: last, changedAt: changed, safe: !!last && last >= changed };
}
