// 내보내기 (PRD F8): 엑셀 + 사진을 zip 하나로.
//   위생점검_{연도}_{n}Q.zip
//   ├─ 위생점검결과_{연도}_{n}Q.xlsx   휴게소별 시트, 보고서 도구(make_report.py) 입력 형식
//   └─ 사진/{YYYYMMDD}/{휴게소}_{점검항목}[_{입점매장}][_조치후][_{n}].jpg
//      (지적 사진은 점검일 폴더, 조치 후 사진은 재방문일 폴더)
//
// 보고서와 숫자가 같도록(2026-11-18 평균점수 통일 결정) 모든 칸을 한 행씩 내보낸다.
//   - 저장된 항목: 입력값 그대로
//   - 작성중 항목: 결과는 비우고(보고서에서 '판정누락'으로 표시), 점수·지적내용·사진은 그대로,
//                 비고에 "작성중(입력값: …) — 빠진 값" 기록
//   - 미입력 항목: 점검일·매장·항목만, 비고 "미입력"
import { entryKey, checkEntry, localDate } from "./entry.js";
import { assignFileNames } from "./photo.js";
import { maskName, maskPhone } from "./master.js";
import { buildXlsx } from "./xlsx-write.js";
import { makeZip } from "./zip.js";

export const COLUMNS = ["점검일", "입점매장", "점검항목", "결과", "점수", "지적내용", "매장담당자", "연락처",
  "지적일", "재방문일", "완료여부", "사진파일", "조치사진파일", "비고"];
const DATE_COLS = ["점검일", "지적일", "재방문일"];
const WIDTHS = { 점검일: 11, 입점매장: 12, 점검항목: 16, 결과: 8, 점수: 6, 지적내용: 40, 매장담당자: 11, 연락처: 15,
  지적일: 11, 재방문일: 11, 완료여부: 9, 사진파일: 34, 조치사진파일: 34, 비고: 40 };

export const exportBaseName = (round) => `${round.year}_${round.quarter}Q`;

/**
 * @param {object} round
 * @param {Map<string, object>} photos  photoId → 사진 레코드 (blob, createdAt)
 * @param {{ mask?: boolean, today?: string }} opt
 */
export function buildRows(round, photos, { mask = false, today = localDate() } = {}) {
  const names = assignFileNames(round, photos);
  const files = []; // { path, id }
  const stats = { rows: 0, saved: 0, draft: 0, empty: 0, photos: 0, missingPhotos: [] };
  const sheets = round.master.stops.map((stop, si) => {
    const rows = [];
    stop.shops.forEach((shop, sj) => {
      round.master.items.forEach((item, k) => {
        const key = entryKey(si, sj, k);
        const saved = round.entries?.[key];
        const draft = round.drafts?.[key];
        const e = saved || draft;
        const row = {
          입점매장: shop.name, 점검항목: item.name,
          매장담당자: mask ? maskName(shop.person) : shop.person,
          연락처: mask ? maskPhone(shop.phone) : shop.phone,
        };
        stats.rows++;
        if (!e) {
          stats.empty++;
          row.비고 = "미입력";
          rows.push(row);
          return;
        }
        const date = e.점검일 || today;
        const photoNames = [];
        for (const id of e.사진 || []) {
          const p = photos.get(id);
          const fname = names.get(id);
          if (!p || !fname) { stats.missingPhotos.push({ key, id }); continue; }
          files.push({ path: `사진/${date.replace(/-/g, "")}/${fname}`, id });
          photoNames.push(fname);
        }
        // 조치 후 사진: 재방문일 폴더 (재방문일이 없으면 찍은 날 폴더)
        const afterNames = [];
        for (const id of e.조치사진 || []) {
          const p = photos.get(id);
          const fname = names.get(id);
          if (!p || !fname) { stats.missingPhotos.push({ key, id }); continue; }
          const day = e.재방문일 || localDate(new Date(p.createdAt));
          files.push({ path: `사진/${day.replace(/-/g, "")}/${fname}`, id });
          afterNames.push(fname);
        }
        Object.assign(row, {
          점검일: e.점검일 || null,
          점수: typeof e.점수 === "number" ? e.점수 : null,
          지적내용: mask ? maskPhone(e.지적내용 || "") : e.지적내용 || "",
          지적일: e.지적일 || null,
          재방문일: e.재방문일 || null,
          완료여부: e.완료여부 || "",
          사진파일: photoNames.join(", "),
          조치사진파일: afterNames.join(", "),
        });
        if (saved) {
          stats.saved++;
          row.결과 = saved.결과;
        } else {
          stats.draft++;
          row.결과 = null;
          row.비고 = `작성중(입력값: ${draft.결과 || "판정 없음"}) — ${checkEntry(draft, item).problems.join(" ")}`;
        }
        rows.push(row);
      });
    });
    return { name: stop.name, header: COLUMNS, rows, dateCols: DATE_COLS, widths: WIDTHS };
  });
  stats.photos = files.length;
  return { sheets, files, stats };
}

/** @returns {Promise<{ blob: Blob, fileName: string, stats: object }>} */
export async function buildExportZip(round, photos, opt = {}) {
  const { sheets, files, stats } = buildRows(round, photos, opt);
  const base = exportBaseName(round);
  const xlsx = await buildXlsx(sheets);
  const entries = [{ name: `위생점검결과_${base}.xlsx`, data: xlsx }];
  for (const f of files) {
    entries.push({ name: f.path, data: new Uint8Array(await photos.get(f.id).blob.arrayBuffer()) });
  }
  return { blob: makeZip(entries), fileName: `위생점검_${base}.zip`, stats };
}
