// 시정 기한 관리 (PRD F9, 시나리오 C)
// 업무_맥락.md: 지적일 +14일 안에 재방문해서 확인해야 시정 완료.
// 판정 규칙은 보고서 도구(make_report.py findings_of)와 같게 맞춘다.
import { entryKey } from "./entry.js";

export const DEADLINE_DAYS = 14;
const DAY = 86400000;

const toUtc = (iso) => {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
  return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) : null;
};
export const addDays = (iso, n) => {
  const t = toUtc(iso);
  return t == null ? null : new Date(t + n * DAY).toISOString().slice(0, 10);
};
export const daysBetween = (a, b) => Math.round((toUtc(b) - toUtc(a)) / DAY); // b - a

/**
 * @returns {{ state: 'done'|'check'|'overdue'|'open'|'nodate', label: string, due: string|null, dday: number|null }}
 */
export function followupStatus(e, today) {
  const d0 = e.지적일 || e.점검일 || null;
  const due = d0 ? addDays(d0, DEADLINE_DAYS) : null;
  const done = ["Y", "O", "완료", "예"].includes(String(e.완료여부 || "").trim().toUpperCase());
  const rv = e.재방문일 || null;
  if (done && rv && due && rv <= due) return { state: "done", label: "시정 완료", due, dday: null };
  if (done) return { state: "check", label: rv ? "확인 필요(기한 후 확인)" : "확인 필요(재방문일 없음)", due, dday: null };
  if (!due) return { state: "nodate", label: "확인 필요(지적일 없음)", due, dday: null };
  const left = daysBetween(today, due);
  if (left < 0) return { state: "overdue", label: `기한 초과(${-left}일)`, due, dday: left };
  return { state: "open", label: left === 0 ? "진행 중(D-day)" : `진행 중(D-${left})`, due, dday: left };
}

const ORDER = { overdue: 0, nodate: 1, check: 2, open: 3, done: 4 };

/** 회차의 저장된 부적합 항목 전체를 시정 상태와 함께, 급한 순서로 */
export function followupList(round, today) {
  const out = [];
  round.master.stops.forEach((stop, si) => stop.shops.forEach((shop, sj) => round.master.items.forEach((item, k) => {
    const key = entryKey(si, sj, k);
    const e = round.entries?.[key];
    if (!e || e.결과 !== "부적합") return;
    out.push({ key, si, sj, k, stop: stop.name, shop: shop.name, item: item.name, e, st: followupStatus(e, today) });
  })));
  out.sort((a, b) => ORDER[a.st.state] - ORDER[b.st.state]
    || (a.st.due || "").localeCompare(b.st.due || "") || a.key.localeCompare(b.key, undefined, { numeric: true }));
  return out;
}

/** 재방문 기록 검사. 문제가 있으면 문구 배열 */
export function checkRevisit({ 재방문일, 완료여부 }, e, today) {
  const p = [];
  const d0 = e.지적일 || e.점검일;
  if (재방문일) {
    if (!toUtc(재방문일)) p.push("재방문일 형식이 올바르지 않습니다.");
    else {
      if (d0 && 재방문일 < d0) p.push(`재방문일이 지적일(${d0})보다 빠릅니다.`);
      if (재방문일 > today) p.push("재방문일이 오늘보다 늦습니다.");
    }
  }
  if (완료여부 === "Y" && !재방문일) p.push("시정 완료로 기록하려면 재방문일이 필요합니다.");
  return p;
}
