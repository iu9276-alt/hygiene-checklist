// 점검 입력 규칙 (PRD F3·F4·F5)
// - 필수값이 다 채워진 항목만 round.entries 에 "저장"된다 (진행률·적합/부적합 집계에 포함).
// - 필수값이 빠진 항목은 round.drafts 에 "작성중"으로 보관한다 (새로고침해도 입력값 유지, 진행률 제외).
//   내보낼 때는 결과 칸을 비워 보고서에서 '판정누락'으로 보이게 한다 (export.js).
// - 부적합이면 지적내용과 사진 1장 이상이 필수.

export const entryKey = (si, sj, k) => `${si}/${sj}/${k}`;

export function localDate(d = new Date()) {
  const z = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`;
}

/** v: { 결과, 점수, 지적내용, 사진[] } → { state: 'empty'|'draft'|'done', problems: string[] } */
export function checkEntry(v, item) {
  const memo = (v.지적내용 || "").trim();
  const photos = v.사진 || [];
  if (!v.결과 && v.점수 == null && !memo && !photos.length) return { state: "empty", problems: [] };
  const problems = [];
  if (!v.결과) problems.push("적합 / 부적합을 선택하세요.");
  if (v.점수 == null) problems.push("점수를 선택하세요.");
  else if (v.점수 < item.scale.min || v.점수 > item.scale.max) problems.push(`점수는 ${item.scale.min}~${item.scale.max} 사이여야 합니다.`);
  if (v.결과 === "부적합" && !memo) problems.push("부적합이면 지적내용을 입력해야 합니다.");
  if (v.결과 === "부적합" && !photos.length) problems.push("부적합이면 사진을 1장 이상 찍어야 합니다.");
  return { state: problems.length ? "draft" : "done", problems };
}

export function currentValue(round, key) {
  const src = round.entries?.[key] || round.drafts?.[key] || {};
  return { 결과: src.결과 ?? null, 점수: src.점수 ?? null, 지적내용: src.지적내용 ?? "", 사진: [...(src.사진 || [])] };
}

export function itemStatus(round, key) {
  const e = round.entries?.[key];
  if (e) return e.결과 === "부적합" ? "bad" : "ok";
  return round.drafts?.[key] ? "draft" : "empty";
}

/** 입력값을 회차에 반영한다 (round 를 직접 바꿈). 점검일·지적일 등 기존 기록은 유지. */
export function applyInput(round, key, v, item, today = localDate()) {
  round.entries ||= {};
  round.drafts ||= {};
  const prev = round.entries[key] || round.drafts[key] || {};
  const check = checkEntry(v, item);
  const meta = {
    점검일: prev.점검일 || today,
    지적일: prev.지적일 ?? null,
    재방문일: prev.재방문일 ?? null,
    완료여부: prev.완료여부 ?? "",
    조치사진: [...(prev.조치사진 || [])], // 재방문 때 찍은 조치 후 사진 (시정 대기 화면에서 관리)
  };
  if (check.state === "empty") {
    delete round.entries[key];
    delete round.drafts[key];
  } else if (check.state === "draft") {
    delete round.entries[key];
    round.drafts[key] = { ...v, 사진: [...(v.사진 || [])], ...meta, updatedAt: new Date().toISOString() };
  } else {
    const bad = v.결과 === "부적합";
    delete round.drafts[key];
    round.entries[key] = {
      결과: v.결과,
      점수: v.점수,
      지적내용: bad ? v.지적내용.trim() : "",
      사진: bad ? [...v.사진] : [],
      점검일: meta.점검일,
      지적일: bad ? meta.지적일 || today : null,
      재방문일: bad ? meta.재방문일 : null,
      완료여부: bad ? meta.완료여부 || "N" : "",
      조치사진: bad ? meta.조치사진 : [],
      updatedAt: new Date().toISOString(),
    };
  }
  return check;
}

/** 회차 모양 맞추기: 저장·작성중 칸이 없으면 만든다 (새 회차, 예전 버전 회차, 백업에서 가져온 회차). 바뀌면 true */
export function normalizeRound(round) {
  let changed = false;
  if (!round.entries || typeof round.entries !== "object") { round.entries = {}; changed = true; }
  if (!round.drafts || typeof round.drafts !== "object") { round.drafts = {}; changed = true; }
  return changed;
}

/**
 * 규칙이 바뀐 뒤(예: 사진 필수 추가) 이미 저장된 항목을 다시 검사한다.
 * 지금 규칙에 맞지 않는 항목은 작성중으로 옮긴다. 바뀐 게 있으면 true.
 */
export function revalidate(round) {
  let changed = false;
  for (const [key, e] of Object.entries(round.entries || {})) {
    const item = round.master.items[Number(key.split("/")[2])];
    if (!item || checkEntry(e, item).state !== "done") {
      round.drafts ||= {};
      round.drafts[key] = { ...e, 사진: [...(e.사진 || [])] };
      delete round.entries[key];
      changed = true;
    }
  }
  return changed;
}
