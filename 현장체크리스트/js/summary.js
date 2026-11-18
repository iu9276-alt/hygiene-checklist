// 회차 집계 (PRD F6·F7). 값은 원본 입력 그대로 쓰고 추정하지 않는다.
// - 적합/부적합 건수·부적합%: 저장된 항목(entries)만
// - 평균점수: 보고서 도구(make_report.py)와 같은 방식 — 점수가 입력된 모든 항목(작성중 포함, 판정 유무 무관)의 평균
//   (2026-11-18 결정: 평균점수 계산은 보고서 형식으로 전체 통일)
import { entryKey, checkEntry } from "./entry.js";

const round2 = (x) => Math.round(x * 100) / 100;

export function summarize(round) {
  const { stops, items } = round.master;
  const entries = round.entries || {};
  const drafts = round.drafts || {};
  const byStop = [];
  const byItem = items.map((it) => ({ name: it.name, perStop: stops.map(() => 0), total: 0 }));
  const findings = [];
  const pending = []; // 확인 필요: 작성중·미입력

  stops.forEach((stop, si) => {
    const s = { name: stop.name, total: 0, done: 0, ok: 0, bad: 0, draft: 0, empty: 0, scoreSum: 0, scoreN: 0 };
    stop.shops.forEach((shop, sj) => {
      const emptyItems = [];
      items.forEach((item, k) => {
        const key = entryKey(si, sj, k);
        s.total++;
        const e = entries[key];
        if (e) {
          s.done++;
          s.scoreSum += e.점수;
          s.scoreN++;
          if (e.결과 === "부적합") {
            s.bad++;
            byItem[k].perStop[si]++;
            byItem[k].total++;
            findings.push({ key, si, sj, k, stop: stop.name, shop: shop.name, item: item.name, ...e });
          } else s.ok++;
        } else if (drafts[key]) {
          s.draft++;
          const sc = drafts[key].점수;
          if (typeof sc === "number" && sc >= item.scale.min && sc <= item.scale.max) { s.scoreSum += sc; s.scoreN++; }
          pending.push({ type: "draft", key, si, sj, k, stop: stop.name, shop: shop.name, item: item.name,
            problems: checkEntry(drafts[key], item).problems });
        } else {
          s.empty++;
          emptyItems.push(k);
        }
      });
      if (emptyItems.length) pending.push({ type: "empty", si, sj, stop: stop.name, shop: shop.name, items: emptyItems });
    });
    s.avg = s.scoreN ? round2(s.scoreSum / s.scoreN) : null;
    s.badRate = s.done ? s.bad / s.done : null;
    byStop.push(s);
  });

  const tot = byStop.reduce((a, s) => {
    for (const f of ["total", "done", "ok", "bad", "draft", "empty", "scoreSum", "scoreN"]) a[f] += s[f];
    return a;
  }, { total: 0, done: 0, ok: 0, bad: 0, draft: 0, empty: 0, scoreSum: 0, scoreN: 0 });
  tot.avg = tot.scoreN ? round2(tot.scoreSum / tot.scoreN) : null;
  tot.badRate = tot.done ? tot.bad / tot.done : null;

  // 자체 검산: 숫자끼리 맞는지 (안 맞으면 계산 오류이므로 화면에 표시)
  const checks = [
    ["입력완료 + 작성중 + 미입력 = 점검건수", tot.done + tot.draft + tot.empty, tot.total],
    ["적합 + 부적합 = 입력완료", tot.ok + tot.bad, tot.done],
    ["지적사항 목록 = 부적합", findings.length, tot.bad],
    ["항목별 부적합 합계 = 부적합", byItem.reduce((n, x) => n + x.total, 0), tot.bad],
  ].map(([label, a, b]) => ({ label, a, b, ok: a === b }));

  return { byStop, byItem, tot, findings, pending, checks };
}
