// 자가 점검 (자동 테스트). 번호는 검증기준.md 와 같다.
// 실제 데이터(IndexedDB)는 쓰지 않는다: 모든 시험은 메모리 안의 가짜 회차로 한다.
import { readXlsx, unzip } from "../js/xlsx.js";
import { validateMaster, parseScale, maskName, maskPhone } from "../js/master.js";
import { checkEntry, applyInput, revalidate, entryKey, normalizeRound, currentValue, itemStatus } from "../js/entry.js";
import { resizeToJpeg, assignFileNames, fileSafe } from "../js/photo.js";
import { summarize } from "../js/summary.js";
import { followupStatus, followupList, checkRevisit } from "../js/followup.js";
import { buildNotice } from "../js/notice.js";
import { buildRows, buildExportZip, COLUMNS } from "../js/export.js";
import { buildXlsx, isoToSerial } from "../js/xlsx-write.js";
import { makeZip, crc32 } from "../js/zip.js";
import { buildBackup, readBackup, savedSinceChange } from "../js/backup.js";

// ---------- 작은 테스트 도구 ----------
const tests = [];
const test = (group, id, name, fn) => tests.push({ group, id, name, fn });
class Fail extends Error {}
const show = (v) => JSON.stringify(v);
const eq = (actual, expected, what = "") => {
  if (show(actual) !== show(expected)) throw new Fail(`${what}\n  기대: ${show(expected)}\n  실제: ${show(actual)}`);
};
const ok = (cond, what) => { if (!cond) throw new Fail(what); };
const has = (arr, part, what = "") => {
  const list = Array.isArray(arr) ? arr : [arr];
  if (!list.some((x) => String(x).includes(part))) throw new Fail(`${what} '${part}' 없음\n  실제: ${show(list)}`);
};
async function rejects(fn, part) {
  try { await fn(); } catch (e) { has([e.message], part, "오류 문구"); return; }
  throw new Fail(`오류가 나야 하는데 나지 않음 ('${part}')`);
}

// ---------- 시험용 데이터 ----------
const sample = async (f) => (await fetch(`../sample/${f}`)).arrayBuffer();
async function jpeg(w, h) {
  const c = document.createElement("canvas");
  c.width = w; c.height = h;
  const g = c.getContext("2d");
  g.fillStyle = "#3d6b8c"; g.fillRect(0, 0, w, h);
  g.fillStyle = "#fff"; g.fillRect(w * 0.1, h * 0.1, w * 0.2, h * 0.1);
  return new Promise((r) => c.toBlob(r, "image/jpeg", 0.8));
}
let MASTER;
async function master() {
  if (!MASTER) MASTER = validateMaster(await readXlsx(await sample("checklist_master_dummy.xlsx"))).master;
  return MASTER;
}
const T = "2026-11-10";
// 가람(0) 한식당(0)·분식코너(1) / 항목 0 조리장 청결, 1 냉장·냉동 온도, 5 방충·방서 / 다솜(2) 편의점(1)
async function fixture() {
  const m = await master();
  const r = { id: "2099-1Q", year: 2099, quarter: 1, name: "시험 회차", createdAt: "2026-11-01T00:00:00Z",
    master: JSON.parse(JSON.stringify(m)), entries: {}, drafts: {} };
  const photos = new Map();
  const addPhoto = async (id, key, created) => photos.set(id, { id, roundId: r.id, key, width: 64, height: 48, createdAt: created, blob: await jpeg(64, 48) });
  await addPhoto("p1", "0/0/1", "2026-11-10T01:00:00Z");
  await addPhoto("p2", "0/0/5", "2026-11-10T02:00:00Z");
  await addPhoto("p3", "0/1/1", "2026-11-10T03:00:00Z");
  const it = r.master.items;
  applyInput(r, "0/0/0", { 결과: "적합", 점수: 5, 지적내용: "", 사진: [] }, it[0], T);
  applyInput(r, "0/0/1", { 결과: "부적합", 점수: 2, 지적내용: "냉장고 8℃", 사진: ["p1"] }, it[1], T);
  r.entries["0/0/1"].재방문일 = "2026-11-16"; r.entries["0/0/1"].완료여부 = "Y";
  applyInput(r, "0/0/5", { 결과: "부적합", 점수: 2, 지적내용: "방충망 파손", 사진: ["p2"] }, it[5], T);
  applyInput(r, "0/1/1", { 결과: "부적합", 점수: 1, 지적내용: "냉동고 성에 (담당 010-3692-4815 연락)", 사진: ["p3"] }, it[1], T);
  applyInput(r, "2/1/5", { 결과: null, 점수: 3, 지적내용: "", 사진: [] }, it[5], "2026-11-12");
  return { r, photos };
}

// ---------- F1 기준정보 ----------
const sheet = (rows) => rows;
const shopsOk = sheet([["휴게소", "입점매장", "매장담당자", "연락처"], ["가람휴게소", "한식당", "김민수", "010-2481-3579"]]);
const itemsOk = sheet([["순서", "점검항목", "점수척도"], [1, "조리장 청결", "1~5"]]);

test("F1 기준정보 가져오기", "F1-1", "정상 더미 파일: 3곳·6매장·8항목, 오류 0", async () => {
  const v = validateMaster(await readXlsx(await sample("checklist_master_dummy.xlsx")));
  eq(v.errors, []); eq(v.master.stops.length, 3); eq(v.master.stops.flatMap((s) => s.shops).length, 6); eq(v.master.items.length, 8);
});
test("F1 기준정보 가져오기", "F1-2", "점수척도 1~5 · 1-3 · 0~10 읽기, 5~1·상중하는 못 읽음", () => {
  eq(parseScale("1~5"), { min: 1, max: 5 }); eq(parseScale("1-3"), { min: 1, max: 3 }); eq(parseScale("0~10"), { min: 0, max: 10 });
  eq(parseScale("5~1"), null); eq(parseScale("상중하"), null);
});
test("F1 기준정보 가져오기", "F1-3", "열 순서가 달라도 열 이름으로 읽음", () => {
  const v = validateMaster({ 매장: [["연락처", "입점매장", "휴게소"], ["010-1111-2222", "카페", "누리휴게소"]], 점검항목: itemsOk });
  eq(v.errors, []); eq(v.master.stops[0].name, "누리휴게소"); eq(v.master.stops[0].shops[0].phone, "010-1111-2222");
});
test("F1 기준정보 가져오기", "F1-4", "연락처 열 없음 → 경고만, 저장 가능", () => {
  const v = validateMaster({ 매장: [["휴게소", "입점매장", "매장담당자"], ["가람휴게소", "한식당", "김민수"]], 점검항목: itemsOk });
  eq(v.errors, []); has(v.warnings, "'연락처' 열이 없습니다");
});
test("F1 기준정보 가져오기", "F1-5", "빈 행은 건너뜀", () => {
  const v = validateMaster({ 매장: [...shopsOk, [null, null, null, null], ["가람휴게소", "분식코너", "", ""]], 점검항목: itemsOk });
  eq(v.errors, []); eq(v.master.stops[0].shops.length, 2);
});
test("F1 기준정보 가져오기", "F1-6", "점검항목 시트 없음 → 오류", () => {
  has(validateMaster({ 매장: shopsOk }).errors, "'점검항목' 시트가 없습니다");
});
test("F1 기준정보 가져오기", "F1-7", "입점매장 빈칸·매장 중복 → 행 번호와 오류 (오류 시험 파일)", async () => {
  const v = validateMaster(await readXlsx(await sample("checklist_master_bad.xlsx")));
  has(v.errors, "5행: 입점매장"); has(v.errors, "6행: 가람휴게소 · 한식당이(가) 중복"); eq(v.master, null);
});
test("F1 기준정보 가져오기", "F1-8", "점수척도 상중하·항목 중복 → 오류", async () => {
  const v = validateMaster(await readXlsx(await sample("checklist_master_bad.xlsx")));
  has(v.errors, "점수척도 '상중하'"); has(v.errors, "'조리장 청결'이(가) 중복");
});
test("F1 기준정보 가져오기", "F1-9", "엑셀이 아닌 파일 → 오류", async () => {
  await rejects(() => readXlsx(new TextEncoder().encode("hello").buffer), "엑셀(.xlsx) 파일이 아닙니다");
});

// ---------- F3·F5 입력 규칙 ----------
const item5 = { name: "x", scale: { min: 1, max: 5 } };
test("F3·F5 입력과 필수값", "F3-1", "적합 + 5점 → 저장", () => {
  const r = { entries: {}, drafts: {} };
  eq(applyInput(r, "0/0/0", { 결과: "적합", 점수: 5, 지적내용: "", 사진: [] }, item5, T).state, "done");
  eq(r.entries["0/0/0"].결과, "적합");
});
test("F3·F5 입력과 필수값", "F3-2", "부적합 + 점수 + 지적내용 + 사진 → 저장, 지적일=오늘, 완료여부 N", () => {
  const r = { entries: {}, drafts: {} };
  applyInput(r, "k", { 결과: "부적합", 점수: 2, 지적내용: " 방충망 파손 ", 사진: ["a"] }, item5, T);
  const e = r.entries.k;
  eq([e.지적내용, e.지적일, e.완료여부, e.사진], ["방충망 파손", T, "N", ["a"]]);
});
test("F3·F5 입력과 필수값", "F3-3", "적합만 (점수 없음) → 작성중", () => {
  const c = checkEntry({ 결과: "적합", 점수: null, 지적내용: "", 사진: [] }, item5);
  eq(c.state, "draft"); has(c.problems, "점수를 선택");
});
test("F3·F5 입력과 필수값", "F3-4", "부적합 + 지적내용 공백만 → 작성중", () => {
  has(checkEntry({ 결과: "부적합", 점수: 2, 지적내용: "   ", 사진: ["a"] }, item5).problems, "지적내용을 입력");
});
test("F3·F5 입력과 필수값", "F3-5", "부적합 + 사진 없음 → 작성중", () => {
  has(checkEntry({ 결과: "부적합", 점수: 2, 지적내용: "x", 사진: [] }, item5).problems, "사진을 1장 이상");
});
test("F3·F5 입력과 필수값", "F3-6", "부적합 → 적합: 지적내용·사진·조치사진 비움", () => {
  const r = { entries: {}, drafts: {} };
  applyInput(r, "k", { 결과: "부적합", 점수: 2, 지적내용: "x", 사진: ["a"] }, item5, T);
  r.entries.k.조치사진 = ["b"];
  applyInput(r, "k", { 결과: "적합", 점수: 5, 지적내용: "x", 사진: ["a"] }, item5, T);
  eq([r.entries.k.지적내용, r.entries.k.사진, r.entries.k.조치사진, r.entries.k.지적일], ["", [], [], null]);
});
test("F3·F5 입력과 필수값", "F3-7", "점수 7 (척도 1~5) → 작성중", () => {
  has(checkEntry({ 결과: "적합", 점수: 7, 지적내용: "", 사진: [] }, item5).problems, "1~5 사이");
});
test("F3·F5 입력과 필수값", "F3-8", "모두 지움 → 미입력", () => {
  const r = { entries: {}, drafts: {} };
  applyInput(r, "k", { 결과: "적합", 점수: null, 지적내용: "", 사진: [] }, item5, T);
  eq(Object.keys(r.drafts), ["k"]);
  eq(applyInput(r, "k", { 결과: null, 점수: null, 지적내용: "", 사진: [] }, item5, T).state, "empty");
  eq([Object.keys(r.entries), Object.keys(r.drafts)], [[], []]);
});
test("F3·F5 입력과 필수값", "F3-9", "규칙 변경 뒤 사진 없는 부적합 저장값 → 작성중으로", () => {
  const r = { master: { items: [item5, item5] }, entries: {
    "0/0/0": { 결과: "적합", 점수: 5, 지적내용: "", 사진: [] },
    "0/0/1": { 결과: "부적합", 점수: 2, 지적내용: "x", 사진: [] } }, drafts: {} };
  eq(revalidate(r), true); eq([Object.keys(r.entries), Object.keys(r.drafts)], [["0/0/0"], ["0/0/1"]]);
});

test("F3·F5 입력과 필수값", "F3-10", "새 회차(작성중 칸 없음)에서 첫 입력이 오류 없이 저장 (2026-11-18 발견 버그 재발 방지)", () => {
  const r = { id: "2099-2Q", entries: {} }; // 예전 버전·백업에서 온 회차 모양
  eq(normalizeRound(r), true);
  eq([typeof r.drafts, normalizeRound(r)], ["object", false]);
  eq(currentValue(r, "0/0/0"), { 결과: null, 점수: null, 지적내용: "", 사진: [] });
  eq(itemStatus(r, "0/0/0"), "empty");
  applyInput(r, "0/0/0", { 결과: "부적합", 점수: null, 지적내용: "", 사진: [] }, item5, T);
  eq(itemStatus(r, "0/0/0"), "draft");
});

// ---------- F4 사진 ----------
test("F4 사진", "F4-1", "4032×3024 → 1600×1200 JPEG", async () => {
  const out = await resizeToJpeg(new File([await jpeg(4032, 3024)], "a.jpg", { type: "image/jpeg" }));
  eq([out.width, out.height, out.blob.type], [1600, 1200, "image/jpeg"]);
});
test("F4 사진", "F4-2", "세로 3024×4032 → 1200×1600", async () => {
  const out = await resizeToJpeg(new File([await jpeg(3024, 4032)], "a.jpg", { type: "image/jpeg" }));
  eq([out.width, out.height], [1200, 1600]);
});
test("F4 사진", "F4-3", "사진이 아닌 파일 → 오류", async () => {
  await rejects(() => resizeToJpeg(new File(["x"], "a.txt", { type: "text/plain" })), "사진 파일이 아닙니다");
});
test("F4 사진", "F4-4", "파일명: 가운뎃점·띄어쓰기·금지문자 제거", async () => {
  eq(fileSafe("냉장·냉동 온도"), "냉장냉동온도"); eq(fileSafe('a/b:c*?"<>|'), "abc");
  const { r, photos } = await fixture();
  eq(assignFileNames(r, photos).get("p2"), "가람휴게소_방충방서.jpg");
});
test("F4 사진", "F4-5", "같은 휴게소·항목이 두 매장 → 매장명 붙임", async () => {
  const { r, photos } = await fixture();
  const n = assignFileNames(r, photos);
  eq([n.get("p1"), n.get("p3")], ["가람휴게소_냉장냉동온도_한식당.jpg", "가람휴게소_냉장냉동온도_분식코너.jpg"]);
});
test("F4 사진", "F4-6", "한 항목 2장 → _2, 조치 후 사진 → _조치후", async () => {
  const { r, photos } = await fixture();
  r.entries["0/0/5"].사진.push("p4"); r.entries["0/0/5"].조치사진 = ["a1"];
  const n = assignFileNames(r, photos);
  eq([n.get("p2"), n.get("p4"), n.get("a1")], ["가람휴게소_방충방서.jpg", "가람휴게소_방충방서_2.jpg", "가람휴게소_방충방서_조치후.jpg"]);
});

// ---------- F7 집계 ----------
test("F6·F7 진행률·집계", "F7-1", "건수 48/적합1/부적합3/작성중1/미입력43, 평균 2.6 (작성중 점수 포함)", async () => {
  const { r } = await fixture();
  const t = summarize(r).tot;
  eq([t.total, t.ok, t.bad, t.draft, t.empty, t.avg], [48, 1, 3, 1, 43, 2.6]);
});
test("F6·F7 진행률·집계", "F7-2", "자체 검산 4가지 모두 일치", async () => {
  const { r } = await fixture();
  eq(summarize(r).checks.map((c) => c.ok), [true, true, true, true]);
});
test("F6·F7 진행률·집계", "F7-3", "항목별 부적합: 냉장·냉동 온도 2, 방충·방서 1 / 확인 필요에 작성중 1건", async () => {
  const { r } = await fixture();
  const S = summarize(r);
  eq([S.byItem[1].total, S.byItem[5].total], [2, 1]);
  eq(S.pending.filter((p) => p.type === "draft").map((p) => p.key), ["2/1/5"]);
});

// ---------- F8 내보내기 ----------
test("F8 내보내기", "F8-1", "48행·14열(조치사진파일 포함), 사진 폴더 사진/20261110/, zip 안 파일 목록", async () => {
  const { r, photos } = await fixture();
  const { blob, fileName, stats } = await buildExportZip(r, photos, { today: "2026-11-18" });
  eq([fileName, stats.rows, stats.photos, COLUMNS.length], ["위생점검_2099_1Q.zip", 48, 3, 14]);
  const z = await unzip(await blob.arrayBuffer());
  has(z.names, "위생점검결과_2099_1Q.xlsx"); has(z.names, "사진/20261110/가람휴게소_냉장냉동온도_한식당.jpg");
  const book = await readXlsx((await z.bytes("위생점검결과_2099_1Q.xlsx")).slice().buffer);
  eq(Object.keys(book), ["가람휴게소", "누리휴게소", "다솜휴게소"]);
  eq(book["가람휴게소"][0], COLUMNS);
});
test("F8 내보내기", "F8-2", "작성중 → 결과 빈칸, 점수 유지, 비고 작성중", async () => {
  const { r, photos } = await fixture();
  const row = buildRows(r, photos).sheets[2].rows[8 + 5];
  eq([row.결과, row.점수], [null, 3]); has([row.비고], "작성중(입력값: 판정 없음)");
});
test("F8 내보내기", "F8-3", "미입력 → 비고 미입력", async () => {
  const { r, photos } = await fixture();
  eq(buildRows(r, photos).sheets[1].rows[0].비고, "미입력");
});
test("F8 내보내기", "F8-4", "가려서 내보내기: 이름·연락처·지적내용 속 번호", async () => {
  const { r, photos } = await fixture();
  const rows = buildRows(r, photos, { mask: true }).sheets[0].rows;
  eq([rows[0].매장담당자, rows[0].연락처], ["김*수", "010-****-3579"]);
  has([rows[8 + 1].지적내용], "010-****-4815");
  eq(maskName("남궁민"), "남*민"); eq(maskPhone("010.1234.5678"), "010-****-5678");
});
test("F8 내보내기", "F8-5", "날짜는 엑셀 날짜 숫자로 (2026-11-10 = 46336)", async () => {
  eq(isoToSerial("2026-11-10"), 46336);
  const x = await buildXlsx([{ name: "s", header: ["점검일", "이름"], rows: [{ 점검일: "2026-11-10", 이름: "가<&>" }], dateCols: ["점검일"] }]);
  const back = await readXlsx(x.slice().buffer);
  eq(back.s, [["점검일", "이름"], [46336, "가<&>"]]);
});
test("F8 내보내기", "F8-6", "zip 압축 확인값(CRC32) 표준과 같음", async () => {
  eq(crc32(new TextEncoder().encode("123456789")).toString(16), "cbf43926");
  const z = await unzip(await makeZip([{ name: "한글/파일.txt", data: "내용" }]).arrayBuffer());
  eq([z.names, await z.text("한글/파일.txt")], [["한글/파일.txt"], "내용"]);
});

// ---------- F9 시정 대기 ----------
const fe = (extra = {}) => ({ 결과: "부적합", 지적일: T, 점검일: T, 완료여부: "N", 재방문일: null, ...extra });
test("F9 시정 대기", "F9-1", "기준일 11-18 → 진행 중(D-6)", () => eq(followupStatus(fe(), "2026-11-18").label, "진행 중(D-6)"));
test("F9 시정 대기", "F9-2", "기준일 11-24 (기한 당일) → 진행 중(D-day)", () => eq(followupStatus(fe(), "2026-11-24").label, "진행 중(D-day)"));
test("F9 시정 대기", "F9-3", "기준일 11-25 → 기한 초과(1일)", () => eq(followupStatus(fe(), "2026-11-25").label, "기한 초과(1일)"));
test("F9 시정 대기", "F9-4", "재방문 11-16 + 완료 → 시정 완료", () => eq(followupStatus(fe({ 재방문일: "2026-11-16", 완료여부: "Y" }), "2026-12-01").state, "done"));
test("F9 시정 대기", "F9-5", "기한 후 재방문 + 완료 → 확인 필요(기한 후 확인)", () =>
  eq(followupStatus(fe({ 재방문일: "2026-11-26", 완료여부: "Y" }), "2026-12-01").label, "확인 필요(기한 후 확인)"));
test("F9 시정 대기", "F9-6", "완료인데 재방문일 없음 → 확인 필요(재방문일 없음)", () =>
  eq(followupStatus(fe({ 완료여부: "Y" }), "2026-11-18").label, "확인 필요(재방문일 없음)"));
test("F9 시정 대기", "F9-7", "재방문일 오류 3가지 막힘, 정상은 통과", () => {
  has(checkRevisit({ 재방문일: "2026-11-09", 완료여부: "Y" }, fe(), "2026-11-18"), "지적일");
  has(checkRevisit({ 재방문일: "2026-11-20", 완료여부: "Y" }, fe(), "2026-11-18"), "오늘보다 늦습니다");
  has(checkRevisit({ 재방문일: "", 완료여부: "Y" }, fe(), "2026-11-18"), "재방문일이 필요");
  eq(checkRevisit({ 재방문일: "2026-11-17", 완료여부: "Y" }, fe(), "2026-11-18"), []);
});
test("F9 시정 대기", "F9-8", "정렬: 기한 초과 → 진행 중 → 완료", async () => {
  const { r } = await fixture();
  eq(followupList(r, "2026-11-25").map((x) => x.st.state), ["overdue", "overdue", "done"]);
  eq(followupList(r, "2026-11-18").map((x) => x.st.state), ["open", "open", "done"]);
});

// ---------- F11·F12 백업 ----------
test("F11 백업·복원", "F11-1", "백업 → 읽기: 회차·사진 그대로 (사진 바이트 동일)", async () => {
  const { r, photos } = await fixture();
  const b = await buildBackup(r, photos);
  const back = await readBackup(await b.blob.arrayBuffer());
  eq([back.round.id, Object.keys(back.round.entries).length, back.photos.length, back.missing], ["2099-1Q", 4, 3, []]);
  const a = new Uint8Array(await photos.get("p2").blob.arrayBuffer());
  const c = new Uint8Array(await back.photos.find((p) => p.id === "p2").blob.arrayBuffer());
  eq([a.length === c.length, a.every((x, i) => x === c[i])], [true, true]);
});
test("F11 백업·복원", "F11-2", "zip이 아닌 파일 → 오류", async () => {
  await rejects(() => readBackup(new TextEncoder().encode("not zip").buffer), "백업 파일(zip)이 아닙니다");
});
test("F11 백업·복원", "F11-3", "엑셀 내보내기 파일 → 오류", async () => {
  const { r, photos } = await fixture();
  const z = await buildExportZip(r, photos);
  await rejects(async () => readBackup(await z.blob.arrayBuffer()), "엑셀 내보내기 파일은 백업으로 가져올 수 없습니다");
});
test("F11 백업·복원", "F11-4", "다른 형식의 backup.json → 오류", async () => {
  const z = makeZip([{ name: "backup.json", data: JSON.stringify({ format: "other", version: 1 }) }]);
  await rejects(async () => readBackup(await z.arrayBuffer()), "이 웹앱의 백업 파일이 아닙니다");
});
test("F11 백업·복원", "F11-5", "사진 하나가 빠진 백업 → 빠진 수 알림", async () => {
  const { r, photos } = await fixture();
  photos.delete("p3");
  const back = await readBackup(await (await buildBackup(r, photos)).blob.arrayBuffer());
  eq(back.missing, ["p3"]);
});
test("F12 삭제·미백업 경고", "F12-1", "변경 뒤 백업함 → 안전", () => {
  eq(savedSinceChange({ changedAt: "2026-11-18T01:00:00Z", backups: [{ at: "2026-11-18T02:00:00Z" }] }).safe, true);
});
test("F12 삭제·미백업 경고", "F12-2", "백업 뒤 다시 변경 / 기록 없음 → 경고", () => {
  eq(savedSinceChange({ changedAt: "2026-11-18T03:00:00Z", backups: [{ at: "2026-11-18T02:00:00Z" }] }).safe, false);
  eq(savedSinceChange({ createdAt: "2026-11-18T00:00:00Z" }).safe, false);
});

// ---------- L4 시정 요구 문구 ----------
const notice = async (today, operator = "", mutate) => {
  const { r, photos } = await fixture();
  if (mutate) mutate(r);
  return buildNotice(r, 0, { today, operator, photoNames: assignFileNames(r, photos) });
};
test("L4 시정 요구 문구", "L4-1", "업체명 입력 → 메일·공문·문자 3종에 반영, 첨부 파일명", async () => {
  const N = await notice("2026-11-18", "가람푸드");
  has([N.texts.mail], "가람푸드 담당자님께"); has([N.texts.mail], "가람휴게소_방충방서.jpg");
  has([N.texts.doc], "붙임  지적사항 사진 2부"); has([N.texts.sms], "지적 2건"); has([N.texts.sms], "11/24");
});
test("L4 시정 요구 문구", "L4-2", "완료(재방문 확인) 항목은 제외", async () => {
  const N = await notice("2026-11-18", "x");
  eq(N.rows.find((x) => x.text.includes("냉장고 8℃")).include, false); eq(N.included, 2);
});
test("L4 시정 요구 문구", "L4-3", "완료 표시인데 재방문일 없음 → 제외 + 확인 질문", async () => {
  const N = await notice("2026-11-18", "x", (r) => { r.entries["0/0/5"].완료여부 = "Y"; });
  eq(N.rows.find((x) => x.text.includes("방충망")).include, false); has(N.questions, "재방문일이 없습니다");
});
test("L4 시정 요구 문구", "L4-4", "기한 경과 → (기한 경과) [확인 필요: 새 기한] + 질문", async () => {
  const N = await notice("2026-11-25", "x");
  has([N.texts.mail], "2026-11-24 (기한 경과) → [확인 필요: 새 기한]"); has(N.questions, "새 기한을 줄지");
  has([N.texts.sms], "[확인 필요: 새 기한]");
});
test("L4 시정 요구 문구", "L4-5", "업체명 비움 → [확인 필요: 운영업체명] + 질문", async () => {
  const N = await notice("2026-11-18", "");
  has([N.texts.mail], "[확인 필요: 운영업체명] 담당자님께"); has(N.questions, "운영업체명");
});

// ---------- L8 조치 후 사진 ----------
test("L8 조치 후 사진", "L8-1", "내보내기: 조치사진파일 열, 재방문일 폴더", async () => {
  const { r, photos } = await fixture();
  photos.set("a1", { id: "a1", key: "0/0/1", width: 64, height: 48, createdAt: "2026-11-16T05:00:00Z", blob: await jpeg(64, 48) });
  r.entries["0/0/1"].조치사진 = ["a1"];
  const { sheets, files } = buildRows(r, photos);
  eq(sheets[0].rows[1].조치사진파일, "가람휴게소_냉장냉동온도_한식당_조치후.jpg");
  has(files.map((f) => f.path), "사진/20261116/가람휴게소_냉장냉동온도_한식당_조치후.jpg");
});
test("L8 조치 후 사진", "L8-2", "부적합 → 적합으로 바꾸면 조치사진 비움 (F3-6과 같은 규칙)", () => {
  const r = { entries: { k: { 결과: "부적합", 점수: 2, 지적내용: "x", 사진: ["a"], 조치사진: ["b"], 지적일: T } }, drafts: {} };
  applyInput(r, "k", { 결과: "적합", 점수: 4, 지적내용: "", 사진: [] }, item5, T);
  eq(r.entries.k.조치사진, []);
});

// ---------- 실행·화면 ----------
const esc = (s) => String(s).replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));
async function runAll() {
  const out = document.getElementById("out");
  const sum = document.getElementById("summary");
  const btn = document.getElementById("run");
  btn.disabled = true; btn.textContent = "실행 중…";
  out.innerHTML = "";
  const results = [];
  const t0 = performance.now();
  let group = null;
  let box;
  for (const t of tests) {
    if (t.group !== group) {
      group = t.group;
      out.insertAdjacentHTML("beforeend", `<section class="t-group"><h2>${esc(group)}</h2><div class="card" style="padding:0 8px"></div></section>`);
      box = out.lastElementChild.querySelector(".card");
    }
    let err = null;
    try { await t.fn(); } catch (e) { err = e instanceof Fail ? e.message : `예상하지 못한 오류: ${e?.stack || e}`; }
    results.push({ id: t.id, name: t.name, pass: !err, err });
    box.insertAdjacentHTML("beforeend", `<div class="t-row ${err ? "fail" : "pass"}"><span class="mark">${err ? "✗" : "✓"}</span>
      <span class="id">${t.id}</span><span>${esc(t.name)}</span>${err ? `<span class="err">${esc(err)}</span>` : ""}</div>`);
  }
  const failed = results.filter((r) => !r.pass);
  sum.hidden = false;
  sum.className = `notice ${failed.length ? "bad" : "ok"}`;
  sum.innerHTML = `<b>${failed.length ? `실패 ${failed.length}건` : "모두 통과"}</b> — 통과 ${results.length - failed.length} / 전체 ${results.length}
    <span class="small">(${Math.round(performance.now() - t0)}ms)</span>`;
  btn.disabled = false; btn.textContent = "다시 실행";
  window.__testResults = { total: results.length, passed: results.length - failed.length, failed };
  return window.__testResults;
}
document.getElementById("run").onclick = runAll;
window.runAll = runAll;
runAll();
