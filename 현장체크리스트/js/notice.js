// 업체 시정 요구 문구 초안 (PRD L4)
// 형식·규칙은 기존 skill(hygiene-correction-request)과 같게 맞춘다.
//  - 문구만 만든다. 보내지 않는다 (발송·상신은 담당자).
//  - 시정 기한 = 지적일 + 14일. 기한 안에 담당자가 재방문해 확인해야 완료.
//  - 완료(완료 표시 + 재방문일)는 문구에서 뺀다. 완료 표시인데 재방문일이 없으면 빼고 [확인 필요]로 알린다.
//  - 기한이 지난 항목은 "(기한 경과)"를 붙이고 새 기한은 [확인 필요: 새 기한] — 처리 방법은 담당자가 매번 정한다.
//  - 모르는 값(운영업체명 등)은 추측하지 않고 [확인 필요: …]로 둔다.
import { entryKey } from "./entry.js";
import { addDays, DEADLINE_DAYS } from "./followup.js";

const NEED = (what) => `[확인 필요: ${what}]`;
const isDone = (e) => ["Y", "O", "완료", "예"].includes(String(e.완료여부 || "").trim().toUpperCase());

/**
 * @param {object} round
 * @param {number} si 휴게소 번호
 * @param {{ today: string, operator?: string, photoNames: Map<string,string> }} opt
 */
export function buildNotice(round, si, { today, operator = "", photoNames }) {
  const stop = round.master.stops[si];
  const rows = [];      // 지적사항 수집 표
  const questions = [];
  stop.shops.forEach((shop, sj) => round.master.items.forEach((item, k) => {
    const e = round.entries?.[entryKey(si, sj, k)];
    if (!e || e.결과 !== "부적합") return;
    const d0 = e.지적일 || null;
    const due = d0 ? addDays(d0, DEADLINE_DAYS) : null;
    const text = `[${shop.name}] ${item.name} — ${e.지적내용}`;
    const row = { shop: shop.name, item: item.name, text, d0, due, done: e.완료여부 || "N", 재방문일: e.재방문일 || null,
      photos: (e.사진 || []).map((id) => photoNames.get(id)).filter(Boolean) };
    if (isDone(e) && e.재방문일) {
      row.include = false; row.reason = "완료 (재방문 확인)";
    } else if (isDone(e)) {
      row.include = false; row.reason = "확인 필요: 완료 표시인데 재방문일 없음";
      questions.push(`${text}: 완료로 표시되어 있지만 재방문일이 없습니다. 재방문으로 확인하셨나요? (확인 전이면 미완료로 바꾸면 문구에 들어갑니다)`);
    } else {
      row.include = true;
      row.overdue = !!due && due < today;
      row.reason = row.overdue ? "포함 (기한 경과)" : "포함";
      if (!d0) questions.push(`${text}: 지적일이 없어 시정 기한을 계산할 수 없습니다.`);
      if (row.overdue) {
        questions.push(`${text}: 기한(${due})이 지났습니다. 새 기한을 줄지, 며칠로 할지, 문구 강도를 정해 주세요.`);
      }
    }
    rows.push(row);
  }));

  const inc = rows.filter((r) => r.include);
  const op = operator.trim() || NEED("운영업체명");
  if (!operator.trim()) questions.unshift("운영업체명을 알려 주세요. (위 입력칸에 적으면 문구에 들어갑니다)");
  const dates = [...new Set(inc.map((r) => r.d0).filter(Boolean))].sort();
  const dateText = dates.length ? dates.join(", ") : NEED("점검일");
  const dueText = (r) => (!r.due ? NEED("시정 기한") : r.overdue ? `${r.due} (기한 경과) → ${NEED("새 기한")}` : r.due);
  const photos = inc.flatMap((r) => r.photos);
  const missingPhoto = inc.some((r) => !r.photos.length);
  const nPhoto = missingPhoto && !photos.length ? NEED("사진 장수") : String(photos.length);
  const firstDue = inc.map((r) => (r.overdue ? null : r.due)).filter(Boolean).sort()[0];

  const mail = inc.length ? [
    `제목: [부여지사] ${stop.name} 위생점검 지적사항 시정 요청 (${dateText})`,
    "",
    `${op} 담당자님께,`,
    "",
    "한국도로공사 부여지사 영업파트입니다.",
    `${dateText}에 실시한 ${stop.name} 위생점검에서 아래 사항이 확인되어 시정을 요청드립니다.`,
    "",
    "| 번호 | 지적사항 | 지적일 | 시정 기한 |",
    "|---|---|---|---|",
    ...inc.map((r, i) => `| ${i + 1} | ${r.text} | ${r.d0 || NEED("지적일")} | ${dueText(r)} |`),
    "",
    "- 시정 기한 안에 조치를 마쳐 주시기 바랍니다.",
    "- 기한 안에 담당자가 재방문하여 조치 결과를 확인하며, 재방문 확인이 끝나야 시정 완료로 처리됩니다.",
    "- 조치가 끝나면 조치 내용을 회신해 주시면 재방문 일정을 조율하겠습니다.",
    "",
    `첨부: 지적사항 사진 ${nPhoto}장${photos.length ? ` (${photos.join(", ")})` : ""}`,
    "",
    "부여지사 영업파트 휴게시설 담당 드림",
  ].join("\n") : "";

  const doc = inc.length ? [
    "1. 귀 사의 무궁한 발전을 기원합니다.",
    `2. ${dateText} 실시한 ${stop.name} 위생점검 결과, 아래와 같이 지적사항이 있어 시정을 요청하오니 기한 내 조치하여 주시기 바랍니다.`,
    "",
    "  가. 지적사항",
    ...inc.map((r, i) => `    ${i + 1}) ${r.text} (지적일 ${r.d0 || NEED("지적일")}, 시정 기한 ${dueText(r)})`),
    "  나. 시정 확인: 시정 기한 내 우리 지사 담당자가 재방문하여 확인",
    "  다. 협조사항: 조치 완료 시 조치 결과 회신",
    "",
    `붙임  지적사항 사진 ${nPhoto}부.  끝.`,
  ].join("\n") : "";

  const sms = inc.length
    ? `[부여지사] ${stop.name} 위생점검 지적 ${inc.length}건 시정 요청드립니다. 가장 이른 기한 ${firstDue ? firstDue.slice(5).replace("-", "/") : NEED(inc.some((r) => r.overdue) ? "새 기한" : "기한")}. 상세 내용은 메일(공문)로 안내드리며, 기한 내 재방문하여 확인하겠습니다.`
    : "";

  return { stop: stop.name, rows, included: inc.length, questions, texts: { mail, doc, sms }, photos, firstDue,
    missingPhotoItems: inc.filter((r) => !r.photos.length).map((r) => r.text) };
}
