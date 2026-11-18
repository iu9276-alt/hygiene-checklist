// 점검 기준정보(점검기준.xlsx) 검사. 값을 임의로 고치지 않고, 문제는 오류·경고로 돌려준다.
// 오류가 하나라도 있으면 저장할 수 없다.

const SHOP_SHEET = "매장";
const ITEM_SHEET = "점검항목";

const str = (v) => (v == null ? "" : String(v).trim());
const isBlankRow = (row) => !row || row.every((v) => str(v) === "");

function headerIndex(rows) {
  const header = (rows[0] || []).map(str);
  const idx = {};
  header.forEach((h, i) => { if (h) idx[h] = i; });
  return idx;
}

export function parseScale(v) {
  const m = str(v).match(/^(\d+)\s*[~\-–]\s*(\d+)$/);
  if (!m) return null;
  const min = Number(m[1]);
  const max = Number(m[2]);
  return min < max ? { min, max } : null;
}

/**
 * @param {Record<string, any[][]>} sheets readXlsx 결과
 * @returns {{ master: object|null, errors: string[], warnings: string[] }}
 */
export function validateMaster(sheets) {
  const errors = [];
  const warnings = [];

  const shopRows = sheets[SHOP_SHEET];
  const itemRows = sheets[ITEM_SHEET];
  if (!shopRows) errors.push(`'${SHOP_SHEET}' 시트가 없습니다.`);
  if (!itemRows) errors.push(`'${ITEM_SHEET}' 시트가 없습니다.`);
  if (errors.length) return { master: null, errors, warnings };

  // --- 매장 시트 ---
  const si = headerIndex(shopRows);
  for (const col of ["휴게소", "입점매장"]) if (!(col in si)) errors.push(`'${SHOP_SHEET}' 시트에 '${col}' 열이 없습니다.`);
  for (const col of ["매장담당자", "연락처"]) if (!(col in si)) warnings.push(`'${SHOP_SHEET}' 시트에 '${col}' 열이 없습니다. 비워 둔 채로 저장합니다.`);

  const stops = new Map();
  const seen = new Set();
  if (!errors.length) {
    shopRows.slice(1).forEach((row, i) => {
      const line = i + 2;
      if (isBlankRow(row)) return;
      const stop = str(row[si["휴게소"]]);
      const shop = str(row[si["입점매장"]]);
      if (!stop || !shop) {
        errors.push(`'${SHOP_SHEET}' ${line}행: ${!stop ? "휴게소" : "입점매장"}이(가) 비어 있습니다.`);
        return;
      }
      const key = `${stop}\u0000${shop}`;
      if (seen.has(key)) { errors.push(`'${SHOP_SHEET}' ${line}행: ${stop} · ${shop}이(가) 중복입니다.`); return; }
      seen.add(key);
      if (!stops.has(stop)) stops.set(stop, []);
      stops.get(stop).push({
        name: shop,
        person: "매장담당자" in si ? str(row[si["매장담당자"]]) : "",
        phone: "연락처" in si ? str(row[si["연락처"]]) : "",
      });
    });
    if (!seen.size) errors.push(`'${SHOP_SHEET}' 시트에 매장이 하나도 없습니다.`);
  }

  // --- 점검항목 시트 ---
  const ii = headerIndex(itemRows);
  const itemErrStart = errors.length;
  if (!("점검항목" in ii)) errors.push(`'${ITEM_SHEET}' 시트에 '점검항목' 열이 없습니다.`);
  if (!("점수척도" in ii)) errors.push(`'${ITEM_SHEET}' 시트에 '점수척도' 열이 없습니다.`);
  const items = [];
  if (errors.length === itemErrStart) {
    const names = new Set();
    itemRows.slice(1).forEach((row, i) => {
      const line = i + 2;
      if (isBlankRow(row)) return;
      const name = str(row[ii["점검항목"]]);
      if (!name) { errors.push(`'${ITEM_SHEET}' ${line}행: 점검항목이 비어 있습니다.`); return; }
      if (names.has(name)) { errors.push(`'${ITEM_SHEET}' ${line}행: '${name}'이(가) 중복입니다.`); return; }
      names.add(name);
      const scale = parseScale(row[ii["점수척도"]]);
      if (!scale) {
        errors.push(`'${ITEM_SHEET}' ${line}행: 점수척도 '${str(row[ii["점수척도"]]) || "빈칸"}'을(를) 읽을 수 없습니다. (예: 1~5)`);
        return;
      }
      const order = "순서" in ii ? Number(row[ii["순서"]]) : NaN;
      items.push({ name, scale, order: Number.isFinite(order) ? order : null, line });
    });
    if (!names.size) errors.push(`'${ITEM_SHEET}' 시트에 점검항목이 하나도 없습니다.`);
    if (items.some((x) => x.order == null) && "순서" in ii) warnings.push("순서가 비어 있는 점검항목은 엑셀 행 순서대로 둡니다.");
  }
  items.sort((a, b) => (a.order ?? Infinity) - (b.order ?? Infinity) || a.line - b.line);

  if (errors.length) return { master: null, errors, warnings };
  return {
    master: {
      stops: [...stops].map(([name, shops]) => ({ name, shops })),
      items: items.map(({ name, scale }) => ({ name, scale })),
    },
    errors,
    warnings,
  };
}

// 화면 표시용 마스킹 (원본 값은 그대로 저장)
export function maskName(n) {
  n = str(n);
  if (n.length <= 1) return n;
  if (n.length === 2) return n[0] + "*";
  return n[0] + "*".repeat(n.length - 2) + n[n.length - 1];
}
export function maskPhone(p) {
  return str(p).replace(/(01\d)[-. ]?(\d{3,4})[-. ]?(\d{4})/g, (_, a, b, c) => `${a}-${"*".repeat(b.length)}-${c}`);
}
