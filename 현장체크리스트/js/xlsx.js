// 외부 라이브러리 없이 .xlsx를 읽는다.
// .xlsx = zip 압축 안의 XML 파일 묶음. 브라우저 기본 기능(DecompressionStream, DOMParser)만 사용.

const utf8 = new TextDecoder("utf-8");

async function inflateRaw(bytes) {
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

export async function unzip(buffer) {
  const dv = new DataView(buffer);
  let eocd = -1;
  for (let i = buffer.byteLength - 22; i >= Math.max(0, buffer.byteLength - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("엑셀(.xlsx) 파일이 아닙니다.");
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const entries = new Map();
  for (let n = 0; n < count; n++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error("손상된 엑셀 파일입니다.");
    const method = dv.getUint16(p + 10, true);
    const csize = dv.getUint32(p + 20, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commentLen = dv.getUint16(p + 32, true);
    const local = dv.getUint32(p + 42, true);
    const name = utf8.decode(new Uint8Array(buffer, p + 46, nameLen));
    const start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
    entries.set(name, { method, data: new Uint8Array(buffer, start, csize) });
    p += 46 + nameLen + extraLen + commentLen;
  }
  async function bytes(name) {
    const e = entries.get(name);
    if (!e) return null;
    if (e.method === 0) return e.data;
    if (e.method === 8) return inflateRaw(e.data);
    throw new Error(`지원하지 않는 압축 방식(${e.method})입니다.`);
  }
  return {
    names: [...entries.keys()],
    bytes,
    async text(name) {
      const b = await bytes(name);
      return b == null ? null : utf8.decode(b);
    },
  };
}

const xml = (s) => new DOMParser().parseFromString(s, "application/xml");
const els = (node, tag) => [...node.getElementsByTagNameNS("*", tag)];

function colIndex(ref) {
  const letters = ref.match(/^[A-Z]+/)[0];
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function resolveTarget(target) {
  if (target.startsWith("/")) return target.slice(1);
  return "xl/" + target.replace(/^\.\//, "");
}

/** @returns {Promise<Record<string, any[][]>>} 시트 이름 -> 행 배열 (각 행은 셀 값 배열) */
export async function readXlsx(buffer) {
  const zip = await unzip(buffer);
  const wbText = await zip.text("xl/workbook.xml");
  if (!wbText) throw new Error("엑셀(.xlsx) 파일이 아닙니다.");
  const rels = {};
  const relText = await zip.text("xl/_rels/workbook.xml.rels");
  if (relText) for (const r of els(xml(relText), "Relationship")) rels[r.getAttribute("Id")] = r.getAttribute("Target");

  const shared = [];
  const ssText = await zip.text("xl/sharedStrings.xml");
  if (ssText) for (const si of els(xml(ssText), "si")) shared.push(els(si, "t").map((t) => t.textContent).join(""));

  const out = {};
  for (const sh of els(xml(wbText), "sheet")) {
    const name = sh.getAttribute("name");
    const rid = sh.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")
      || sh.getAttribute("r:id");
    const text = rels[rid] && (await zip.text(resolveTarget(rels[rid])));
    if (!text) continue;
    const rows = [];
    for (const row of els(xml(text), "row")) {
      const r = Number(row.getAttribute("r")) - 1;
      const cells = [];
      for (const c of els(row, "c")) {
        const t = c.getAttribute("t");
        const v = els(c, "v")[0]?.textContent;
        let val = null;
        if (t === "s") val = shared[Number(v)] ?? null;
        else if (t === "inlineStr") val = els(c, "t").map((x) => x.textContent).join("");
        else if (t === "str") val = v ?? null;
        else if (t === "b") val = v === "1";
        else if (v != null && v !== "") val = Number(v);
        cells[colIndex(c.getAttribute("r"))] = val;
      }
      rows[r] = cells;
    }
    out[name] = Array.from(rows, (r) => Array.from(r || [], (v) => (v === undefined ? null : v)));
  }
  return out;
}
