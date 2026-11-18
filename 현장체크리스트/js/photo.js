// 사진 처리 (PRD F4)
// - 웹앱 안 카메라로 찍은 사진을 긴 변 1600px JPEG로 줄여 웹앱 저장소에만 둔다 (갤러리 보관 안 함, 2026-11-18 결정).
// - 파일명 규칙: 휴게소_점검항목.jpg (띄어쓰기·가운뎃점·파일명에 못 쓰는 글자 제거)
//   같은 휴게소·같은 항목 사진이 두 매장 이상이면 _입점매장, 같은 칸에 여러 장이면 _2, _3 을 붙인다.
//   조치 후 사진은 _조치후 를 붙인다 (예: 가람휴게소_방충방서_조치후.jpg, …_한식당_조치후_2.jpg).

export const MAX_SIDE = 1600;
export const JPEG_QUALITY = 0.85;

export const newId = () =>
  (crypto.randomUUID ? crypto.randomUUID() : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`);

/** 이미지 파일 → { blob, width, height } (JPEG, 긴 변 MAX_SIDE 이하). 사진 방향(EXIF)도 바로잡는다. */
export async function resizeToJpeg(file, maxSide = MAX_SIDE, quality = JPEG_QUALITY) {
  if (!file?.type?.startsWith("image/")) throw new Error("사진 파일이 아닙니다.");
  const bmp = await createImageBitmap(file, { imageOrientation: "from-image" });
  const scale = Math.min(1, maxSide / Math.max(bmp.width, bmp.height));
  const width = Math.round(bmp.width * scale);
  const height = Math.round(bmp.height * scale);
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  canvas.getContext("2d").drawImage(bmp, 0, 0, width, height);
  bmp.close?.();
  const blob = await new Promise((res, rej) =>
    canvas.toBlob((b) => (b ? res(b) : rej(new Error("사진을 줄이지 못했습니다."))), "image/jpeg", quality));
  return { blob, width, height };
}

// 파일명용 정리: 띄어쓰기·가운뎃점·윈도우 파일명에 못 쓰는 글자 제거
export const fileSafe = (s) => String(s ?? "").replace(/[\s·・.\\/:*?"<>|]/g, "");

/**
 * 회차 전체의 사진 파일명을 정한다. 저장된 항목(entries)과 작성중(drafts) 모두 대상.
 * @returns {Map<string, string>} photoId → 파일명
 */
export function assignFileNames(round, photoMeta = new Map()) {
  const all = [];
  for (const bucket of [round.entries || {}, round.drafts || {}]) {
    for (const [key, e] of Object.entries(bucket)) {
      const [si, sj, k] = key.split("/").map(Number);
      const stop = round.master.stops[si];
      const shop = stop?.shops[sj];
      const item = round.master.items[k];
      if (!shop || !item) continue;
      const base = `${fileSafe(stop.name)}_${fileSafe(item.name)}`;
      (e.사진 || []).forEach((id, order) => all.push({
        id, order, shopIdx: sj, base, kind: "", shop: fileSafe(shop.name),
        created: photoMeta.get(id)?.createdAt || "",
      }));
      // 조치 후 사진: 같은 규칙 + "_조치후"
      (e.조치사진 || []).forEach((id, order) => all.push({
        id, order, shopIdx: sj, base, kind: "_조치후", shop: fileSafe(shop.name),
        created: photoMeta.get(id)?.createdAt || "",
      }));
    }
  }
  const byBase = new Map();
  for (const p of all) (byBase.get(p.base) || byBase.set(p.base, []).get(p.base)).push(p);

  const names = new Map();
  for (const [base, list] of byBase) {
    const multiShop = new Set(list.map((p) => p.shopIdx)).size > 1;
    const groups = new Map();
    for (const p of list) {
      const stem = (multiShop ? `${base}_${p.shop}` : base) + p.kind;
      (groups.get(stem) || groups.set(stem, []).get(stem)).push(p);
    }
    for (const [stem, ps] of groups) {
      ps.sort((a, b) => a.shopIdx - b.shopIdx || a.order - b.order || a.created.localeCompare(b.created));
      ps.forEach((p, i) => names.set(p.id, `${stem}${i ? `_${i + 1}` : ""}.jpg`));
    }
  }
  return names;
}
