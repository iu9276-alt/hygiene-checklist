// 위생점검 현장 체크리스트 — 화면과 이동 (PRD 2장 S1·S2·S7)
import { db, requestPersistentStorage } from "./db.js";
import { readXlsx } from "./xlsx.js";
import { validateMaster, maskName, maskPhone } from "./master.js";
import { entryKey, currentValue, itemStatus, checkEntry, applyInput, revalidate, normalizeRound } from "./entry.js";
import { resizeToJpeg, newId, assignFileNames } from "./photo.js";
import { summarize } from "./summary.js";
import { buildRows, buildExportZip, exportBaseName } from "./export.js";
import { followupList, followupStatus, checkRevisit, DEADLINE_DAYS } from "./followup.js";
import { localDate } from "./entry.js";
import { buildBackup, readBackup, savedSinceChange } from "./backup.js";
import { buildNotice } from "./notice.js";
import { capturePhoto } from "./camera.js";
import { makeZip } from "./zip.js";

async function copyText(text) {
  try { await navigator.clipboard.writeText(text); return true; } catch { /* 아래 방식으로 */ }
  const ta = document.createElement("textarea");
  ta.value = text;
  document.body.appendChild(ta);
  ta.select();
  const ok = document.execCommand("copy");
  ta.remove();
  return ok;
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}
let pendingRestore = null; // 가져오기 전 확인 중인 백업 { fileName, data | error }

// 주소 끝 ?d=YYYY-MM-DD 로 기준일을 바꿀 수 있다 (시험용). 없으면 오늘.
let routeParams = new URLSearchParams();
const todayISO = () => (/^\d{4}-\d{2}-\d{2}$/.test(routeParams.get("d") || "") ? routeParams.get("d") : localDate());
const dq = () => (routeParams.get("d") ? `?d=${routeParams.get("d")}` : "");

// 회차를 읽을 때마다 현재 규칙으로 다시 검사 (규칙이 바뀐 뒤 예전 기록이 '저장됨'으로 남지 않게)
async function loadRound(id) {
  const round = await db.getRound(id);
  if (!round) return round;
  normalizeRound(round); // 새 회차·옛 버전 회차·백업에서 가져온 회차 모두 같은 모양으로
  if (revalidate(round)) await db.putRound(round);
  return round;
}

const app = document.getElementById("app");
const titleEl = document.getElementById("title");
const backBtn = document.getElementById("back");

const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const fmtDateTime = (iso) => (iso ? new Date(iso).toLocaleString("ko-KR", { dateStyle: "medium", timeStyle: "short" }) : "");

let pendingImport = null; // 저장 전 미리보기 중인 기준정보 { fileName, result }

function toast(msg) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => { t.hidden = true; }, 2500);
}

function setHeader(title, backHash) {
  titleEl.textContent = title;
  document.title = title === "위생점검 체크리스트" ? title : `${title} · 위생점검`;
  backBtn.hidden = !backHash;
  backBtn.onclick = () => { location.hash = backHash; };
}

// 회차 진행률: 입력한 항목 수 / (매장 수 × 점검항목 수). 입력 기능은 2단계에서 추가.
function progress(round, stopIdx) {
  const stops = stopIdx == null ? round.master.stops : [round.master.stops[stopIdx]];
  const total = stops.reduce((n, s) => n + s.shops.length, 0) * round.master.items.length;
  const prefix = stopIdx == null ? "" : `${stopIdx}/`;
  const done = Object.keys(round.entries || {}).filter((k) => k.startsWith(prefix)).length;
  return { done, total, pct: total ? Math.round((done / total) * 100) : 0 };
}

const bar = (p) => `<div class="progress" aria-hidden="true"><span style="width:${p.pct}%"></span></div>`;

// ---------- S1 회차 선택 (홈) ----------
async function renderHome() {
  setHeader("위생점검 체크리스트", null);
  const [master, rounds] = await Promise.all([db.getMaster(), db.listRounds()]);
  for (const r of rounds) if (revalidate(r)) await db.putRound(r);
  rounds.sort((a, b) => b.id.localeCompare(a.id));

  if (!master) {
    app.innerHTML = `
      <div class="card empty">
        <p><strong>점검 기준정보가 없습니다.</strong></p>
        <p class="muted">휴게소·매장·점검항목이 담긴 엑셀을 가져오면 회차를 만들 수 있습니다.</p>
        <a class="btn block" href="#/settings">기준정보 가져오기</a>
      </div>`;
    return;
  }

  app.innerHTML = `
    <h2>점검 회차</h2>
    ${rounds.length ? rounds.map((r) => {
      const p = progress(r);
      return `<a class="card" href="#/round/${encodeURIComponent(r.id)}">
        <div class="row"><span class="name">${esc(r.name)}</span><span class="badge${p.done === p.total && p.total ? " ok" : ""}">${p.done}/${p.total}</span></div>
        <div class="muted">휴게소 ${r.master.stops.length}곳 · 점검항목 ${r.master.items.length}개</div>
        ${bar(p)}
      </a>`;
    }).join("") : `<div class="card empty muted">아직 만든 회차가 없습니다.</div>`}
    <a class="btn block" href="#/round-new">+ 새 회차 만들기</a>
    <p class="muted" style="margin-top:16px">기준정보: ${esc(master.fileName)} (${fmtDateTime(master.importedAt)} 가져옴)</p>`;
}

// ---------- 새 회차 ----------
async function renderNewRound() {
  setHeader("새 회차 만들기", "#/");
  const master = await db.getMaster();
  if (!master) { location.hash = "#/settings"; return; }
  const now = new Date();
  const y = now.getFullYear();
  const q = Math.floor(now.getMonth() / 3) + 1;
  app.innerHTML = `
    <form id="f" class="card">
      <label class="field"><span>연도</span>
        <select name="year">${[y - 1, y, y + 1].map((v) => `<option ${v === y ? "selected" : ""}>${v}</option>`).join("")}</select>
      </label>
      <label class="field"><span>분기</span>
        <select name="quarter">${[1, 2, 3, 4].map((v) => `<option value="${v}" ${v === q ? "selected" : ""}>${v}분기</option>`).join("")}</select>
      </label>
      <p class="muted">현재 기준정보(휴게소 ${master.stops.length}곳, 점검항목 ${master.items.length}개)를 이 회차에 복사합니다.
      나중에 기준정보를 바꿔도 이미 만든 회차는 바뀌지 않습니다.</p>
      <button class="btn block" type="submit">회차 만들기</button>
    </form>`;
  document.getElementById("f").onsubmit = async (ev) => {
    ev.preventDefault();
    const fd = new FormData(ev.target);
    const year = Number(fd.get("year"));
    const quarter = Number(fd.get("quarter"));
    const id = `${year}-${quarter}Q`;
    if (await db.getRound(id)) { toast(`${year}년 ${quarter}분기 회차가 이미 있습니다.`); return; }
    await db.putRound({
      id, year, quarter, name: `${year}년 ${quarter}분기`,
      createdAt: new Date().toISOString(),
      master: { stops: master.stops, items: master.items, fileName: master.fileName },
      entries: {},
      drafts: {},
    });
    toast("회차를 만들었습니다.");
    location.hash = `#/round/${encodeURIComponent(id)}`;
  };
}

// ---------- S2 휴게소 선택 ----------
async function renderRound(id) {
  const round = await loadRound(id);
  if (!round) { location.hash = "#/"; return; }
  setHeader(round.name, "#/");
  app.innerHTML = `
    <h2>휴게소</h2>
    ${round.master.stops.map((s, i) => {
      const p = progress(round, i);
      return `<a class="card" href="#/round/${encodeURIComponent(id)}/stop/${i}">
        <div class="row"><span class="name">${esc(s.name)}</span><span class="badge${p.done === p.total ? " ok" : ""}">${p.done}/${p.total}</span></div>
        <div class="muted">입점매장 ${s.shops.length}개</div>
        ${bar(p)}
      </a>`;
    }).join("")}
    <div class="stack">
      ${(() => {
        const list = followupList(round, todayISO());
        const pending = list.filter((x) => x.st.state !== "done").length;
        const overdue = list.filter((x) => x.st.state === "overdue").length;
        return list.length ? `<a class="btn ${overdue ? "danger-fill" : "secondary"} block" href="#/round/${encodeURIComponent(id)}/followup${dq()}">
          시정 대기 ${pending}건${overdue ? ` · 기한 초과 ${overdue}건` : ""}</a>` : "";
      })()}
      <a class="btn secondary block" href="#/round/${encodeURIComponent(id)}/review">검토 · 집계 보기</a>
      <a class="btn secondary block" href="#/round/${encodeURIComponent(id)}/export">엑셀 + 사진 내보내기</a>
    </div>
    ${manageSection(round)}`;
  bindManage(round);
}

// ---------- 회차 관리: 백업(F11) · 이 기기에서 삭제(F12) ----------
function manageSection(round) {
  const s = savedSinceChange(round);
  const lastB = (round.backups || []).at(-1);
  return `
    <h2>회차 관리</h2>
    <div class="card stack">
      <p class="small" style="margin:0">
        마지막 변경 ${s.changedAt ? fmtDateTime(s.changedAt) : "-"} ·
        마지막 백업 ${lastB ? fmtDateTime(lastB.at) : "없음"} ·
        마지막 내보내기 ${(round.exports || []).at(-1) ? fmtDateTime(round.exports.at(-1).at) : "없음"}
      </p>
      ${s.safe ? "" : `<p class="small bad-text" style="margin:0">마지막 변경 뒤로 백업·내보내기를 하지 않았습니다.</p>`}
      <button type="button" id="backup" class="btn secondary block">백업 파일 내려받기 (대체 담당자 인계용)</button>
      <p class="muted small" style="margin:0">백업 파일에는 담당자 연락처·사진이 들어 있습니다. 사내 메일 등 사내 경로로만 전달하세요.
        받은 사람은 설정(⚙) → "백업 파일 가져오기"로 이어서 입력합니다.</p>
      <button type="button" id="del-round" class="btn secondary danger block">이 기기에서 회차 삭제</button>
    </div>`;
}

function bindManage(round) {
  document.getElementById("backup").onclick = async (ev) => {
    const btn = ev.currentTarget;
    btn.disabled = true;
    try {
      const photos = new Map((await db.photosByRound(round.id)).map((p) => [p.id, p]));
      const { blob, fileName, photoCount } = await buildBackup(round, photos);
      downloadBlob(blob, fileName);
      round.backups = [...(round.backups || []), { at: new Date().toISOString(), photos: photoCount, bytes: blob.size }];
      await db.putRound(round, false);
      toast(`${fileName}을 내려받았습니다.`);
      route();
    } catch (e) {
      console.error(e);
      toast(`백업하지 못했습니다: ${e.message || e}`);
      btn.disabled = false;
    }
  };
  document.getElementById("del-round").onclick = async () => {
    const photos = await db.photosByRound(round.id);
    const n = Object.keys(round.entries || {}).length + Object.keys(round.drafts || {}).length;
    if (!confirm(`${round.name} 회차를 이 기기에서 삭제합니다.\n입력 ${n}건, 사진 ${photos.length}장이 지워지며 되돌릴 수 없습니다.\n계속할까요?`)) return;
    if (!savedSinceChange(round).safe
      && !confirm("마지막 변경 뒤로 백업·내보내기를 하지 않았습니다.\n지금 삭제하면 그 뒤의 입력은 어디에도 남지 않습니다.\n그래도 삭제할까요?")) return;
    await db.deletePhotos(photos.map((p) => p.id));
    await db.deleteRound(round.id);
    toast(`${round.name} 회차를 이 기기에서 삭제했습니다.`);
    location.hash = "#/";
  };
}

// ---------- 백업 가져오기 확인 ----------
async function renderRestore() {
  if (!pendingRestore) { location.hash = "#/settings"; return; }
  setHeader("백업 가져오기", "#/settings");
  const { fileName, data, error } = pendingRestore;
  if (error) {
    app.innerHTML = `<p class="muted">파일: ${esc(fileName)}</p><div class="notice bad">${esc(error)}</div>
      <a class="btn secondary block" href="#/settings">돌아가기</a>`;
    return;
  }
  const r = data.round;
  const local = await db.getRound(r.id);
  const cnt = (x) => ({ saved: Object.keys(x.entries || {}).length, draft: Object.keys(x.drafts || {}).length });
  const b = cnt(r);
  const l = local && cnt(local);
  app.innerHTML = `
    <p class="muted">파일: ${esc(fileName)}</p>
    ${data.missing.length ? `<div class="notice bad">입력 기록이 가리키는 사진 ${data.missing.length}장이 백업 안에 없습니다. 가져오면 해당 항목은 사진 없이 열립니다.</div>` : ""}
    <div class="card">
      <div class="name">${esc(r.name)}</div>
      <ul class="plain">
        <li>백업 시각: ${fmtDateTime(data.createdAt)}</li>
        <li>입력 저장 ${b.saved}건 · 작성중 ${b.draft}건 · 사진 ${data.photos.length}장</li>
        <li>휴게소 ${r.master.stops.length}곳 · 점검항목 ${r.master.items.length}개</li>
      </ul>
    </div>
    ${local ? `<div class="notice warn"><b>이 기기에 같은 회차가 이미 있습니다.</b>
      (저장 ${l.saved}건 · 작성중 ${l.draft}건 · 마지막 변경 ${fmtDateTime(local.changedAt || local.createdAt)})<br>
      가져오면 이 기기의 ${esc(r.name)} 기록과 사진을 <b>백업 내용으로 바꿉니다.</b> 되돌릴 수 없습니다.</div>`
      : `<div class="notice ok">이 기기에 없는 회차입니다. 새로 추가합니다.</div>`}
    <div class="stack">
      <button type="button" id="restore" class="btn block">${local ? "백업 내용으로 바꾸기" : "가져오기"}</button>
      <a class="btn secondary block" href="#/settings">취소</a>
    </div>`;
  document.getElementById("restore").onclick = async () => {
    if (local && !confirm(`이 기기의 ${r.name} 기록을 백업 내용으로 바꿉니다. 계속할까요?`)) return;
    if (local) await db.deletePhotos((await db.photosByRound(r.id)).map((p) => p.id));
    for (const p of data.photos) await db.putPhoto(p);
    r.restoredAt = new Date().toISOString();
    await db.putRound(r, false);
    pendingRestore = null;
    toast(`${r.name} 회차를 가져왔습니다.`);
    location.hash = `#/round/${encodeURIComponent(r.id)}`;
  };
}

// ---------- S5 시정 대기 ----------
const md = (iso) => (iso ? iso.slice(5) : "-");
const ST_CLASS = { overdue: "bad", nodate: "warn", check: "warn", open: "", done: "ok" };

async function renderFollowup(id) {
  const round = await loadRound(id);
  if (!round) { location.hash = "#/"; return; }
  setHeader(`${round.name} 시정 대기`, `#/round/${encodeURIComponent(id)}${dq()}`);
  const today = todayISO();
  const list = followupList(round, today);
  const groups = [
    ["기한 초과", list.filter((x) => x.st.state === "overdue")],
    ["확인 필요", list.filter((x) => x.st.state === "check" || x.st.state === "nodate")],
    ["진행 중 (기한 가까운 순)", list.filter((x) => x.st.state === "open")],
  ];
  const done = list.filter((x) => x.st.state === "done");
  const card = (x) => `<a class="card pending" href="#/round/${encodeURIComponent(id)}/followup/${x.si}/${x.sj}/${x.k}${dq()}">
      <div class="row"><b>${esc(x.stop)} · ${esc(x.shop)}</b><span class="badge ${ST_CLASS[x.st.state]}">${esc(x.st.label)}</span></div>
      <div class="small">${esc(x.item)}</div>
      <div class="muted small">지적 ${md(x.e.지적일)} → 기한 ${md(x.st.due)}${x.e.재방문일 ? ` · 재방문 ${md(x.e.재방문일)}` : ""}${(x.e.조치사진 || []).length ? ` · 조치 사진 ${x.e.조치사진.length}장` : ""}</div></a>`;

  const pendingByStop = round.master.stops.map((s, si) => list.filter((x) => x.si === si && x.st.state !== "done").length);
  app.innerHTML = `
    ${routeParams.get("d") ? `<div class="notice warn">기준일을 <b>${esc(today)}</b>로 바꿔 보는 중입니다 (시험용).</div>` : ""}
    ${list.length ? `<div class="card">
      <div class="field-label">업체 시정 요구 문구 만들기 (메일·공문·문자)</div>
      <div class="chips-row">${round.master.stops.map((s, si) => `<a class="btn secondary" href="#/round/${encodeURIComponent(id)}/notice/${si}${dq()}">
        ${esc(short(s.name))} ${pendingByStop[si]}건</a>`).join("")}</div>
    </div>` : ""}
    <p class="muted small">시정 기한 = 지적일 + ${DEADLINE_DAYS}일. 기한 안에 재방문해 확인해야 '시정 완료'입니다. 기준일 ${esc(today)}</p>
    ${list.length ? "" : `<div class="card muted">부적합 항목이 없습니다.</div>`}
    ${groups.filter(([, xs]) => xs.length).map(([title, xs]) => `<h2>${title} (${xs.length})</h2>${xs.map(card).join("")}`).join("")}
    ${done.length ? `<details class="done-list"><summary><h2>시정 완료 (${done.length})</h2></summary>${done.map(card).join("")}</details>` : ""}`;
}

async function renderFollowupItem(id, si, sj, k) {
  const round = await loadRound(id);
  const key = entryKey(si, sj, k);
  const e = round?.entries?.[key];
  const listHash = `#/round/${encodeURIComponent(id)}/followup${dq()}`;
  if (!e || e.결과 !== "부적합") { location.hash = listHash; return; }
  const stop = round.master.stops[si];
  const shop = stop.shops[sj];
  const item = round.master.items[k];
  setHeader(`${stop.name} · ${shop.name}`, listHash);
  const today = todayISO();
  const st = followupStatus(e, today);
  let v = { 재방문일: e.재방문일 || "", 완료여부: e.완료여부 === "Y" ? "Y" : "N" };

  for (const u of photoUrls) URL.revokeObjectURL(u);
  photoUrls = [];
  const thumbs = [];
  for (const pid of e.사진 || []) {
    const p = await db.getPhoto(pid);
    if (!p) continue;
    const url = URL.createObjectURL(p.blob);
    photoUrls.push(url);
    thumbs.push(`<img src="${url}" alt="지적 사진">`);
  }

  app.innerHTML = `
    <section class="card">
      <div class="row"><h2 class="item-name" style="margin:0">${esc(item.name)}</h2><span class="badge ${ST_CLASS[st.state]}">${esc(st.label)}</span></div>
      <p style="margin:8px 0">${esc(maskPhone(e.지적내용))}</p>
      ${thumbs.length ? `<div class="photos">${thumbs.map((t) => `<div class="thumb">${t}</div>`).join("")}</div>` : ""}
      <p class="muted small">점수 ${e.점수} · 지적일 ${esc(e.지적일 || "-")} · <b>기한 ${esc(st.due || "-")}</b></p>
    </section>
    <section class="card">
      <h2 style="margin-top:0">재방문 확인</h2>
      <label class="field"><span>재방문일</span><input type="date" id="rv" value="${esc(v.재방문일)}" max="${esc(today)}"></label>
      <button type="button" id="rv-today" class="link-btn">오늘(${esc(today)})로 입력</button>
      <div class="field-label" style="margin-top:12px">시정 결과</div>
      <div class="seg" role="group" aria-label="시정 결과">
        <button type="button" class="seg-btn res-ok" data-v="Y" aria-pressed="${v.완료여부 === "Y"}">시정 완료</button>
        <button type="button" class="seg-btn res-bad" data-v="N" aria-pressed="${v.완료여부 === "N"}">미완료</button>
      </div>
      <p id="fu-status" class="save-status" role="status"></p>
      <button type="button" id="fu-save" class="btn block">저장</button>
      <button type="button" id="fu-clear" class="link-btn" style="margin-top:8px">재방문 기록 지우기</button>
    </section>
    <section class="card">
      <h2 style="margin-top:0">조치 후 사진 <span class="muted small">(선택)</span></h2>
      <div id="after-photos" class="photos"></div>
      <button type="button" id="after-cam" class="btn secondary block">📷 조치 후 사진 촬영</button>
      <p class="muted small">재방문 때 시정된 모습을 찍어 지적 사진과 함께 남깁니다. 찍으면 바로 저장되며, 웹앱 안에만 저장됩니다.</p>
    </section>
    <dialog id="viewer" class="viewer">
      <img id="viewer-img" alt="">
      <p id="viewer-name" class="muted small"></p>
      <div class="viewer-actions">
        <button type="button" id="viewer-del" class="btn secondary danger">사진 삭제</button>
        <button type="button" id="viewer-close" class="btn">닫기</button>
      </div>
    </dialog>`;

  // ----- 조치 후 사진 -----
  e.조치사진 ||= [];
  const afterEl = document.getElementById("after-photos");
  const viewer = document.getElementById("viewer");
  const afterUrls = [];
  async function renderAfter() {
    afterUrls.splice(0).forEach((u) => URL.revokeObjectURL(u));
    const names = assignFileNames(round);
    const cards = [];
    for (const pid of e.조치사진) {
      const p = await db.getPhoto(pid);
      if (!p) continue;
      const url = URL.createObjectURL(p.blob);
      afterUrls.push(url);
      photoUrls.push(url);
      cards.push(`<button type="button" class="thumb" data-id="${esc(pid)}" data-url="${url}" data-name="${esc(names.get(pid) || "")}">
        <img src="${url}" alt="조치 후 사진 ${cards.length + 1}"><span>${esc(names.get(pid) || "")}</span></button>`);
    }
    afterEl.innerHTML = cards.join("") || `<p class="muted small" style="margin:0;grid-column:1/-1">아직 조치 후 사진이 없습니다.</p>`;
    afterEl.querySelectorAll(".thumb").forEach((t) => (t.onclick = () => {
      document.getElementById("viewer-img").src = t.dataset.url;
      document.getElementById("viewer-name").textContent = `저장될 파일명: ${t.dataset.name}`;
      viewer.dataset.id = t.dataset.id;
      viewer.showModal();
    }));
  }
  document.getElementById("viewer-close").onclick = () => viewer.close();
  document.getElementById("viewer-del").onclick = async () => {
    if (!confirm("이 조치 후 사진을 삭제할까요? 되돌릴 수 없습니다.")) return;
    viewer.close();
    const pid = viewer.dataset.id;
    e.조치사진 = e.조치사진.filter((x) => x !== pid);
    await db.putRound(round);
    await db.deletePhotos([pid]);
    await renderAfter();
  };
  const afterBtn = document.getElementById("after-cam");
  afterBtn.onclick = async () => {
    const f = await capturePhoto(`${item.name} · 조치 후 사진`);
    if (!f) return;
    afterBtn.classList.add("busy");
    try {
      const { blob, width, height } = await resizeToJpeg(f);
      const pid = newId();
      await db.putPhoto({ id: pid, roundId: id, key, kind: "조치후", blob, width, height, createdAt: new Date().toISOString() });
      e.조치사진.push(pid);
      await db.putRound(round);
      await renderAfter();
      toast("조치 후 사진을 저장했습니다.");
    } catch (err) {
      toast(err.message || "사진을 저장하지 못했습니다.");
    } finally {
      afterBtn.classList.remove("busy");
    }
  };
  renderAfter();

  const rv = document.getElementById("rv");
  const statusEl = document.getElementById("fu-status");
  const preview = () => {
    const problems = checkRevisit(v, e, today);
    if (problems.length) {
      statusEl.className = "save-status bad";
      statusEl.textContent = problems.join(" ");
    } else {
      const s = followupStatus({ ...e, ...v }, today);
      statusEl.className = `save-status ${s.state === "done" ? "ok" : "muted"}`;
      statusEl.textContent = `저장하면: ${s.label}`;
    }
    return problems;
  };
  rv.onchange = () => { v.재방문일 = rv.value; preview(); };
  document.getElementById("rv-today").onclick = () => { rv.value = today; v.재방문일 = today; preview(); };
  app.querySelectorAll(".seg-btn").forEach((b) => (b.onclick = () => {
    v.완료여부 = b.dataset.v;
    app.querySelectorAll(".seg-btn").forEach((x) => x.setAttribute("aria-pressed", String(x.dataset.v === v.완료여부)));
    preview();
  }));
  preview();

  async function save(next) {
    const problems = checkRevisit(next, e, today);
    if (problems.length) { toast(problems[0]); return; }
    round.entries[key] = { ...e, 조치사진: [...e.조치사진], 재방문일: next.재방문일 || null, 완료여부: next.완료여부, updatedAt: new Date().toISOString() };
    await db.putRound(round);
    toast(`저장했습니다 — ${followupStatus(round.entries[key], today).label}`);
    location.hash = listHash;
  }
  document.getElementById("fu-save").onclick = () => save(v);
  document.getElementById("fu-clear").onclick = async () => {
    if (!confirm("재방문일과 시정 결과를 지울까요?")) return;
    if (e.조치사진.length && confirm(`조치 후 사진 ${e.조치사진.length}장도 지울까요? (취소하면 사진은 남겨 둡니다)`)) {
      const ids = [...e.조치사진];
      e.조치사진 = [];
      await db.putRound(round);
      await db.deletePhotos(ids);
    }
    save({ 재방문일: "", 완료여부: "N" });
  };
}

// ---------- 업체 시정 요구 문구 (L4) ----------
async function renderNotice(id, si) {
  const round = await loadRound(id);
  const stop = round?.master.stops[si];
  if (!stop) { location.hash = "#/"; return; }
  const followHash = `#/round/${encodeURIComponent(id)}/followup${dq()}`;
  setHeader(`${stop.name} 시정 요구`, followHash);
  const today = todayISO();
  const photoList = await db.photosByRound(id);
  const photoMap = new Map(photoList.map((p) => [p.id, p]));
  const photoNames = assignFileNames(round, photoMap);
  round.operators ||= {};
  const make = () => buildNotice(round, si, { today, operator: round.operators[stop.name] || "", photoNames });
  let N = make();
  const KINDS = [["mail", "A. 메일 (사내 메일·Gmail 공용)"], ["doc", "B. 공문 본문 (결재 시스템용)"], ["sms", "C. 문자·메신저"]];

  app.innerHTML = `
    ${routeParams.get("d") ? `<div class="notice warn">기준일 <b>${esc(today)}</b> (시험용)</div>` : ""}
    <div class="chips-row" style="margin-bottom:12px">${round.master.stops.map((s, i) => `<a class="btn ${i === si ? "" : "secondary"}"
      href="#/round/${encodeURIComponent(id)}/notice/${i}${dq()}">${esc(short(s.name))}</a>`).join("")}</div>
    <div class="notice warn">문구만 만듭니다. <b>보내지 않습니다.</b> 받는 사람 확인, 사진 첨부, 발송·상신은 직접 해 주세요.</div>
    <label class="field"><span>운영업체명 (모르면 비워 두세요 — 문구에 [확인 필요]로 남습니다)</span>
      <input type="text" id="op" value="${esc(round.operators[stop.name] || "")}" placeholder="예) ○○푸드">
    </label>

    <h2>지적사항 수집 (${N.rows.length}건 중 문구 포함 ${N.included}건)</h2>
    <div class="table-wrap"><table>
      <tr><th>#</th><th>지적사항</th><th>지적일</th><th>기한</th><th>완료</th><th>포함</th></tr>
      ${N.rows.map((r, i) => `<tr><td>${i + 1}</td><td>${esc(r.text)}</td><td>${esc(r.d0 || "-")}</td><td>${esc(r.due || "-")}</td>
        <td>${esc(r.done)}</td><td class="${r.include ? (r.overdue ? "bad-text" : "") : "muted"}">${esc(r.reason)}</td></tr>`).join("")
        || `<tr><td colspan="6" class="muted">이 휴게소에는 부적합 항목이 없습니다.</td></tr>`}
    </table></div>

    <h2>확인 부탁드릴 사항</h2>
    <ol id="questions" class="plain"></ol>

    ${N.included ? `
      <p class="muted small">지적내용은 입력한 그대로 들어갑니다. 업체에 보내면 안 되는 내부 메모(연락 기록 등)가 있으면 아래 칸에서 지우고 복사하세요.
        운영업체명을 바꾸면 아래 문구가 다시 만들어져 직접 고친 내용은 사라집니다.</p>
      ${KINDS.map(([k, label]) => `
        <div class="card notice-text">
          <div class="row"><b>${label}</b><button type="button" class="btn secondary" data-copy="${k}">복사</button></div>
          <textarea id="t-${k}" rows="${k === "sms" ? 3 : 12}"></textarea>
        </div>`).join("")}
      <button type="button" id="photos-zip" class="btn secondary block" ${N.photos.length ? "" : "disabled"}>
        첨부할 사진 ${N.photos.length}장 zip 받기</button>
      ${N.missingPhotoItems.length ? `<p class="small bad-text">사진이 없는 지적사항: ${N.missingPhotoItems.map(esc).join(" / ")}</p>` : ""}`
      : `<div class="card muted">문구에 넣을 미완료 지적사항이 없습니다.</div>`}`;

  const fill = () => {
    document.getElementById("questions").innerHTML = N.questions.length
      ? N.questions.map((q) => `<li>${esc(q)}</li>`).join("") : `<li class="muted">없습니다.</li>`;
    for (const [k] of KINDS) { const t = document.getElementById(`t-${k}`); if (t) t.value = N.texts[k]; }
  };
  fill();

  let timer;
  document.getElementById("op").oninput = (ev) => {
    round.operators[stop.name] = ev.target.value;
    clearTimeout(timer);
    timer = setTimeout(async () => { N = make(); fill(); await db.putRound(round, false); }, 300);
  };
  app.querySelectorAll("[data-copy]").forEach((b) => (b.onclick = async () => {
    const ok = await copyText(document.getElementById(`t-${b.dataset.copy}`).value);
    toast(ok ? "복사했습니다. 메일·결재·문자 창에 붙여 넣으세요." : "복사하지 못했습니다. 칸을 길게 눌러 직접 복사하세요.");
  }));
  const zipBtn = document.getElementById("photos-zip");
  if (zipBtn) zipBtn.onclick = async () => {
    const files = [];
    for (const r of N.rows.filter((x) => x.include)) {
      for (const name of r.photos) {
        const id2 = [...photoNames].find(([, n]) => n === name)?.[0];
        const p = id2 && photoMap.get(id2);
        if (p) files.push({ name, data: new Uint8Array(await p.blob.arrayBuffer()) });
      }
    }
    const ymd = today.replace(/-/g, "");
    downloadBlob(makeZip(files), `시정요구_사진_${stop.name}_${ymd}.zip`);
    toast(`사진 ${files.length}장을 내려받았습니다.`);
  };
}

// ---------- S6 내보내기 ----------
async function renderExport(id) {
  const round = await loadRound(id);
  if (!round) { location.hash = "#/"; return; }
  setHeader(`${round.name} 내보내기`, `#/round/${encodeURIComponent(id)}`);
  const photoList = await db.photosByRound(id);
  const photos = new Map(photoList.map((p) => [p.id, p]));
  const { stats } = buildRows(round, photos);
  const S = summarize(round);
  const base = exportBaseName(round);
  const last = (round.exports || []).at(-1);

  app.innerHTML = `
    ${stats.draft || stats.empty ? `<div class="notice warn"><b>아직 끝나지 않은 항목이 있습니다.</b>
      작성중 ${stats.draft}건 · 미입력 ${stats.empty}건은 결과 칸을 비운 채 내보내며, 보고서에는 <b>판정누락</b>으로 표시됩니다.
      <a href="#/round/${encodeURIComponent(id)}/review">검토 화면에서 확인</a></div>`
      : `<div class="notice ok"><b>모든 항목 입력 완료</b> — 그대로 내보낼 수 있습니다.</div>`}
    ${stats.missingPhotos.length ? `<div class="notice bad">사진 ${stats.missingPhotos.length}장을 저장소에서 찾지 못해 빠집니다.</div>` : ""}

    <div class="card">
      <div class="name">위생점검_${esc(base)}.zip</div>
      <ul class="plain">
        <li>위생점검결과_${esc(base)}.xlsx — 휴게소 ${round.master.stops.length}개 시트, ${stats.rows}행
          <span class="muted small">(저장 ${stats.saved} · 작성중 ${stats.draft} · 미입력 ${stats.empty})</span></li>
        <li>사진/날짜별 폴더 — ${stats.photos}장</li>
      </ul>
      <p class="muted small">적합 ${S.tot.ok} · 부적합 ${S.tot.bad} · 평균점수 ${S.tot.avg ?? "-"} (보고서와 같은 계산)</p>
      <label class="check"><input type="checkbox" id="mask"> 담당자 이름·연락처를 가려서 내보내기</label>
      <p class="muted small">끄면 원본 그대로 내보내고, 한글 보고서를 만들 때 보고서 도구가 가립니다.</p>
      <button type="button" id="go" class="btn block">zip 내려받기</button>
    </div>
    <p class="muted small">압축을 <b>D:\\영업파트\\위생점검\\${round.year}\\${round.quarter}분기\\</b>에 풀면 기존 폴더 구조와 같습니다.
      그 폴더로 보고서 도구(make_report.py)를 실행하면 한글 보고서 초안이 만들어집니다.</p>
    ${last ? `<p class="muted small">마지막 내보내기: ${fmtDateTime(last.at)} (${last.rows}행 · 사진 ${last.photos}장)</p>` : ""}`;

  document.getElementById("go").onclick = async (ev) => {
    const btn = ev.currentTarget;
    btn.disabled = true;
    btn.textContent = "만드는 중…";
    try {
      const { blob, fileName, stats: st } = await buildExportZip(round, photos, { mask: document.getElementById("mask").checked });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 60000);
      round.exports = [...(round.exports || []), { at: new Date().toISOString(), rows: st.rows, photos: st.photos, bytes: blob.size }];
      await db.putRound(round, false);
      toast(`${fileName} (${Math.round(blob.size / 1024)}KB)을 내려받았습니다.`);
      renderExport(id);
    } catch (e) {
      console.error(e);
      toast(`내보내지 못했습니다: ${e.message || e}`);
      btn.disabled = false;
      btn.textContent = "zip 내려받기";
    }
  };
}

// ---------- S4 검토 · 집계 ----------
const pct = (x) => (x == null ? "-" : `${(x * 100).toFixed(1)}%`);
const short = (n) => String(n).replace(/휴게소$/, ""); // 표 머리글용 (가람휴게소 → 가람)
const num = (x) => (x == null ? "-" : x.toFixed(2));

async function renderReview(id) {
  const round = await loadRound(id);
  if (!round) { location.hash = "#/"; return; }
  setHeader(`${round.name} 검토`, `#/round/${encodeURIComponent(id)}`);
  const S = summarize(round);
  const { tot, byStop, byItem, findings, pending, checks } = S;
  const enc = encodeURIComponent(id);
  const itemHref = (si, sj, k) => `#/round/${enc}/stop/${si}/shop/${sj}/${k}`;
  const items = round.master.items;
  const allOk = checks.every((c) => c.ok);
  const finished = tot.draft === 0 && tot.empty === 0;

  app.innerHTML = `
    <div class="notice ${finished ? "ok" : "warn"}">
      ${finished ? `<b>모든 항목 입력 완료</b> (${tot.total}건)` : `<b>입력 ${tot.done}/${tot.total}건</b> · 작성중 ${tot.draft}건 · 미입력 ${tot.empty}건`}
    </div>
    ${allOk ? "" : `<div class="notice bad"><b>집계 검산 불일치</b> — 계산 오류입니다. 개발 담당에게 알려주세요.<ul>${checks.filter((c) => !c.ok).map((c) => `<li>${esc(c.label)}: ${c.a} ≠ ${c.b}</li>`).join("")}</ul></div>`}

    <h2>휴게소별 집계</h2>
    <div class="table-wrap"><table class="num">
      <tr><th>휴게소</th><th>점검</th><th>적합</th><th>부적합</th><th>작성중</th><th>미입력</th><th>부적합%</th><th>평균</th></tr>
      ${byStop.map((s) => `<tr><td title="${esc(s.name)}">${esc(short(s.name))}</td><td>${s.total}</td><td>${s.ok}</td><td>${s.bad}</td><td>${s.draft}</td><td>${s.empty}</td><td>${pct(s.badRate)}</td><td>${num(s.avg)}</td></tr>`).join("")}
      <tr class="total"><td>합계</td><td>${tot.total}</td><td>${tot.ok}</td><td>${tot.bad}</td><td>${tot.draft}</td><td>${tot.empty}</td><td>${pct(tot.badRate)}</td><td>${num(tot.avg)}</td></tr>
    </table></div>
    <p class="muted small">부적합%는 입력완료 항목만으로, 평균점수는 보고서와 같은 방식으로 점수가 입력된 모든 항목(작성중 포함)으로 계산합니다.</p>

    <h2>항목별 부적합</h2>
    <div class="table-wrap"><table class="num">
      <tr><th>점검항목</th>${byStop.map((s) => `<th title="${esc(s.name)}">${esc(short(s.name))}</th>`).join("")}<th>합계</th></tr>
      ${byItem.map((it) => `<tr><td>${esc(it.name)}</td>${it.perStop.map((n) => `<td>${n || ""}</td>`).join("")}<td>${it.total || ""}</td></tr>`).join("")}
      <tr class="total"><td>합계</td>${byStop.map((s) => `<td>${s.bad}</td>`).join("")}<td>${tot.bad}</td></tr>
    </table></div>

    <h2>확인 필요 (${pending.reduce((n, p) => n + (p.type === "draft" ? 1 : p.items.length), 0)}건)</h2>
    ${pending.length ? pending.map((p) => p.type === "draft"
      ? `<a class="card pending draft" href="${itemHref(p.si, p.sj, p.k)}">
          <div class="row"><span><b>${esc(p.stop)} · ${esc(p.shop)}</b> — ${p.k + 1}. ${esc(p.item)}</span><span class="badge warn">작성중</span></div>
          <div class="small bad-text">${p.problems.map(esc).join(" ")}</div></a>`
      : `<a class="card pending" href="${itemHref(p.si, p.sj, p.items[0])}">
          <div class="row"><span><b>${esc(p.stop)} · ${esc(p.shop)}</b></span><span class="badge">미입력 ${p.items.length}</span></div>
          <div class="small muted">${p.items.length === items.length ? "아직 시작 안 함" : p.items.map((k) => `${k + 1}. ${esc(items[k].name)}`).join(", ")}</div></a>`).join("")
    : `<div class="card muted">없습니다.</div>`}

    <h2>지적사항 (${findings.length}건)</h2>
    ${findings.length ? findings.map((f) => `<a class="card finding" href="${itemHref(f.si, f.sj, f.k)}">
        <div class="row"><b>${esc(f.stop)} · ${esc(f.shop)}</b><span class="badge bad">${f.점수}점 · 사진 ${(f.사진 || []).length}</span></div>
        <div class="small">${esc(f.item)}</div>
        <div class="muted small">${esc(maskPhone(f.지적내용))}</div></a>`).join("")
    : `<div class="card muted">부적합 항목이 없습니다.</div>`}
    <p class="muted small">지적내용 속 전화번호는 화면에서 가려서 표시합니다.</p>`;
}

// ---------- S2 매장 선택 ----------
async function renderStop(id, si) {
  const round = await loadRound(id);
  const stop = round?.master.stops[si];
  if (!stop) { location.hash = "#/"; return; }
  setHeader(stop.name, `#/round/${encodeURIComponent(id)}`);
  const n = round.master.items.length;
  app.innerHTML = `
    <h2>입점매장</h2>
    ${stop.shops.map((sh, j) => {
      const done = Object.keys(round.entries || {}).filter((k) => k.startsWith(`${si}/${j}/`)).length;
      const drafts = Object.keys(round.drafts || {}).filter((k) => k.startsWith(`${si}/${j}/`)).length;
      return `<a class="card" href="#/round/${encodeURIComponent(id)}/stop/${si}/shop/${j}">
        <div class="row"><span class="name">${esc(sh.name)}</span>
          <span>${drafts ? `<span class="badge warn">작성중 ${drafts}</span> ` : ""}<span class="badge${done === n ? " ok" : ""}">${done}/${n}</span></span></div>
        <div class="muted">담당 ${esc(maskName(sh.person) || "-")} · ${esc(maskPhone(sh.phone) || "-")}</div>
      </a>`;
    }).join("")}`;
}

// ---------- S3 체크리스트 입력 ----------
// 저장은 순서대로 한 번에 하나씩 (빠르게 눌러도 나중 값이 먼저 저장되지 않게)
let saveChain = Promise.resolve();
let pendingCommit = null; // 지적내용 입력 중 아직 반영 안 된 저장
const saveRound = (round) => (saveChain = saveChain.then(() => db.putRound(round)));
async function flushPending() {
  if (pendingCommit) { const f = pendingCommit; pendingCommit = null; f(); }
  await saveChain;
}

const STATUS_LABEL = { ok: "적합", bad: "부적합", draft: "작성중", empty: "미입력" };
let photoUrls = []; // 화면에 띄운 사진 주소 (다시 그릴 때 해제)

// 지적내용은 글자마다 즉시 임시 버퍼(localStorage)에도 적어 둔다.
// 저장(IndexedDB)이 끝나기 전에 앱이 꺼져도, 다음 실행 때 recoverMemoBuffer()가 되살린다.
const MEMO_BUF = "memo-buffer";
function writeMemoBuffer(id, key, text) {
  try { localStorage.setItem(MEMO_BUF, JSON.stringify({ id, key, text })); } catch { /* 저장 불가 환경이면 무시 */ }
}
function clearMemoBuffer(id, key, text) {
  try {
    const b = JSON.parse(localStorage.getItem(MEMO_BUF) || "null");
    if (b && b.id === id && b.key === key && b.text === text) localStorage.removeItem(MEMO_BUF);
  } catch { /* 무시 */ }
}
async function recoverMemoBuffer() {
  let b;
  try { b = JSON.parse(localStorage.getItem(MEMO_BUF) || "null"); } catch { return; }
  if (!b) return;
  const round = await db.getRound(b.id);
  const [si, sj, k] = String(b.key).split("/").map(Number);
  const item = round?.master.items[k];
  if (round && item && round.master.stops[si]?.shops[sj]) {
    const v = currentValue(round, b.key);
    if (v.지적내용 !== b.text) {
      applyInput(round, b.key, { ...v, 지적내용: b.text }, item);
      await db.putRound(round);
    }
  }
  try { localStorage.removeItem(MEMO_BUF); } catch { /* 무시 */ }
}

async function renderShop(id, si, sj, k = 0) {
  const round = await loadRound(id);
  const stop = round?.master.stops[si];
  const shop = stop?.shops[sj];
  const items = round?.master.items || [];
  if (!shop || !items[k]) { location.hash = "#/"; return; }
  const base = `#/round/${encodeURIComponent(id)}/stop/${si}/shop/${sj}`;
  const stopHash = `#/round/${encodeURIComponent(id)}/stop/${si}`;
  setHeader(`${stop.name} · ${shop.name}`, stopHash);

  const item = items[k];
  const key = entryKey(si, sj, k);
  let v = currentValue(round, key);
  const last = k === items.length - 1;

  const scores = [];
  for (let s = item.scale.min; s <= item.scale.max; s++) scores.push(s);

  app.innerHTML = `
    <nav class="chips" aria-label="점검항목 바로가기">
      ${items.map((it, i) => {
        const st = itemStatus(round, entryKey(si, sj, i));
        return `<button type="button" class="chip st-${st}${i === k ? " current" : ""}" data-k="${i}"
          title="${esc(it.name)} · ${STATUS_LABEL[st]}" aria-label="${i + 1}번 ${esc(it.name)}, ${STATUS_LABEL[st]}">${i + 1}</button>`;
      }).join("")}
    </nav>
    <section class="card item-card">
      <div class="muted">${k + 1} / ${items.length}</div>
      <h2 class="item-name">${esc(item.name)}</h2>
      <div class="seg" role="group" aria-label="판정">
        <button type="button" class="seg-btn res-ok" data-res="적합" aria-pressed="${v.결과 === "적합"}">적합</button>
        <button type="button" class="seg-btn res-bad" data-res="부적합" aria-pressed="${v.결과 === "부적합"}">부적합</button>
      </div>
      <div class="field-label">점수 (${item.scale.min}~${item.scale.max})</div>
      <div class="scores" role="group" aria-label="점수">
        ${scores.map((s) => `<button type="button" class="score" data-score="${s}" aria-pressed="${v.점수 === s}">${s}</button>`).join("")}
      </div>
      <div id="memo-wrap" ${v.결과 === "부적합" ? "" : "hidden"}>
        <label class="field"><span>지적내용 <b class="req">필수</b></span>
          <textarea id="memo" rows="3" placeholder="예) 냉장고 내부 온도 8℃, 기준(5℃ 이하) 초과">${esc(v.지적내용)}</textarea>
        </label>
        <div class="field-label">사진 <b class="req">필수</b></div>
        <div id="photos" class="photos"></div>
        <button type="button" class="btn secondary block" id="cam-btn">📷 사진 촬영</button>
        <p class="muted small">웹앱 안의 카메라로 찍습니다. 사진은 휴대폰 갤러리에 남지 않고 이 웹앱 안에만 저장됩니다.</p>
      </div>
      <p id="status" class="save-status" role="status"></p>
      <button type="button" id="clear" class="link-btn">이 항목 입력 지우기</button>
    </section>
    <div class="navbar">
      <button type="button" id="prev" class="btn secondary" ${k === 0 ? "disabled" : ""}>← 이전</button>
      <button type="button" id="next" class="btn">${last ? "완료 · 매장 목록" : "다음 →"}</button>
    </div>
    <dialog id="viewer" class="viewer">
      <img id="viewer-img" alt="">
      <p id="viewer-name" class="muted small"></p>
      <div class="viewer-actions">
        <button type="button" id="viewer-del" class="btn secondary danger">사진 삭제</button>
        <button type="button" id="viewer-close" class="btn">닫기</button>
      </div>
    </dialog>`;

  const statusEl = document.getElementById("status");
  const memo = document.getElementById("memo");

  function showStatus(check) {
    if (check.state === "done") {
      statusEl.className = "save-status ok";
      statusEl.textContent = `✓ 저장됨 (${v.결과} · ${v.점수}점)`;
    } else if (check.state === "draft") {
      statusEl.className = "save-status bad";
      statusEl.innerHTML = `저장 안 됨 — ${check.problems.map(esc).join(" ")}`;
    } else {
      statusEl.className = "save-status muted";
      statusEl.textContent = "아직 입력하지 않았습니다.";
    }
    const chip = app.querySelector(`.chip[data-k="${k}"]`);
    const st = itemStatus(round, key);
    chip.className = `chip st-${st} current`;
    chip.setAttribute("aria-label", `${k + 1}번 ${item.name}, ${STATUS_LABEL[st]}`);
  }

  function commit() {
    const check = applyInput(round, key, v, item);
    const text = v.지적내용;
    saveRound(round).then(() => clearMemoBuffer(id, key, text));
    showStatus(check);
    return check;
  }

  showStatus(checkEntry(v, item));

  // ----- 사진 -----
  const photosEl = document.getElementById("photos");
  const viewer = document.getElementById("viewer");
  async function renderPhotos() {
    for (const u of photoUrls) URL.revokeObjectURL(u);
    photoUrls = [];
    const names = assignFileNames(round);
    const cards = [];
    for (const pid of v.사진) {
      const p = await db.getPhoto(pid);
      if (!p) continue;
      const url = URL.createObjectURL(p.blob);
      photoUrls.push(url);
      cards.push(`<button type="button" class="thumb" data-id="${esc(pid)}" data-url="${url}" data-name="${esc(names.get(pid) || "")}">
        <img src="${url}" alt="사진 ${cards.length + 1}"><span>${esc(names.get(pid) || "")}</span></button>`);
    }
    photosEl.innerHTML = cards.join("") || `<p class="muted small" style="margin:0;grid-column:1/-1">아직 사진이 없습니다.</p>`;
    photosEl.querySelectorAll(".thumb").forEach((t) => (t.onclick = () => {
      document.getElementById("viewer-img").src = t.dataset.url;
      document.getElementById("viewer-name").textContent = `저장될 파일명: ${t.dataset.name}`;
      viewer.dataset.id = t.dataset.id;
      viewer.showModal();
    }));
  }
  async function removePhotos(ids) {
    if (!ids.length) return;
    v.사진 = v.사진.filter((x) => !ids.includes(x));
    commit();
    await saveChain;
    await db.deletePhotos(ids);
  }
  document.getElementById("viewer-close").onclick = () => viewer.close();
  document.getElementById("viewer-del").onclick = async () => {
    if (!confirm("이 사진을 삭제할까요? 되돌릴 수 없습니다.")) return;
    viewer.close();
    await removePhotos([viewer.dataset.id]);
    await renderPhotos();
  };
  const camBtn = document.getElementById("cam-btn");
  camBtn.onclick = async () => {
    const f = await capturePhoto(`${item.name} · 지적 사진`);
    if (f) await savePhoto(f);
  };
  async function savePhoto(file) {
    camBtn.classList.add("busy");
    try {
      const { blob, width, height } = await resizeToJpeg(file);
      const pid = newId();
      await db.putPhoto({ id: pid, roundId: id, key, blob, width, height, createdAt: new Date().toISOString() });
      v.사진.push(pid);
      commit();
      await saveChain;
      await renderPhotos();
      toast("사진을 저장했습니다.");
    } catch (e) {
      toast(e.message || "사진을 저장하지 못했습니다.");
    } finally {
      camBtn.classList.remove("busy");
    }
  };
  renderPhotos();

  app.querySelectorAll(".seg-btn").forEach((b) => (b.onclick = async () => {
    const next = v.결과 === b.dataset.res ? null : b.dataset.res; // 다시 누르면 선택 해제
    const after = (round.entries?.[key] || round.drafts?.[key] || {}).조치사진 || [];
    if (next !== "부적합" && (v.사진.length || after.length)) {
      const what = [v.사진.length ? `지적 사진 ${v.사진.length}장` : "", after.length ? `조치 후 사진 ${after.length}장` : ""].filter(Boolean).join(", ");
      if (!confirm(`부적합이 아니면 사진이 필요 없습니다. 이 항목의 ${what}을 삭제할까요?`)) return;
      await removePhotos([...v.사진]);
      if (after.length) {
        for (const bucket of [round.entries, round.drafts]) if (bucket?.[key]) bucket[key].조치사진 = [];
        await saveRound(round);
        await db.deletePhotos(after);
      }
      await renderPhotos();
    }
    v.결과 = next;
    app.querySelectorAll(".seg-btn").forEach((x) => x.setAttribute("aria-pressed", String(x.dataset.res === v.결과)));
    document.getElementById("memo-wrap").hidden = v.결과 !== "부적합";
    commit();
    if (v.결과 === "부적합" && !v.지적내용.trim()) memo.focus();
  }));
  app.querySelectorAll(".score").forEach((b) => (b.onclick = () => {
    const s = Number(b.dataset.score);
    v.점수 = v.점수 === s ? null : s;
    app.querySelectorAll(".score").forEach((x) => x.setAttribute("aria-pressed", String(Number(x.dataset.score) === v.점수)));
    commit();
  }));
  let timer;
  memo.oninput = () => {
    v.지적내용 = memo.value;
    writeMemoBuffer(id, key, memo.value);
    clearTimeout(timer);
    pendingCommit = commit;
    timer = setTimeout(() => { if (pendingCommit === commit) { pendingCommit = null; commit(); } }, 400);
  };
  memo.onblur = () => { if (pendingCommit === commit) { clearTimeout(timer); pendingCommit = null; commit(); } };

  document.getElementById("clear").onclick = async () => {
    const after = (round.entries?.[key] || round.drafts?.[key] || {}).조치사진 || [];
    const nPhotos = v.사진.length + after.length;
    if (nPhotos && !confirm(`이 항목의 입력과 사진 ${nPhotos}장(조치 후 사진 포함)을 모두 지울까요? 되돌릴 수 없습니다.`)) return;
    clearTimeout(timer);
    if (pendingCommit === commit) pendingCommit = null;
    const photoIds = [...v.사진, ...after];
    v = { 결과: null, 점수: null, 지적내용: "", 사진: [] };
    commit();
    await saveChain; // 저장이 끝난 뒤 다시 그려야 지운 상태가 보인다
    await db.deletePhotos(photoIds);
    renderShop(id, si, sj, k);
  };

  // 필수값이 빠진 항목(작성중)에서는 다른 항목으로 넘어가지 못하게 막는다
  async function go(target) {
    clearTimeout(timer);
    if (pendingCommit === commit) pendingCommit = null;
    const check = commit();
    await saveChain;
    if (check.state === "draft") {
      statusEl.scrollIntoView({ block: "center", behavior: "smooth" });
      toast("필수값을 채우거나 '이 항목 입력 지우기'를 누르세요.");
      return;
    }
    location.hash = target;
  }
  document.getElementById("prev").onclick = () => go(`${base}/${k - 1}`);
  document.getElementById("next").onclick = () => go(last ? stopHash : `${base}/${k + 1}`);
  app.querySelectorAll(".chip").forEach((c) => (c.onclick = () => {
    const t = Number(c.dataset.k);
    if (t !== k) go(`${base}/${t}`);
  }));
}

// ---------- S7 설정·기준정보 ----------
async function renderSettings() {
  setHeader("설정 · 기준정보", "#/");
  const master = await db.getMaster();
  app.innerHTML = `
    <h2>점검 기준정보 가져오기</h2>
    <div class="card stack">
      <p class="muted" style="margin:0">엑셀 파일에 <b>매장</b> 시트(휴게소, 입점매장, 매장담당자, 연락처)와
      <b>점검항목</b> 시트(순서, 점검항목, 점수척도)가 있어야 합니다. 파일은 이 기기 안에서만 읽습니다.</p>
      <label class="btn block file-btn">엑셀 파일 선택
        <input id="file" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">
      </label>
      <button id="sample" class="btn secondary block" type="button">시험용 더미 기준정보 불러오기</button>
    </div>
    <h2>자가 점검</h2>
    <div class="card stack">
      <p class="muted" style="margin:0">계산·검사 규칙이 제대로 동작하는지 시험용 가짜 데이터로 자동 확인합니다 (실제 데이터는 건드리지 않음).</p>
      <a class="btn secondary block" href="tests/">자가 점검 실행</a>
    </div>
    <h2>백업 파일 가져오기</h2>
    <div class="card stack">
      <p class="muted" style="margin:0">다른 담당자가 보낸 회차 백업(위생점검백업_….zip)을 가져와 이어서 입력합니다.
        가져오기 전에 내용을 먼저 보여줍니다.</p>
      <label class="btn secondary block file-btn">백업 파일 선택
        <input id="restore-file" type="file" accept=".zip,application/zip">
      </label>
    </div>
    ${master ? masterView(master, `현재 기준정보 · ${esc(master.fileName)}`, fmtDateTime(master.importedAt) + " 가져옴") : ""}`;

  document.getElementById("restore-file").onchange = async (ev) => {
    const f = ev.target.files[0];
    if (!f) return;
    await startRestore(f.name, await f.arrayBuffer());
  };
  document.getElementById("file").onchange = async (ev) => {
    const f = ev.target.files[0];
    if (f) await startImport(f.name, await f.arrayBuffer());
  };
  document.getElementById("sample").onclick = async () => {
    const res = await fetch("sample/checklist_master_dummy.xlsx");
    await startImport("checklist_master_dummy.xlsx (더미)", await res.arrayBuffer());
  };
}

async function startRestore(fileName, buffer) {
  try {
    pendingRestore = { fileName, data: await readBackup(buffer) };
  } catch (e) {
    pendingRestore = { fileName, error: e.message || String(e) };
  }
  location.hash = "#/restore";
}

async function startImport(fileName, buffer) {
  let result;
  try {
    result = validateMaster(await readXlsx(buffer));
  } catch (e) {
    result = { master: null, errors: [e.message || String(e)], warnings: [] };
  }
  pendingImport = { fileName, result };
  location.hash = "#/settings/preview";
}

function masterView(m, heading, sub) {
  const shops = m.stops.flatMap((s) => s.shops.map((sh) => ({ stop: s.name, ...sh })));
  return `
    <h2>${heading}</h2>
    ${sub ? `<p class="muted">${esc(sub)}</p>` : ""}
    <p>휴게소 <b>${m.stops.length}</b>곳 · 입점매장 <b>${shops.length}</b>개 · 점검항목 <b>${m.items.length}</b>개
      → 회차당 점검 <b>${shops.length * m.items.length}</b>건</p>
    <div class="table-wrap"><table>
      <tr><th>휴게소</th><th>입점매장</th><th>담당자</th><th>연락처</th></tr>
      ${shops.map((s) => `<tr><td>${esc(s.stop)}</td><td>${esc(s.name)}</td><td>${esc(maskName(s.person) || "-")}</td><td>${esc(maskPhone(s.phone) || "-")}</td></tr>`).join("")}
    </table></div>
    <p class="muted">담당자 이름·연락처는 화면에 가려서 표시합니다.</p>
    <div class="table-wrap"><table>
      <tr><th>순서</th><th>점검항목</th><th>점수척도</th></tr>
      ${m.items.map((it, k) => `<tr><td>${k + 1}</td><td>${esc(it.name)}</td><td>${it.scale.min}~${it.scale.max}</td></tr>`).join("")}
    </table></div>`;
}

function renderPreview() {
  if (!pendingImport) { location.hash = "#/settings"; return; }
  setHeader("기준정보 확인", "#/settings");
  const { fileName, result } = pendingImport;
  const { master, errors, warnings } = result;
  const list = (arr) => `<ul>${arr.map((x) => `<li>${esc(x)}</li>`).join("")}</ul>`;
  app.innerHTML = `
    <p class="muted">파일: ${esc(fileName)}</p>
    ${errors.length ? `<div class="notice bad"><b>오류 ${errors.length}건 — 엑셀을 고친 뒤 다시 가져오세요.</b>${list(errors)}</div>` : ""}
    ${warnings.length ? `<div class="notice warn"><b>확인 ${warnings.length}건</b>${list(warnings)}</div>` : ""}
    ${!errors.length ? `<div class="notice ok">문제없이 읽었습니다. 아래 내용이 맞으면 저장하세요.</div>` : ""}
    ${master ? masterView(master, "가져올 내용") : ""}
    <div class="stack" style="margin-top:16px">
      <button id="save" class="btn block" ${errors.length ? "disabled" : ""}>이 기준정보로 저장</button>
      <a class="btn secondary block" href="#/settings">취소</a>
    </div>`;
  document.getElementById("save").onclick = async () => {
    await db.putMaster({ ...master, fileName, importedAt: new Date().toISOString() });
    pendingImport = null;
    toast("기준정보를 저장했습니다.");
    location.hash = "#/";
  };
}

// ---------- 라우터 ----------
const routes = [
  [/^#?\/?$/, renderHome],
  [/^#\/settings$/, renderSettings],
  [/^#\/settings\/preview$/, renderPreview],
  [/^#\/restore$/, renderRestore],
  [/^#\/round-new$/, renderNewRound],
  [/^#\/round\/([^/]+)$/, (id) => renderRound(decodeURIComponent(id))],
  [/^#\/round\/([^/]+)\/review$/, (id) => renderReview(decodeURIComponent(id))],
  [/^#\/round\/([^/]+)\/export$/, (id) => renderExport(decodeURIComponent(id))],
  [/^#\/round\/([^/]+)\/followup$/, (id) => renderFollowup(decodeURIComponent(id))],
  [/^#\/round\/([^/]+)\/notice\/(\d+)$/, (id, si) => renderNotice(decodeURIComponent(id), Number(si))],
  [/^#\/round\/([^/]+)\/followup\/(\d+)\/(\d+)\/(\d+)$/,
    (id, si, sj, k) => renderFollowupItem(decodeURIComponent(id), Number(si), Number(sj), Number(k))],
  [/^#\/round\/([^/]+)\/stop\/(\d+)$/, (id, si) => renderStop(decodeURIComponent(id), Number(si))],
  [/^#\/round\/([^/]+)\/stop\/(\d+)\/shop\/(\d+)(?:\/(\d+))?$/,
    (id, si, sj, k) => renderShop(decodeURIComponent(id), Number(si), Number(sj), Number(k || 0))],
];

async function route() {
  await flushPending(); // 화면을 떠나기 전에 입력 중이던 값 저장
  const [h0, q = ""] = (location.hash || "#/").split("?");
  routeParams = new URLSearchParams(q);
  const h = h0;
  for (const [re, fn] of routes) {
    const m = h.match(re);
    if (m) {
      try { await fn(...m.slice(1)); } catch (e) {
        console.error(e);
        app.innerHTML = `<div class="notice bad">오류가 발생했습니다: ${esc(e.message || e)}</div>`;
      }
      window.scrollTo(0, 0);
      return;
    }
  }
  location.hash = "#/";
}

window.addEventListener("hashchange", route);
// 앱을 닫거나 백그라운드로 보낼 때 입력 중이던 지적내용 저장
document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") flushPending(); });
requestPersistentStorage();

// 오프라인 동작 (sw.js) 과 통신 상태 표시
if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch((e) => console.warn("서비스 워커 등록 실패", e));
const offlineBar = document.getElementById("offline");
const syncOnline = () => { offlineBar.hidden = navigator.onLine; };
window.addEventListener("online", syncOnline);
window.addEventListener("offline", syncOnline);
syncOnline();
recoverMemoBuffer().catch(console.error).finally(route);
