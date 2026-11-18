// 웹앱 안 카메라 (getUserMedia). 휴대폰 기본 카메라 앱을 쓰지 않으므로 갤러리에 사진이 남지 않는다.
// (2026-11-18 결정: 사진은 웹앱 안에만 저장, 갤러리 보관 안 함)
// capturePhoto() → 찍은 사진(Blob) / PC 시험용으로 고른 파일(File) / 취소하면 null

let dlg;
function ensureDialog() {
  if (dlg) return dlg;
  dlg = document.createElement("dialog");
  dlg.id = "camera";
  dlg.className = "viewer camera";
  dlg.setAttribute("aria-label", "사진 촬영");
  dlg.innerHTML = `
    <p id="cam-title" class="small" style="margin:0 0 8px;font-weight:700"></p>
    <video id="cam-video" playsinline muted autoplay></video>
    <p id="cam-msg" class="muted small" role="status"></p>
    <div class="viewer-actions">
      <button type="button" id="cam-cancel" class="btn secondary">취소</button>
      <button type="button" id="cam-shot" class="btn" disabled>촬영</button>
    </div>
    <label class="link-btn file-btn" id="cam-file-wrap" hidden>파일에서 가져오기 (PC 시험용 · 휴대폰에서는 쓰지 마세요)
      <input id="cam-file" type="file" accept="image/*">
    </label>`;
  document.body.appendChild(dlg);
  return dlg;
}

export function capturePhoto(title = "사진 촬영") {
  const d = ensureDialog();
  const video = d.querySelector("#cam-video");
  const msg = d.querySelector("#cam-msg");
  const shot = d.querySelector("#cam-shot");
  const fileWrap = d.querySelector("#cam-file-wrap");
  const file = d.querySelector("#cam-file");
  d.querySelector("#cam-title").textContent = title;
  let stream = null;
  let result = null;

  return new Promise((resolve) => {
    const finish = () => {
      stream?.getTracks().forEach((t) => t.stop());
      stream = null;
      video.srcObject = null;
      resolve(result);
    };
    d.addEventListener("close", finish, { once: true });
    d.querySelector("#cam-cancel").onclick = () => d.close();
    shot.onclick = async () => {
      if (!video.videoWidth) return;
      const c = document.createElement("canvas");
      c.width = video.videoWidth;
      c.height = video.videoHeight;
      c.getContext("2d").drawImage(video, 0, 0);
      result = await new Promise((r) => c.toBlob(r, "image/jpeg", 0.92));
      d.close();
    };
    file.onchange = () => {
      result = file.files[0] || null;
      file.value = "";
      d.close();
    };

    shot.disabled = true;
    fileWrap.hidden = true;
    video.hidden = false;
    msg.textContent = "카메라를 켜는 중…";
    d.showModal();
    (async () => {
      try {
        if (!navigator.mediaDevices?.getUserMedia) throw new Error("이 브라우저에서는 웹앱 카메라를 쓸 수 없습니다.");
        stream = await navigator.mediaDevices.getUserMedia({
          video: { facingMode: { ideal: "environment" }, width: { ideal: 1920 }, height: { ideal: 1440 } }, audio: false,
        });
        if (!d.open) { stream.getTracks().forEach((t) => t.stop()); return; } // 켜지는 사이 취소함
        video.srcObject = stream;
        // 카메라 화면이 4초 안에 나오지 않으면 무한정 기다리지 않고 안내
        const ready = await Promise.race([
          video.play().then(() => true).catch(() => false),
          new Promise((r) => setTimeout(() => r(false), 4000)),
        ]);
        if (!d.open) return;
        if (!ready && !video.videoWidth) {
          msg.textContent = "카메라 화면이 나오지 않습니다. 취소한 뒤 다시 시도하거나, 다른 앱이 카메라를 쓰고 있지 않은지 확인하세요.";
          return;
        }
        msg.textContent = "";
        shot.disabled = false;
      } catch (e) {
        video.hidden = true;
        const denied = e && (e.name === "NotAllowedError" || e.name === "SecurityError");
        msg.textContent = denied
          ? "카메라 사용이 허용되지 않았습니다. 브라우저 설정에서 이 사이트의 카메라를 허용해 주세요."
          : `카메라를 열 수 없습니다. (${e?.message || e})`;
        fileWrap.hidden = false;
      }
    })();
  });
}
