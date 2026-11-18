r"""배포 묶음 만들기 — 방법 ① 외부 정적 호스팅(GitHub Pages 등) / ② 사내 서버 공용

실행:  python make_deploy.py
결과:
  배포\dist\                          ← 이 폴더 안의 파일을 그대로 올리면 됨 (하위 경로에 둬도 동작)
  배포\위생점검웹앱_배포_YYYYMMDD.zip   ← 전산 담당에게 전달할 압축 파일

- 화면 파일만 넣는다 (정해진 목록). README·시험용 결과 파일·개발 도구는 넣지 않는다.
- 사내 서버용 web.config(IIS)·.htaccess(Apache)를 함께 넣는다. GitHub Pages 에서는 이 두 파일이 무시된다.
- GitHub Pages 용 .nojekyll 을 넣는다 (밑줄로 시작하는 파일 등을 그대로 제공하게).
- 원본 현장체크리스트 폴더는 읽기만 한다.
"""
import datetime as dt
import hashlib
import os
import shutil
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
APP = os.path.join(HERE, "..", "현장체크리스트")
DIST = os.path.join(HERE, "dist")

FILES = [
    "index.html", "manifest.json", "sw.js",
    "css/app.css",
    "icons/icon-192.png", "icons/icon-512.png",
    "js/app.js", "js/backup.js", "js/camera.js", "js/db.js", "js/entry.js", "js/export.js", "js/followup.js",
    "js/master.js", "js/notice.js", "js/photo.js", "js/summary.js", "js/xlsx-write.js", "js/xlsx.js", "js/zip.js",
    "tests/index.html", "tests/tests.js",                       # 설치 후 자가 점검용
    "sample/checklist_master_dummy.xlsx", "sample/checklist_master_bad.xlsx",  # 가짜 데이터 (자가 점검·시험용)
]
SERVER_FILES = {"web.config": "서버설정/web.config", ".htaccess": "서버설정/.htaccess"}


def check_sw_list():
    """sw.js 의 오프라인 파일 목록이 실제 화면 파일과 맞는지 (빠지면 오프라인에서 안 열림)"""
    sw = open(os.path.join(APP, "sw.js"), encoding="utf-8").read()
    need = [f for f in FILES if not f.startswith(("tests/", "sample/")) and f != "sw.js"]  # sw.js 는 자기 자신을 담지 않음
    missing = [f for f in need if f"./{f}" not in sw]
    if missing:
        raise SystemExit(f"sw.js 오프라인 목록에 빠진 파일: {missing}")
    version = sw.split('const CACHE = "', 1)[1].split('"', 1)[0]
    return version


def main():
    version = check_sw_list()
    if os.path.isdir(DIST):
        shutil.rmtree(DIST)
    for f in FILES:
        src = os.path.join(APP, f)
        if not os.path.isfile(src):
            raise SystemExit(f"파일 없음: {f}")
        dst = os.path.join(DIST, f)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)
    for name, src in SERVER_FILES.items():
        shutil.copy2(os.path.join(HERE, src), os.path.join(DIST, name))
    open(os.path.join(DIST, ".nojekyll"), "w").close()

    # 무엇이 들어갔는지 기록 (전산 담당 확인용)
    lines, total = [], 0
    for dp, _, fs in os.walk(DIST):
        for f in sorted(fs):
            p = os.path.join(dp, f)
            rel = os.path.relpath(p, DIST).replace(os.sep, "/")
            data = open(p, "rb").read()
            total += len(data)
            lines.append(f"{hashlib.sha256(data).hexdigest()[:16]}  {len(data):>8}  {rel}")
    lines.sort(key=lambda s: s.split()[-1])
    today = dt.date.today().strftime("%Y%m%d")
    manifest = [f"위생점검 체크리스트 웹앱 배포 묶음 · {today} · 오프라인 버전 {version}",
                f"파일 {len(lines)}개 · {total / 1024:.0f} KB", "", "SHA-256(앞16자)  크기(바이트)  파일", *lines]
    open(os.path.join(DIST, "배포목록.txt"), "w", encoding="utf-8").write("\n".join(manifest) + "\n")

    zpath = os.path.join(HERE, f"위생점검웹앱_배포_{today}.zip")
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for dp, _, fs in os.walk(DIST):
            for f in fs:
                p = os.path.join(dp, f)
                z.write(p, os.path.relpath(p, DIST))
    print("\n".join(manifest[:2]))
    print("dist:", os.path.abspath(DIST))
    print("zip :", os.path.abspath(zpath), f"({os.path.getsize(zpath) / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
