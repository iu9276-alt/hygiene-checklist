"""위생점검 결과 엑셀 -> 한글(HWP) 보고서 초안 + 검산표

사용 예:
  python make_report.py --excel "...\위생점검결과_2026_4Q.xlsx" --photos "...\사진" ^
      --manual "...\수작업집계_2026_4Q.xlsx" --out "...\결과\2026_4Q" --round "2026년 4분기"

원칙
  - 원본 엑셀·사진은 읽기만 한다 (덮어쓰지 않음).
  - 열은 위치가 아니라 열 이름으로 찾는다. 모르는 열·없는 열은 검산표에 보고한다.
  - 결과가 비었거나 형식이 다른 값은 고치지 않고 '확인 필요'로 보고한다.
  - 원본에 없는 수치·문장은 만들지 않는다. 요약 문장은 집계값만으로 만든다.
  - 사람 이름·전화번호는 보고서에 마스킹해서 넣는다.
"""
import argparse
import datetime as dt
import html
import os
import random
import re
import shutil
import subprocess
import sys
from collections import Counter, defaultdict

from openpyxl import load_workbook

REQUIRED = ["점검일", "입점매장", "점검항목", "결과", "점수"]
OPTIONAL = ["지적내용", "매장담당자", "연락처", "지적일", "재방문일", "완료여부"]
# 현장 체크리스트 웹앱이 내보내는 열.
#   사진파일: 지적사항에 연결할 사진 파일명 (쉼표로 여러 개). 있으면 파일명 규칙보다 우선.
#   조치사진파일: 재방문 때 찍은 조치 후 사진 파일명 (재방문일 폴더). 보고서 사진란에 "(조치 후)"로 표시.
#   비고: 작성중·미입력 표시 등. 집계에는 쓰지 않음.
EXTRA_KNOWN = ["사진파일", "조치사진파일", "비고"]
KNOWN = REQUIRED + OPTIONAL
RESULTS = {"적합", "부적합"}
SCORE_MIN, SCORE_MAX = 1, 5
DEADLINE_DAYS = 14          # 업무_맥락.md: 지적일 +14일 안에 재방문 확인해야 완료
SAMPLE_N = 5
PHONE_RE = re.compile(r"(01\d)[-. ]?(\d{3,4})[-. ]?(\d{4})")


# ---------- 공통 ----------
def norm(s):
    return re.sub(r"[\s·・.\-_()]", "", str(s or ""))


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str) and v.strip():
        for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d"):
            try:
                return dt.datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def mask_name(name):
    n = str(name or "").strip()
    if len(n) <= 1:
        return n
    if len(n) == 2:
        return n[0] + "*"
    return n[0] + "*" * (len(n) - 2) + n[-1]


def mask_phone_text(text):
    return PHONE_RE.sub(lambda m: f"{m.group(1)}-{'*' * len(m.group(2))}-{m.group(3)}", str(text or ""))


def fmt_d(d):
    return d.strftime("%Y-%m-%d") if d else ""


def e(s):
    return html.escape(str(s if s is not None else ""))


# ---------- 1. 읽기 ----------
def read_excel(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    records, col_issues, raw_count = [], [], 0
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        header = [str(h).strip() if h is not None else "" for h in rows[0]]
        idx = {h: i for i, h in enumerate(header) if h}
        unknown = [h for h in idx if h not in KNOWN and h not in EXTRA_KNOWN]
        missing = [c for c in KNOWN if c not in idx]
        if unknown or missing:
            col_issues.append((ws.title, unknown, missing))
        if any(c not in idx for c in REQUIRED):
            continue  # 필수 열이 없으면 이 시트는 집계하지 않고 보고만 함
        for rno, row in enumerate(rows[1:], start=2):
            if not any(v not in (None, "") for v in row):
                continue
            raw_count += 1
            rec = {"시트": ws.title, "행": rno, "휴게소": ws.title}
            for c in KNOWN + EXTRA_KNOWN:
                rec[c] = row[idx[c]] if c in idx and idx[c] < len(row) else None
            records.append(rec)
    wb.close()  # 읽기 전용으로 열면 닫을 때까지 파일이 잠겨 있으므로 바로 닫는다
    return records, col_issues, raw_count


# ---------- 2. 판정·집계 ----------
def classify(records):
    issues = []
    for r in records:
        res = str(r["결과"]).strip() if r["결과"] is not None else ""
        r["결과_정리"] = res if res in RESULTS else None
        if r["결과_정리"] is None:
            issues.append((r, f"결과 값이 '{res or '빈칸'}' (적합/부적합 아님)"))
        sc = r["점수"]
        if isinstance(sc, (int, float)) and SCORE_MIN <= sc <= SCORE_MAX:
            r["점수_정리"] = float(sc)
        else:
            r["점수_정리"] = None
            issues.append((r, f"점수 값이 '{sc}' ({SCORE_MIN}~{SCORE_MAX} 숫자 아님)"))
    return issues


def aggregate(records):
    by_stop = defaultdict(lambda: {"n": 0, "ok": 0, "ng": 0, "na": 0, "scores": []})
    by_item = defaultdict(Counter)
    for r in records:
        s = by_stop[r["휴게소"]]
        s["n"] += 1
        if r["결과_정리"] == "적합":
            s["ok"] += 1
        elif r["결과_정리"] == "부적합":
            s["ng"] += 1
            by_item[r["점검항목"]][r["휴게소"]] += 1
        else:
            s["na"] += 1
        if r["점수_정리"] is not None:
            s["scores"].append(r["점수_정리"])
    for s in by_stop.values():
        s["avg"] = round(sum(s["scores"]) / len(s["scores"]), 2) if s["scores"] else None
    return by_stop, by_item


def findings_of(records, today):
    out, seq = [], Counter()
    for r in records:
        if r["결과_정리"] != "부적합":
            continue
        seq[r["휴게소"]] += 1
        no = f"{r['휴게소'].replace('휴게소', '')}-{seq[r['휴게소']]:02d}"
        d0 = to_date(r["지적일"]) or to_date(r["점검일"])
        due = d0 + dt.timedelta(days=DEADLINE_DAYS) if d0 else None
        rv = to_date(r["재방문일"])
        done = str(r["완료여부"] or "").strip().upper() in ("Y", "O", "완료", "예")
        if done and rv and due and rv <= due:
            status = "시정 완료"
        elif done:
            status = "확인 필요(재방문일 없음/기한 후)"
        elif due and today > due:
            status = f"기한 초과({(today - due).days}일)"
        elif due:
            left = (due - today).days
            status = "진행 중(D-day)" if left == 0 else f"진행 중(D-{left})"  # 웹앱 시정 대기 화면과 같은 표기
        else:
            status = "확인 필요(지적일 없음)"
        out.append({"no": no, "rec": r, "지적일": d0, "기한": due, "재방문일": rv, "상태": status})
    return out


# ---------- 3. 사진 매칭 ----------
# 1순위: 행의 '사진파일' 열에 적힌 파일 (현장 체크리스트 웹앱이 기록)
# 2순위: '사진파일' 값이 없는 행만 파일명 규칙(휴게소_점검항목)으로 찾음 (예전 엑셀 호환)
def listed_files(rec, col="사진파일"):
    v = rec.get(col)
    if v is None:
        return []
    return [s.strip() for s in re.split(r"[,\n;]", str(v)) if s.strip()]


def match_photos(findings, photo_root):
    photos = []
    if photo_root and os.path.isdir(photo_root):
        for dp, _, fs in os.walk(photo_root):
            for f in fs:
                if f.lower().endswith((".jpg", ".jpeg", ".png")):
                    photos.append(os.path.join(dp, f))
    by_name = defaultdict(list)
    by_key = defaultdict(list)
    for p in photos:
        base = os.path.basename(p)
        by_name[base.lower()].append(p)
        stem = os.path.splitext(base)[0]
        if "_" in stem:
            a, b = stem.split("_", 1)
            by_key[(norm(a), norm(b))].append(p)

    used, ambiguous, missing_listed = set(), [], []
    fallback = [f for f in findings if not listed_files(f["rec"])]
    keycount = Counter((norm(f["rec"]["휴게소"]), norm(f["rec"]["점검항목"])) for f in fallback)

    def resolve(f, names, day_col, out):
        day = to_date(f["rec"][day_col])
        day = day.strftime("%Y%m%d") if day else None
        for n in names:
            cands = by_name.get(n.lower(), [])
            same_day = [p for p in cands if os.path.basename(os.path.dirname(p)) == day]
            if len(same_day) == 1 or len(cands) == 1:
                out.append((same_day or cands)[0])
            elif not cands:
                missing_listed.append((f, n))
            else:
                ambiguous.append((f, f"사진 '{n}'이(가) 여러 폴더에 있어 어느 것인지 정할 수 없음"))

    for f in findings:
        f["조치사진"] = []
        resolve(f, listed_files(f["rec"], "조치사진파일"), "재방문일", f["조치사진"])
        used.update(f["조치사진"])
        names = listed_files(f["rec"])
        if names:
            f["사진"] = []
            resolve(f, names, "점검일", f["사진"])
        else:
            k = (norm(f["rec"]["휴게소"]), norm(f["rec"]["점검항목"]))
            f["사진"] = list(by_key.get(k, []))
            if keycount[k] > 1:
                ambiguous.append((f, "같은 휴게소·항목 지적이 여러 건이라 사진 파일명만으로 구분 불가"))
        used.update(f["사진"])
    unmatched = [p for p in photos if p not in used]
    return unmatched, ambiguous, missing_listed


# ---------- 4. 과거 수작업 대조 ----------
def compare_manual(path, by_stop):
    if not path:
        return None
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()  # 읽기 전용으로 열면 닫을 때까지 파일이 잠겨 있으므로 바로 닫는다
    hdr = [str(h).strip() if h else "" for h in rows[0]]
    col = {h: i for i, h in enumerate(hdr)}
    pairs = [("점검건수", "n"), ("적합", "ok"), ("부적합", "ng"), ("평균점수", "avg")]
    out = []
    for row in rows[1:]:
        stop = row[col["휴게소"]]
        if not stop:
            continue
        auto = by_stop.get(stop)
        cells = []
        for label, k in pairs:
            m = row[col[label]] if label in col else None
            a = auto[k] if auto else None
            same = (m is not None and a is not None and abs(float(m) - float(a)) < 0.005)
            cells.append((label, m, a, same))
        out.append((stop, cells))
    return out


# ---------- 5. HTML ----------
CSS = """
body{font-family:'맑은 고딕';font-size:10pt}
h1{font-size:18pt;text-align:center}
h2{font-size:13pt;margin-top:14pt}
h3{font-size:11pt}
table{border-collapse:collapse;width:100%}
th,td{border:1px solid #444;padding:2pt 4pt;font-size:9pt}
th{background:#dde7f0}
.bad{background:#fde2e2;font-weight:bold}
.ok{background:#e2f3e5}
.note{font-size:9pt;color:#555}
"""


def tbl(head, rows, widths=None, row_cls=None):
    h = "<table><tr>" + "".join(
        f"<th{f' width={widths[i]}' if widths else ''}>{e(c)}</th>" for i, c in enumerate(head)) + "</tr>"
    for i, r in enumerate(rows):
        cls = row_cls[i] if row_cls else ""
        h += "<tr>" + "".join(
            f"<td{f' class={cls}' if cls else ''}>{c if str(c).startswith('<') else e(c)}</td>" for c in r) + "</tr>"
    return h + "</table>"


def build_html(a):
    by_stop, by_item, findings = a["by_stop"], a["by_item"], a["findings"]
    stops = list(by_stop)
    dates = sorted({to_date(r["점검일"]) for r in a["records"] if to_date(r["점검일"])})
    tot = {k: sum(by_stop[s][k] for s in stops) for k in ("n", "ok", "ng", "na")}
    allscores = [x for s in stops for x in by_stop[s]["scores"]]
    tot_avg = round(sum(allscores) / len(allscores), 2) if allscores else None
    items = sorted({r["점검항목"] for r in a["records"]}, key=str)

    P = ['<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">'
         f'<meta charset="utf-8"><style>{CSS}</style></head><body>']
    P.append(f"<h1>{e(a['round'])} 휴게소 위생점검 결과 보고(초안)</h1>")
    P.append(f"<p class=note>작성: 자동 생성 초안 · 생성일 {fmt_d(a['today'])} · 원본: {e(os.path.basename(a['excel']))}</p>")

    # 1. 개요
    P.append("<h2>1. 점검 개요</h2>")
    P.append(tbl(["구분", "내용"], [
        ["점검 회차", a["round"]],
        ["점검 기간", f"{fmt_d(dates[0])} ~ {fmt_d(dates[-1])}" if dates else "확인 필요"],
        ["대상 휴게소", f"{len(stops)}개소 ({', '.join(stops)})"],
        ["점검 항목", f"{len(items)}개 항목"],
        ["점검 건수", f"{tot['n']}건 (입점매장 × 항목)"],
    ], widths=["25%", "75%"]))

    # 2. 요약
    P.append("<h2>2. 요약</h2><ul>")
    na_txt = f", 판정 누락 {tot['na']}건" if tot["na"] else ""
    P.append(f"<li>{len(stops)}개 휴게소 {tot['n']}건 점검 결과 적합 {tot['ok']}건, 부적합 {tot['ng']}건"
             f"{na_txt}입니다.</li>")
    worst = max(stops, key=lambda s: by_stop[s]["ng"])
    P.append(f"<li>부적합이 가장 많은 곳은 {worst}({by_stop[worst]['ng']}건)이며, 전체 평균 점수는 {tot_avg}점입니다.</li>")
    item_tot = Counter({k: sum(v.values()) for k, v in by_item.items()})
    if item_tot:
        top = item_tot.most_common()
        mx = top[0][1]
        names = [k for k, v in top if v == mx]
        if len(names) == len(top):
            P.append(f"<li>부적합 {len(top)}개 항목이 모두 {mx}건씩으로, 특정 항목에 몰리지 않았습니다.</li>")
        elif len(names) > 1:
            P.append(f"<li>부적합이 가장 많은 항목은 {', '.join(names)}({mx}건씩)입니다.</li>")
        else:
            P.append(f"<li>부적합이 가장 많은 항목은 {names[0]}({mx}건)입니다.</li>")
    st = Counter(f["상태"].split("(")[0] for f in findings)
    P.append(f"<li>지적사항 {len(findings)}건 중 시정 완료 {st.get('시정 완료', 0)}건, 진행 중 {st.get('진행 중', 0)}건, "
             f"기한 초과 {st.get('기한 초과', 0)}건, 확인 필요 {st.get('확인 필요', 0)}건입니다.</li>")
    if a["check_fail"]:
        P.append(f"<li><b>검산에서 불일치·확인 필요 {a['check_fail']}건이 있어 결재 전 부록 확인이 필요합니다.</b></li>")
    P.append("</ul>")

    # 3. 집계표
    P.append("<h2>3. 집계표</h2><h3>가. 휴게소별</h3>")
    rows = []
    for s in stops:
        b = by_stop[s]
        judged = b["ok"] + b["ng"]
        rows.append([s, b["n"], b["ok"], b["ng"], b["na"],
                     f"{b['ng'] / judged * 100:.1f}%" if judged else "-", b["avg"]])
    judged = tot["ok"] + tot["ng"]
    rows.append(["합계", tot["n"], tot["ok"], tot["ng"], tot["na"],
                 f"{tot['ng'] / judged * 100:.1f}%" if judged else "-", tot_avg])
    P.append(tbl(["휴게소", "점검건수", "적합", "부적합", "판정누락", "부적합률", "평균점수"], rows))
    P.append("<h3>나. 항목별 부적합 건수</h3>")
    rows = [[it] + [by_item[it].get(s, 0) for s in stops] + [sum(by_item[it].values())] for it in items]
    rows.append(["합계"] + [by_stop[s]["ng"] for s in stops] + [tot["ng"]])
    P.append(tbl(["점검항목"] + stops + ["합계"], rows))

    # 4. 지적사항
    P.append(f"<h2>4. 지적사항 목록 ({len(findings)}건)</h2>")
    rows, cls = [], []
    md = lambda d: d.strftime("%m-%d") if d else "-"
    for f in findings:
        r = f["rec"]
        rows.append([f["no"], f"{r['휴게소']} / {r['입점매장']}",
                     r["점검항목"], r["점수"], mask_phone_text(r["지적내용"]),
                     f"{mask_name(r['매장담당자'])} {mask_phone_text(r['연락처'])}",
                     f"{md(f['지적일'])} → {md(f['기한'])}", md(f["재방문일"]), f["상태"],
                     "있음" if f["사진"] else "없음"])
        cls.append("bad" if f["상태"].startswith(("기한 초과", "확인 필요")) or not f["사진"] else "")
    P.append(tbl(["번호", "휴게소 / 매장", "항목", "점수", "지적내용", "담당자", "지적일 → 기한",
                  "재방문", "시정상태", "사진"], rows, row_cls=cls,
                 widths=["7%", "13%", "10%", "5%", "21%", "12%", "10%", "6%", "10%", "6%"]))
    P.append(f"<p class=note>시정 기한 = 지적일 + {DEADLINE_DAYS}일. 기한 안에 재방문해 확인해야 '시정 완료'.</p>")

    # 사진
    P.append("<h3>지적사항 사진</h3>")
    cells = []
    for f in findings:
        for p, after in f.get("사진_rel", []):
            cap = f"{f['no']} {f['rec']['휴게소']} {f['rec']['입점매장']} · {f['rec']['점검항목']}" + (" (조치 후)" if after else "")
            cells.append(f"<img src='{p}' width=300><br>{e(cap)}")
    if cells:
        grid = [cells[i:i + 2] + [""] * (2 - len(cells[i:i + 2])) for i in range(0, len(cells), 2)]
        P.append("<table>" + "".join("<tr>" + "".join(f"<td width='50%'>{c}</td>" for c in row) + "</tr>"
                                     for row in grid) + "</table>")
    else:
        P.append("<p>연결된 사진 없음</p>")

    # 5. 부록 검산표
    P.append("<h2>[부록] 검산표</h2>")
    P.append("<h3>① 자동 검산</h3>")
    P.append(tbl(["검산 항목", "기준값", "비교값", "결과", "원인"],
                 [c[:5] for c in a["checks"]],
                 row_cls=["ok" if c[3] == "일치" else "bad" for c in a["checks"]]))

    P.append("<h3>② 과거 수작업 집계 대조</h3>")
    if a["manual"] is None:
        P.append("<p>수작업 집계 파일이 주어지지 않아 대조하지 않음</p>")
    else:
        rows, cls = [], []
        for stop, cells_ in a["manual"]:
            for label, m, au, same in cells_:
                rows.append([stop, label, m, au, "일치" if same else "불일치"])
                cls.append("ok" if same else "bad")
        P.append(tbl(["휴게소", "항목", "수작업", "자동", "결과"], rows, row_cls=cls))

    P.append(f"<h3>③ 표본 확인 (무작위 {len(a['sample'])}건)</h3>")
    P.append("<p class=note>원본 엑셀의 해당 시트·행을 열어 아래 내용과 같은지 눈으로 대조하세요.</p>")
    P.append(tbl(["보고서 번호", "원본 시트", "원본 행", "항목", "결과", "점수", "지적내용"],
                 [[f["no"], f["rec"]["시트"], f["rec"]["행"], f["rec"]["점검항목"], f["rec"]["결과"],
                   f["rec"]["점수"], mask_phone_text(f["rec"]["지적내용"])] for f in a["sample"]]))

    P.append("<h3>④ 확인 필요 데이터 (임의 수정하지 않음)</h3>")
    rows = [[r["시트"], r["행"], r["입점매장"], r["점검항목"], why] for r, why in a["data_issues"]]
    for sheet, unk, mis in a["col_issues"]:
        if unk:
            rows.append([sheet, "머리글", "-", "-", f"처음 보는 열: {', '.join(unk)} (집계에 쓰지 않음)"])
        if mis:
            rows.append([sheet, "머리글", "-", "-", f"없는 열: {', '.join(mis)}"])
    for f, why in a["ambiguous"]:
        rows.append([f["rec"]["시트"], f["rec"]["행"], f["rec"]["입점매장"], f["rec"]["점검항목"], why])
    P.append(tbl(["시트", "행", "매장", "항목", "내용"], rows) if rows else "<p>없음</p>")

    P.append("<h3>⑤ 사진 연결 점검</h3>")
    rows = [[f["no"], f["rec"]["휴게소"], f["rec"]["점검항목"], "부적합인데 사진 없음"] for f in findings if not f["사진"]]
    rows += [[f["no"], f["rec"]["휴게소"], n, "사진파일 열에 적힌 파일이 사진 폴더에 없음"] for f, n in a["missing_listed"]]
    rows += [["-", os.path.basename(os.path.dirname(p)), os.path.basename(p), "어느 지적사항에도 연결 안 된 사진"]
             for p in a["unmatched"]]
    P.append(tbl(["번호", "휴게소/폴더", "항목/파일", "내용"], rows) if rows else "<p>문제 없음</p>")

    P.append("<h3>⑥ 개인정보 점검</h3><p>__PII_RESULT__</p>")
    P.append("</body></html>")
    return "\n".join(P)


def pii_check(doc, records):
    text = re.sub(r"<[^>]+>", " ", doc)
    leaks = []
    for r in records:
        n = str(r["매장담당자"] or "").strip()
        if len(n) >= 2 and n in text:
            leaks.append(f"이름 노출: {mask_name(n)}")
    for m in PHONE_RE.finditer(text):
        leaks.append(f"전화번호 노출: {m.group(1)}-****-{m.group(3)}")
    return sorted(set(leaks))


# ---------- main ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True)
    ap.add_argument("--photos")
    ap.add_argument("--manual")
    ap.add_argument("--out", required=True)
    ap.add_argument("--round", required=True, help="예: 2026년 4분기")
    ap.add_argument("--today", help="기준일 YYYY-MM-DD (기본: 오늘)")
    args = ap.parse_args()
    today = to_date(args.today) if args.today else dt.date.today()

    records, col_issues, raw_count = read_excel(args.excel)
    data_issues = classify(records)
    by_stop, by_item = aggregate(records)
    findings = findings_of(records, today)
    unmatched, ambiguous, missing_listed = match_photos(findings, args.photos)
    manual = compare_manual(args.manual, by_stop)

    tot_n = sum(b["n"] for b in by_stop.values())
    tot_ok = sum(b["ok"] for b in by_stop.values())
    tot_ng = sum(b["ng"] for b in by_stop.values())
    na_rows = [f"{r['시트']} {r['행']}행" for r in records if r["결과_정리"] is None]
    checks = [
        ["원본 행 수 = 집계 대상 건수", raw_count, tot_n, "일치" if raw_count == tot_n else "불일치",
         "" if raw_count == tot_n else "필수 열이 없는 시트가 있음"],
        ["적합 + 부적합 = 전체 점검 건수", tot_n, tot_ok + tot_ng,
         "일치" if tot_ok + tot_ng == tot_n else "불일치", ", ".join(na_rows) + " 결과 빈칸/오류" if na_rows else ""],
        ["지적사항 목록 건수 = 부적합 건수", tot_ng, len(findings),
         "일치" if tot_ng == len(findings) else "불일치", ""],
        ["사진 누락 지적사항 = 0", 0, sum(1 for f in findings if not f["사진"]),
         "일치" if all(f["사진"] for f in findings) else "불일치", "부록 ⑤ 참고"],
    ]
    if any(listed_files(f["rec"]) for f in findings):
        checks.append(["사진파일 열에 적힌 사진이 모두 있음 (없는 파일 = 0)", 0, len(missing_listed),
                       "일치" if not missing_listed else "불일치", "부록 ⑤ 참고"])
    if manual is not None:
        diff = sum(1 for _, cs in manual for c in cs if not c[3])
        checks.append(["과거 수작업 집계와 차이 = 0", 0, diff, "일치" if diff == 0 else "불일치", "부록 ② 참고"])

    rnd = random.Random(os.path.basename(args.excel))
    sample = sorted(rnd.sample(findings, min(SAMPLE_N, len(findings))), key=lambda f: f["no"])

    # 한글이 HTML을 불러올 때 한글·절대 경로 그림을 못 읽으므로, 출력 폴더에 영문 이름 사본을 두고 상대 경로로 넣는다
    os.makedirs(args.out, exist_ok=True)
    img_dir = os.path.join(args.out, "_img")
    os.makedirs(img_dir, exist_ok=True)
    k = 0
    for f in findings:
        f["사진_rel"] = []
        for p, after in [(p, False) for p in f["사진"]] + [(p, True) for p in f.get("조치사진", [])]:
            k += 1
            name = f"p{k:03d}{os.path.splitext(p)[1].lower()}"
            shutil.copyfile(p, os.path.join(img_dir, name))
            f["사진_rel"].append((f"_img/{name}", after))

    a = dict(records=records, by_stop=by_stop, by_item=by_item, findings=findings, unmatched=unmatched,
             ambiguous=ambiguous, missing_listed=missing_listed, manual=manual, checks=checks, sample=sample,
             data_issues=data_issues,
             col_issues=col_issues, today=today, round=args.round, excel=args.excel)
    a["check_fail"] = sum(1 for c in checks if c[3] != "일치") + len(data_issues)
    doc = build_html(a)
    leaks = pii_check(doc.replace("__PII_RESULT__", ""), records)
    checks.append(["마스킹 안 된 개인정보 = 0", 0, len(leaks), "일치" if not leaks else "불일치", "; ".join(leaks)])
    doc = doc.replace("__PII_RESULT__", "이상 없음 (이름·전화번호 모두 마스킹됨)" if not leaks
                      else "<b>노출 발견:</b> " + e("; ".join(leaks)))

    os.makedirs(args.out, exist_ok=True)
    base = os.path.join(os.path.abspath(args.out), f"위생점검_결과보고_{args.round.replace(' ', '')}_초안")
    with open(base + ".html", "w", encoding="utf-8-sig") as fp:
        fp.write(doc)

    print("== 검산 결과 ==")
    for c in checks:
        print(f"  [{c[3]}] {c[0]}: 기준 {c[1]} / 비교 {c[2]} {('- ' + c[4]) if c[4] else ''}")
    print(f"  확인 필요 데이터 {len(data_issues)}건, 열 구성 이슈 {len(col_issues)}개 시트, 연결 안 된 사진 {len(unmatched)}개, "
          f"사진파일 열에 적혔는데 없는 사진 {len(missing_listed)}개")

    ps = os.path.join(os.path.dirname(os.path.abspath(__file__)), "to_hwp.ps1")
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps,
                        "-Html", base + ".html", "-Hwp", base + ".hwp", "-Pdf", base + ".pdf"],
                       capture_output=True, text=True, encoding="cp949", errors="replace")
    print("== 한글 변환 ==", (r.stdout or "").strip(), (r.stderr or "").strip()[:500])
    if r.returncode != 0:
        sys.exit(r.returncode)
    print("출력:", base + ".hwp")


if __name__ == "__main__":
    main()
