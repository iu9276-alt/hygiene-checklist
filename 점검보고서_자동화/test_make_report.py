"""보고서 도구 자동 테스트 (검증기준.md 의 R-1 ~ R-6)

실행:  python test_make_report.py
- 한글(HWP) 변환은 하지 않고 집계·사진 연결·시정 상태 함수만 시험한다.
- 원본 더미 파일은 읽기만 하고, 시험용 파일은 임시 폴더에 만들었다가 지운다.
"""
import datetime as dt
import os
import shutil
import sys
import tempfile
import traceback

from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import make_report as m  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
DUMMY = os.path.join(HERE, "..", "더미데이터", "위생점검", "2026", "4분기")
TESTS = []


def test(tid, name):
    def deco(fn):
        TESTS.append((tid, name, fn))
        return fn
    return deco


def run_pipeline(xlsx, photos, today="2026-11-18"):
    recs, col_issues, raw = m.read_excel(xlsx)
    m.classify(recs)
    by_stop, by_item = m.aggregate(recs)
    findings = m.findings_of(recs, dt.date.fromisoformat(today))
    unmatched, ambiguous, missing = m.match_photos(findings, photos)
    return dict(recs=recs, col_issues=col_issues, raw=raw, by_stop=by_stop, by_item=by_item,
                findings=findings, unmatched=unmatched, ambiguous=ambiguous, missing=missing)


# ---------- 기존 더미 엑셀 ----------
@test("R-1", "기존 더미 엑셀: 48/39/8, 판정누락 1, 평균 4.06·4.19·4.38 (열 순서 다른 시트 포함)")
def _():
    a = run_pipeline(os.path.join(DUMMY, "위생점검결과_2026_4Q.xlsx"), os.path.join(DUMMY, "사진"))
    s = a["by_stop"]
    assert a["raw"] == 48, a["raw"]
    assert [sum(b[k] for b in s.values()) for k in ("ok", "ng", "na")] == [39, 8, 1]
    assert [s[x]["avg"] for x in ("가람휴게소", "누리휴게소", "다솜휴게소")] == [4.06, 4.19, 4.38]
    scores = [r["점수_정리"] for r in a["recs"] if r["점수_정리"] is not None]
    assert round(sum(scores) / len(scores), 2) == 4.21


@test("R-2", "사진 없는 부적합 1건(누리 카페), 연결 안 되는 사진 1장(주차장조명)")
def _():
    a = run_pipeline(os.path.join(DUMMY, "위생점검결과_2026_4Q.xlsx"), os.path.join(DUMMY, "사진"))
    no_photo = [f["rec"]["입점매장"] + "/" + f["rec"]["점검항목"] for f in a["findings"] if not f["사진"]]
    assert no_photo == ["카페/손씻기 시설"], no_photo
    assert [os.path.basename(p) for p in a["unmatched"]] == ["가람휴게소_주차장조명.jpg"], a["unmatched"]


# ---------- 웹앱 형식 엑셀 (사진파일·조치사진파일) ----------
COLS = ["점검일", "입점매장", "점검항목", "결과", "점수", "지적내용", "매장담당자", "연락처",
        "지적일", "재방문일", "완료여부", "사진파일", "조치사진파일", "비고"]
D10, D16 = dt.datetime(2026, 11, 10), dt.datetime(2026, 11, 16)


def webapp_fixture(tmp, drop=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "가람휴게소"
    ws.append(COLS)
    ws.append([D10, "한식당", "냉장·냉동 온도", "부적합", 2, "냉장고 8℃", "김민수", "010-2481-3579", D10, D16, "Y",
               "가람휴게소_냉장냉동온도_한식당.jpg, 가람휴게소_냉장냉동온도_한식당_2.jpg",
               "가람휴게소_냉장냉동온도_한식당_조치후.jpg", None])
    ws.append([D10, "분식코너", "냉장·냉동 온도", "부적합", 1, "냉동고 성에", "이서연", "010-3692-4815", D10, None, "N",
               "가람휴게소_냉장냉동온도_분식코너.jpg", None, None])
    ws.append([D10, "한식당", "방충·방서", "부적합", 2, "방충망 파손", "김민수", "010-2481-3579", D10, None, "N",
               "가람휴게소_방충방서.jpg", None, None])
    ws.append([D10, "한식당", "조리장 청결", None, 3, None, "김민수", "010-2481-3579", None, None, None, None, None,
               "작성중(입력값: 판정 없음) — 적합 / 부적합을 선택하세요."])
    xlsx = os.path.join(tmp, "web.xlsx")
    wb.save(xlsx)
    files = {"20261110": ["가람휴게소_냉장냉동온도_한식당.jpg", "가람휴게소_냉장냉동온도_한식당_2.jpg",
                          "가람휴게소_냉장냉동온도_분식코너.jpg", "가람휴게소_방충방서.jpg"],
             "20261116": ["가람휴게소_냉장냉동온도_한식당_조치후.jpg"]}
    for day, names in files.items():
        os.makedirs(os.path.join(tmp, "사진", day), exist_ok=True)
        for n in names:
            if n != drop:
                open(os.path.join(tmp, "사진", day, n), "wb").write(b"\xff\xd8\xff\xd9")
    return xlsx, os.path.join(tmp, "사진")


@test("R-3", "웹앱 형식: 두 매장·여러 장 사진이 각자 지적사항에 정확히 연결, 모르는 열 0")
def _():
    with tempfile.TemporaryDirectory() as tmp:
        a = run_pipeline(*webapp_fixture(tmp))
        got = {f["rec"]["입점매장"] + "/" + f["rec"]["점검항목"]: [os.path.basename(p) for p in f["사진"]] for f in a["findings"]}
        assert got == {
            "한식당/냉장·냉동 온도": ["가람휴게소_냉장냉동온도_한식당.jpg", "가람휴게소_냉장냉동온도_한식당_2.jpg"],
            "분식코너/냉장·냉동 온도": ["가람휴게소_냉장냉동온도_분식코너.jpg"],
            "한식당/방충·방서": ["가람휴게소_방충방서.jpg"],
        }, got
        assert a["unmatched"] == [] and a["missing"] == [] and a["ambiguous"] == []
        assert a["col_issues"] == [], a["col_issues"]
        assert sum(b["na"] for b in a["by_stop"].values()) == 1  # 작성중 행 → 판정누락


@test("R-4", "사진파일 열에 적힌 파일이 없으면 누락 1건으로 보고")
def _():
    with tempfile.TemporaryDirectory() as tmp:
        a = run_pipeline(*webapp_fixture(tmp, drop="가람휴게소_냉장냉동온도_분식코너.jpg"))
        assert [n for _, n in a["missing"]] == ["가람휴게소_냉장냉동온도_분식코너.jpg"], a["missing"]
        assert [f["rec"]["입점매장"] for f in a["findings"] if not f["사진"]] == ["분식코너"]


@test("R-5", "조치사진파일: 재방문일 폴더의 조치 후 사진 연결, 미연결 사진 0")
def _():
    with tempfile.TemporaryDirectory() as tmp:
        a = run_pipeline(*webapp_fixture(tmp))
        f = next(x for x in a["findings"] if x["rec"]["입점매장"] == "한식당" and x["rec"]["점검항목"] == "냉장·냉동 온도")
        assert [os.path.basename(os.path.dirname(p)) + "/" + os.path.basename(p) for p in f["조치사진"]] == \
            ["20261116/가람휴게소_냉장냉동온도_한식당_조치후.jpg"], f["조치사진"]


@test("R-6", "시정 상태 표기가 웹앱과 같음 (기준일 11-18 / 11-24 / 11-25)")
def _():
    with tempfile.TemporaryDirectory() as tmp:
        xlsx, photos = webapp_fixture(tmp)
        for today, want in (("2026-11-18", "진행 중(D-6)"), ("2026-11-24", "진행 중(D-day)"), ("2026-11-25", "기한 초과(1일)")):
            a = run_pipeline(xlsx, photos, today)
            states = {f["rec"]["입점매장"] + "/" + f["rec"]["점검항목"]: f["상태"] for f in a["findings"]}
            assert states["한식당/냉장·냉동 온도"] == "시정 완료", states
            assert states["한식당/방충·방서"] == want, (today, states)


def main():
    passed = 0
    for tid, name, fn in TESTS:
        try:
            fn()
            passed += 1
            print(f"  ✓ {tid} {name}")
        except Exception as e:  # noqa: BLE001
            print(f"  ✗ {tid} {name}\n      {type(e).__name__}: {e}")
            traceback.print_exc(limit=1)
    print(f"\n{'모두 통과' if passed == len(TESTS) else f'실패 {len(TESTS) - passed}건'} — 통과 {passed} / 전체 {len(TESTS)}")
    sys.exit(0 if passed == len(TESTS) else 1)


if __name__ == "__main__":
    main()
