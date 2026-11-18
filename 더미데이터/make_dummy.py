r"""위생점검 더미 데이터 생성기 (시험용 가짜 데이터)

업무_맥락.md 3장의 구조를 따름:
  위생점검\{연도}\{분기}\위생점검결과_{연도}_{n}Q.xlsx  (휴게소별 시트)
  위생점검\{연도}\{분기}\사진\{YYYYMMDD}\               (사진 파일명 = 휴게소_점검항목.jpg)

휴게소·매장·담당자·연락처는 모두 가상 값이다.
일부러 넣은 시험 상황:
  - 누리휴게소 시트: 열 순서가 다르고 '비고' 열이 추가됨  -> 열 이름 기준 인식 확인
  - 다솜휴게소 시트: 결과(적합/부적합)가 빈 행 1개       -> 임의 보정 없이 보고되는지 확인
  - 부적합인데 사진이 없는 건 1개, 어느 지적사항에도 안 맞는 사진 1개
  - 수작업집계 파일에 일부러 틀린 숫자 1칸               -> 과거 대조에서 잡히는지 확인
"""
import datetime as dt
import os
import random

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from PIL import Image, ImageDraw, ImageFont

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "위생점검", "2026", "4분기")
PHOTO_ROOT = os.path.join(BASE, "사진")

ITEMS = ["조리장 청결", "냉장·냉동 온도", "식재료 유통기한", "종사자 위생복장",
         "손씻기 시설", "방충·방서", "화장실 청결", "폐기물 관리"]

STOPS = {
    "가람휴게소": {"date": dt.date(2026, 11, 10),
                 "shops": [("한식당", "김민수", "010-2481-3579"), ("분식코너", "이서연", "010-3692-4815")]},
    "누리휴게소": {"date": dt.date(2026, 11, 11),
                 "shops": [("푸드코트", "박지훈", "010-5813-2468"), ("카페", "최유진", "010-7924-1357")]},
    "다솜휴게소": {"date": dt.date(2026, 11, 12),
                 "shops": [("한식당", "정하늘", "010-4826-9137"), ("편의점", "남궁민", "010-6357-2849")]},
}

# 부적합으로 만들 (휴게소, 매장, 항목): (점수, 지적내용, 재방문일 오프셋 또는 None, 완료여부)
FAILS = {
    ("가람휴게소", "한식당", "냉장·냉동 온도"): (2, "냉장고 내부 온도 8℃로 기준(5℃ 이하) 초과", 6, "Y"),
    ("가람휴게소", "분식코너", "종사자 위생복장"): (1, "위생모 미착용 종사자 2명", None, "N"),
    ("가람휴게소", "한식당", "방충·방서"): (2, "출입문 방충망 파손", None, "N"),
    ("누리휴게소", "푸드코트", "식재료 유통기한"): (1, "유통기한 경과 소스 1종 보관", 5, "Y"),
    ("누리휴게소", "카페", "손씻기 시설"): (2, "손세정제 비치 안 됨", None, "N"),
    ("누리휴게소", "푸드코트", "폐기물 관리"): (2, "음식물 쓰레기통 뚜껑 미설치", None, "N"),
    ("다솜휴게소", "한식당", "조리장 청결"): (2, "후드 기름때 누적", None, "N"),
    ("다솜휴게소", "편의점", "화장실 청결"): (1, "직원 화장실 비누 없음 (담당 010-6357-2849 연락함)", None, "N"),
}
BLANK_RESULT = ("다솜휴게소", "편의점", "방충·방서")   # 결과 빈칸 행
NO_PHOTO = ("누리휴게소", "카페", "손씻기 시설")        # 사진 없는 부적합
EXTRA_PHOTO = ("가람휴게소", "주차장 조명")             # 매칭 안 되는 사진

STD_COLS = ["점검일", "입점매장", "점검항목", "결과", "점수", "지적내용",
            "매장담당자", "연락처", "지적일", "재방문일", "완료여부"]
NURI_COLS = ["점검일", "입점매장", "점검항목", "점수", "결과", "지적내용", "비고",
             "매장담당자", "연락처", "지적일", "재방문일", "완료여부"]


def rows_for(stop, info):
    rnd = random.Random(stop)
    out = []
    for shop, person, phone in info["shops"]:
        for item in ITEMS:
            key = (stop, shop, item)
            row = {"점검일": info["date"], "입점매장": shop, "점검항목": item,
                   "매장담당자": person, "연락처": phone, "비고": ""}
            if key in FAILS:
                score, memo, rv, done = FAILS[key]
                row.update(결과="부적합", 점수=score, 지적내용=memo, 지적일=info["date"],
                           재방문일=(info["date"] + dt.timedelta(days=rv)) if rv else None,
                           완료여부=done)
            else:
                row.update(결과="적합", 점수=rnd.choice([4, 5, 5]), 지적내용="",
                           지적일=None, 재방문일=None, 완료여부="")
            if key == BLANK_RESULT:
                row["결과"] = None
                row["점수"] = 3
            out.append(row)
    return out


def make_photo(path, title, sub):
    font_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    f_big = ImageFont.truetype(os.path.join(font_dir, "malgun.ttf"), 34)
    f_small = ImageFont.truetype(os.path.join(font_dir, "malgun.ttf"), 22)
    img = Image.new("RGB", (640, 480), (225, 232, 238))
    d = ImageDraw.Draw(img)
    d.rectangle([20, 20, 620, 460], outline=(90, 110, 130), width=4)
    d.text((50, 170), title, font=f_big, fill=(30, 40, 50))
    d.text((50, 240), sub, font=f_small, fill=(70, 80, 90))
    d.text((50, 400), "더미 사진 (시험용)", font=f_small, fill=(150, 60, 60))
    img.save(path, quality=85)


def main():
    os.makedirs(BASE, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)
    truth = {}
    for stop, info in STOPS.items():
        ws = wb.create_sheet(stop)
        cols = NURI_COLS if stop == "누리휴게소" else STD_COLS
        ws.append(cols)
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="DDEBF7")
        rows = rows_for(stop, info)
        for r in rows:
            ws.append([r.get(c) for c in cols])
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value, dt.date):
                    c.number_format = "yyyy-mm-dd"
        widths = {"지적내용": 42, "점검항목": 16, "연락처": 15}
        for i, c in enumerate(cols, 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = widths.get(c, 11)
        ok = sum(1 for r in rows if r["결과"] == "적합")
        ng = sum(1 for r in rows if r["결과"] == "부적합")
        avg = round(sum(r["점수"] for r in rows) / len(rows), 2)
        truth[stop] = (len(rows), ok, ng, avg)

        # 사진
        pdir = os.path.join(PHOTO_ROOT, info["date"].strftime("%Y%m%d"))
        os.makedirs(pdir, exist_ok=True)
        for (s, shop, item), (_, memo, *_rest) in FAILS.items():
            if s != stop or (s, shop, item) == NO_PHOTO:
                continue
            fname = f"{stop}_{item.replace('·', '').replace(' ', '')}.jpg"
            make_photo(os.path.join(pdir, fname), f"{stop} · {shop}", item)
        if stop == EXTRA_PHOTO[0]:
            make_photo(os.path.join(pdir, f"{stop}_{EXTRA_PHOTO[1].replace(' ', '')}.jpg"),
                       stop, EXTRA_PHOTO[1])

    xlsx = os.path.join(BASE, "위생점검결과_2026_4Q.xlsx")
    wb.save(xlsx)

    # 과거 수작업 집계 (누리휴게소 부적합 건수를 일부러 1 틀리게 적음)
    mb = Workbook()
    ms = mb.active
    ms.title = "수작업집계"
    ms.append(["휴게소", "점검건수", "적합", "부적합", "평균점수"])
    for stop, (n, ok, ng, avg) in truth.items():
        if stop == "누리휴게소":
            ng += 1
        ms.append([stop, n, ok, ng, avg])
    mb.save(os.path.join(BASE, "수작업집계_2026_4Q.xlsx"))
    print("생성 완료:", BASE)


if __name__ == "__main__":
    main()
