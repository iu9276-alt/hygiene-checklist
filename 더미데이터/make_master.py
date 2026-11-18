r"""웹앱용 점검 기준정보 더미 엑셀 생성 (PRD 3-1 형식)

출력 (웹앱 폴더 안 sample\):
  checklist_master_dummy.xlsx  정상 파일: 휴게소 3곳 x 매장 2개, 점검항목 8개, 1~5점
  checklist_master_bad.xlsx    검사 기능 시험용: 일부러 오류를 넣은 파일
휴게소·매장·담당자·연락처는 모두 가상 값이다 (make_dummy.py와 같은 값).
"""
import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "현장체크리스트", "sample")

SHOPS = [
    ("가람휴게소", "한식당", "김민수", "010-2481-3579"),
    ("가람휴게소", "분식코너", "이서연", "010-3692-4815"),
    ("누리휴게소", "푸드코트", "박지훈", "010-5813-2468"),
    ("누리휴게소", "카페", "최유진", "010-7924-1357"),
    ("다솜휴게소", "한식당", "정하늘", "010-4826-9137"),
    ("다솜휴게소", "편의점", "남궁민", "010-6357-2849"),
]
ITEMS = ["조리장 청결", "냉장·냉동 온도", "식재료 유통기한", "종사자 위생복장",
         "손씻기 시설", "방충·방서", "화장실 청결", "폐기물 관리"]


def style(ws, widths):
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def build(path, shops, items, shop_header=("휴게소", "입점매장", "매장담당자", "연락처")):
    wb = Workbook()
    ws = wb.active
    ws.title = "매장"
    ws.append(list(shop_header))
    for row in shops:
        ws.append(list(row))
    style(ws, {"A": 14, "B": 14, "C": 12, "D": 16})
    wi = wb.create_sheet("점검항목")
    wi.append(["순서", "점검항목", "점수척도"])
    for row in items:
        wi.append(list(row))
    style(wi, {"A": 8, "B": 18, "C": 10})
    wb.save(path)


def main():
    os.makedirs(OUT, exist_ok=True)
    build(os.path.join(OUT, "checklist_master_dummy.xlsx"), SHOPS,
          [(i + 1, name, "1~5") for i, name in enumerate(ITEMS)])

    # 검사 기능 시험용 오류:
    #  - '연락처' 열 없음 (경고)          - 입점매장 빈칸 (오류)
    #  - 같은 휴게소·매장 중복 (오류)     - 점수척도 '상중하' (오류)
    #  - 같은 점검항목 중복 (오류)
    bad_shops = [r[:3] for r in SHOPS[:3]] + [("누리휴게소", None, "홍길동"), ("가람휴게소", "한식당", "김민수")]
    bad_items = [(1, "조리장 청결", "1~5"), (2, "냉장·냉동 온도", "상중하"), (3, "조리장 청결", "1~5")]
    build(os.path.join(OUT, "checklist_master_bad.xlsx"), bad_shops, bad_items,
          shop_header=("휴게소", "입점매장", "매장담당자"))
    print("생성 완료:", os.path.abspath(OUT))


if __name__ == "__main__":
    main()
