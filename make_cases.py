"""검증 6사례 입력 파일 생성

입력양식의 '작성예시'(가상 실험 EXP-800)를 '실험기록' 시트에 옮긴 뒤 사례별로 오류를 하나씩 주입한다.
생성 위치: 검증사례/사례1_정상.xlsx ~ 사례6_수소0.xlsx
"""
import copy
import os

from openpyxl import load_workbook

FORM = "탄소영수증_입력양식.xlsx"
OUT_DIR = "검증사례"
NCOL = 10  # 실험기록 A~J

# 사례 정의: 주입 내용, 예상 처리, 주입 오류와 관련된 검사(탐지 대상 + 같은 입력에 의존하는 검사)
CASES = [
    dict(no="①", key="사례1_정상", title="정상", inject="작성예시 그대로",
         expect="모든 검사 통과, 영수증 출력, 수작업 계산과 상대오차 0.1% 이내",
         target=None, target_status=None, expect_state="계산 완료", related=[]),
    dict(no="②", key="사례2_단위오기", title="단위 오기", inject="반응 구간 전력 5,000 Wh를 5,000 kWh로 입력",
         expect="이상치 '확인 필요', 계산 보류, 수정 후 재검사·재계산하면 ①과 동일, 수정 이력 기록",
         target="전력 원단위 이상치", target_status="확인 필요", expect_state="계산 보류",
         related=["전력 원단위 이상치"]),
    dict(no="③", key="사례3_필수값누락", title="필수값 누락", inject="수소 생산량 칸 비움",
         expect="'자료 부족', 계산 보류, 필요한 자료 안내 문구 출력",
         target="필수값", target_status="자료 부족", expect_state="계산 보류",
         related=["필수값", "수소 생산량(기능단위)", "전력 원단위 이상치", "화학양론 일관성(H)"]),
    dict(no="④", key="사례4_중복기록", title="중복 기록", inject="반응 구간 전력 행 2번 입력",
         expect="'확인 필요', 자동 삭제하지 않음, 계산 보류",
         target="중복 기록", target_status="확인 필요", expect_state="계산 보류",
         related=["중복 기록", "전력 원단위 이상치"]),
    dict(no="⑤", key="사례5_수지불일치", title="수지 불일치", inject="고체탄소 회수량을 이론량의 50%로 입력",
         expect="'확인 필요' 경고와 함께 계산 수행, 영수증에 경고와 가능 원인 표시",
         target="화학양론 일관성(C)", target_status="확인 필요", expect_state="계산 완료",
         related=["화학양론 일관성(C)"]),
    dict(no="⑥", key="사례6_수소0", title="수소 0", inject="수소 생산량 0 입력",
         expect="기능단위 계산 불가로 '계산 불가' 처리(0으로 나누기 방지), 사유 표시",
         target="수소 생산량(기능단위)", target_status="계산 불가", expect_state="계산 불가",
         related=["수소 생산량(기능단위)", "전력 원단위 이상치", "화학양론 일관성(H)"]),
]


# 추가 검증(리뷰 B3): 탐지 평가(탐지율·오경보율)에는 넣지 않고 따로 판정한다.
EXTRA_CASES = [
    dict(no="⑦", key="사례7_단위표기", title="같은 전력의 kWh 표기",
         inject="①과 같은 전력을 Wh 대신 kWh로 입력(1.2 / 5.0 / 0.3 kWh)",
         expect="모든 검사 통과, 결과가 ①과 같음(상대오차 0.1% 이내)"),
    dict(no="⑧", key="사례8_범위밖정상", title="범위 밖이지만 맞는 전력 기록",
         inject="전력을 승온 6,000 / 반응 25,000 / 중단 1,500 Wh로 입력(원단위 130 kWh/kg-H₂, 범위 5~100 밖). "
                "원자료가 맞다고 가정",
         expect="처음엔 '확인 필요'로 보류 → 연구자 확인 표시 후 경고를 유지한 채 계산, 값을 바꾸면 확인 표시 무효"),
]
ELEC_CASE8 = {"승온": 6000, "반응": 25000, "중단": 1500}


def example_rows(wb):
    ws = wb["작성예시"]
    rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, NCOL + 1)]
        if vals[0]:
            rows.append(vals)
    return rows


def settings_num(wb, sid):
    ws = wb["계산설정"]
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value == sid:
            return float(ws.cell(r, 4).value)
    raise KeyError(sid)


def find(rows, seg, code):
    for i, v in enumerate(rows):
        if v[1] == seg and v[3] == code:
            return i
    raise KeyError((seg, code))


def write_case(rows, path):
    wb = load_workbook(FORM)
    ws = wb["실험기록"]
    style_src = [ws.cell(2, c) for c in range(1, NCOL + 1)]
    for r in range(2, max(ws.max_row, len(rows) + 1) + 1):
        for c in range(1, NCOL + 1):
            ws.cell(r, c).value = None
    for i, vals in enumerate(rows, start=2):
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(i, c, v)
            src = style_src[c - 1]
            cell.font, cell.fill = copy.copy(src.font), copy.copy(src.fill)
            cell.border, cell.alignment = copy.copy(src.border), copy.copy(src.alignment)
    wb.save(path)


def build():
    os.makedirs(OUT_DIR, exist_ok=True)
    wb = load_workbook(FORM)
    base = example_rows(wb)
    paths = {}
    for case in CASES:
        rows = [list(r) for r in base]
        k = case["key"]
        if k == "사례2_단위오기":
            rows[find(rows, "반응", "ELEC")][5] = "kWh"          # 5000 Wh → 5000 kWh
        elif k == "사례3_필수값누락":
            rows[find(rows, "반응", "H2_PROD")][4] = None
        elif k == "사례4_중복기록":
            i = find(rows, "반응", "ELEC")
            rows.insert(i + 1, list(rows[i]))
        elif k == "사례5_수지불일치":
            # 이론 탄소량 = 반응 메탄 × M_C / M_CH4 (몰질량은 계산설정에서 읽음)
            cin = rows[find(rows, "반응", "CH4_IN")][4]
            cout = rows[find(rows, "반응", "CH4_OUT")][4]
            th_kg = (cin - cout) * settings_num(wb, "M_C") / settings_num(wb, "M_CH4")
            i = find(rows, "반응", "C_REC")
            rows[i][4], rows[i][5] = round(th_kg * 0.5 * 1000, 1), "g"
        elif k == "사례6_수소0":
            rows[find(rows, "반응", "H2_PROD")][4] = 0
        p = os.path.join(OUT_DIR, f"{k}.xlsx")
        write_case(rows, p)
        paths[k] = p
    for case in EXTRA_CASES:
        rows = [list(r) for r in base]
        k = case["key"]
        for seg in ("승온", "반응", "중단"):
            i = find(rows, seg, "ELEC")
            if k == "사례7_단위표기":
                rows[i][4], rows[i][5] = rows[i][4] / 1000, "kWh"
            else:
                rows[i][4] = ELEC_CASE8[seg]
        p = os.path.join(OUT_DIR, f"{k}.xlsx")
        write_case(rows, p)
        paths[k] = p
    return paths


if __name__ == "__main__":
    for k, p in build().items():
        print(k, "→", p)
