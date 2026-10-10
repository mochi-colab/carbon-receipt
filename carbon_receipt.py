"""탄소영수증 시범 구현 (초기 개발 범위)

표준 Excel 입력 → 단위 환산 및 검사 → 수소 1 kg 기준 배출량 계산 → 배출원별 기여도와 탄소영수증 출력.

- 입력 양식의 '실험기록', '계산설정', '배출계수' 시트를 직접 읽는다. 수치(배출계수, 허용범위, 몰질량,
  단위환산, 경계·가정 문구)는 코드에 두지 않고 모두 양식에서 읽는다.
- 실험 간 순위 비교, AI 기능은 이 범위에 포함하지 않는다.
- 현재 양식의 배출계수와 실험값은 가상·예시값이다.

사용: python carbon_receipt.py 입력.xlsx -o 출력폴더
"""
import argparse
import csv
import json
import math
import os
import sys
import textwrap
from collections import OrderedDict, defaultdict

from openpyxl import load_workbook

ENGINE_VERSION = "v0.3"

PASS, CHECK, LACK, IMPOSSIBLE, PENDING = "통과", "확인 필요", "자료 부족", "계산 불가", "판정 보류"
ACKED = "확인 필요(연구자 확인)"  # 경고는 유지하되 계산은 진행
STATE_DONE, STATE_HOLD, STATE_IMPOSSIBLE = "계산 완료", "계산 보류", "계산 불가"

CHECK_NAMES = [
    "필수값", "중복 기록", "단위 인식", "값 범위·물리 한계", "배출계수 매칭",
    "수소 생산량(기능단위)", "전력 원단위 이상치", "화학양론 일관성(H)", "화학양론 일관성(C)",
]
# 검사별로 결과가 '확인 필요/자료 부족'일 때 따르는 계산설정 규칙
# (검사 → (규칙ID, 규칙이 적용되는 상태)). 다른 검사의 문제로 생긴 파생 상태(자료 부족·판정 보류)는
# 원인 검사의 규칙만 따른다.
CHECK_RULE = {
    "필수값": [("RULE_MISSING", "자료 부족")],
    "중복 기록": [("RULE_DUPLICATE", "확인 필요")],
    "단위 인식": [("RULE_UNIT_UNKNOWN", "확인 필요")],
    "값 범위·물리 한계": [("RULE_RANGE", "확인 필요")],
    "배출계수 매칭": [("RULE_EF_MISSING", "자료 부족"), ("RULE_EF_INVALID", "확인 필요")],
    "전력 원단위 이상치": [("RULE_ELEC_OUTLIER", "확인 필요")],
    "화학양론 일관성(H)": [("RULE_STOICH", "확인 필요")],
    "화학양론 일관성(C)": [("RULE_STOICH", "확인 필요")],
}
# 탐지율·오경보율은 처음 정한 8개 검사 기준으로 유지한다(3차 검토 후 추가한 검사는 별도 표로 확인)
ORIGINAL_CHECKS = [n for n in CHECK_NAMES if n != "값 범위·물리 한계"]


class FormError(Exception):
    """양식 구조 자체가 맞지 않을 때"""


# ----------------------------------------------------------------------------- 읽기
def _rows_by_header(ws):
    header = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(header) if h}
    for r in ws.iter_rows(min_row=2):
        yield r[0].row, {h: r[i].value if i < len(r) else None for h, i in idx.items()}


def _to_number(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return s  # 텍스트 값 (예: 촉매 종류)


def _blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


class Settings:
    def __init__(self, wb):
        if "계산설정" not in wb.sheetnames:
            raise FormError("'계산설정' 시트가 없습니다.")
        self.rows = OrderedDict()
        for _, d in _rows_by_header(wb["계산설정"]):
            sid = d.get("설정ID")
            if sid:
                self.rows[str(sid).strip()] = d
        if "FU" in self.rows and not (self.num("FU") > 0):
            raise FormError("계산설정의 기능단위(FU)는 0보다 커야 합니다.")

    def val(self, sid):
        if sid not in self.rows:
            raise FormError(f"계산설정에 '{sid}' 행이 없습니다.")
        return self.rows[sid]["값"]

    def num(self, sid):
        v = _to_number(self.val(sid))
        if not isinstance(v, float):
            raise FormError(f"계산설정 '{sid}' 값이 숫자가 아닙니다.")
        return v

    def group(self, prefix):
        return [(k, d) for k, d in self.rows.items() if k.startswith(prefix)]

    def items(self, sid):
        """쉼표로 나열된 설정값 → 목록 (행이 없으면 빈 목록)"""
        if sid not in self.rows:
            return []
        return [x.strip() for x in str(self.val(sid)).split(",") if x.strip()]

    def range_of(self, code):
        sid = f"RANGE_{code}"
        if sid not in self.rows:
            return None
        lo, hi = str(self.val(sid)).split("~")
        return float(lo), float(hi)

    def rule_holds(self, rule_id):
        return "보류" in str(self.val(rule_id))

    def std_unit(self, code):
        sid = f"STD_{code}"
        return str(self.val(sid)).strip() if sid in self.rows else None

    def unit_factor(self, unit):
        """입력단위 → (환산계수, 표준단위). 환산표에 없으면 None"""
        for k, d in self.group("UNIT_"):
            if str(d["항목"]).strip() == unit:
                return float(_to_number(d["값"])), str(d["단위"]).strip()
        return None

    def segs(self, sid):
        return [s.strip() for s in str(self.val(sid)).split("+") if s.strip()]


class EFTable(dict):
    """적용 항목코드 → 계수. 같은 버전에 같은 항목코드가 두 번 이상 있으면 dups에 기록"""
    def __init__(self):
        super().__init__()
        self.dups = set()


def load_ef(wb, version):
    if "배출계수" not in wb.sheetnames:
        raise FormError("'배출계수' 시트가 없습니다.")
    ef = EFTable()
    for _, d in _rows_by_header(wb["배출계수"]):
        if d.get("계수ID") and str(d.get("버전")).strip() == version:
            key = str(d.get("적용 항목코드")).strip()
            if key in ef:
                ef.dups.add(key)
            _v = _to_number(d["값"])
            ef[key] = {"id": d["계수ID"], "target": d["대상"],
                       "value": _v if isinstance(_v, float) else float("nan"),  # 숫자가 아니면 nan → 유한수 검사에서 보류
                       "unit": d["단위"], "source": d["출처"], "version": d["버전"],
                       "base_year": d.get("기준연도"), "condition": d.get("적용 조건")}
    return ef


def load_records(wb, sheet):
    if sheet not in wb.sheetnames:
        raise FormError(f"'{sheet}' 시트가 없습니다.")
    recs = []
    for rown, d in _rows_by_header(wb[sheet]):
        if _blank(d.get("구간")) and _blank(d.get("항목코드")):
            continue
        if _blank(d.get("실험ID")):
            continue  # 빈 양식 행
        recs.append({
            "row": rown, "exp": str(d["실험ID"]).strip(), "seg": str(d["구간"]).strip(),
            "name": d.get("항목"), "code": str(d["항목코드"]).strip(), "raw": d.get("값"),
            "unit": "" if _blank(d.get("단위")) else str(d["단위"]).strip(),
            "desc": d.get("측정 구간 설명") or "", "req": (d.get("필수/선택") or "").strip(),
            "note": d.get("비고") or "",
        })
    return recs


def load_acks(wb):
    """'검토확인' 시트: 연구자가 원자료를 대조해 값이 맞다고 표시한 검사 (correct_record.py --ack 로 기록)"""
    if "검토확인" not in wb.sheetnames:
        return []
    return [d for _, d in _rows_by_header(wb["검토확인"]) if d.get("실험ID")]


def load_history(wb):
    if "수정이력" not in wb.sheetnames:
        return []
    return [d for _, d in _rows_by_header(wb["수정이력"]) if d.get("실험ID")]


# ----------------------------------------------------------------------------- 값 정리
class Values:
    """(구간, 항목코드)별 표준단위 값과 문제(누락·중복·단위)를 관리"""

    def __init__(self, recs, S):
        self.by_key = defaultdict(list)
        for r in recs:
            self.by_key[(r["seg"], r["code"])].append(r)
        self.val, self.issue, self.meta = {}, {}, {}
        self.missing_required, self.dups, self.unit_bad = [], [], []
        for key, rs in self.by_key.items():
            self.meta[key] = rs[0]
            if len(rs) > 1:
                self.issue[key] = "dup"
                self.dups.append((key, rs))
                continue
            r = rs[0]
            v = _to_number(r["raw"])
            if v is None:
                self.issue[key] = "missing"
                if r["req"] == "필수":
                    self.missing_required.append(r)
                continue
            if isinstance(v, str):
                std = S.std_unit(r["code"])
                if std not in (None, "-"):  # 숫자 항목에 텍스트가 들어온 경우
                    self.issue[key] = "unit"
                    self.unit_bad.append((r, "숫자가 아닌 값"))
                    continue
                self.val[key] = v
                continue
            std = S.std_unit(r["code"])
            if std is None:
                self.issue[key] = "unknown"  # 등록되지 않은 항목코드: '값 범위·물리 한계'에서 표시
                continue
            if r["unit"] == std:
                self.val[key] = v
                continue
            fu = S.unit_factor(r["unit"])
            if fu is None or fu[1] != std:
                self.issue[key] = "unit"
                why = "단위환산표에 없는 단위" if fu is None else f"표준단위({std})와 차원이 다름"
                self.unit_bad.append((r, why))
                continue
            self.val[key] = v * fu[0]

    def get(self, code, seg=None):
        """(값, 문제) 반환. seg=None이면 모든 구간 합(텍스트는 단일값)."""
        keys = [k for k in self.by_key if k[1] == code and (seg is None or k[0] == seg)]
        if not keys:
            return None, "absent"
        for pri in ("dup", "unit", "missing"):
            if any(self.issue.get(k) == pri for k in keys):
                return None, pri
        vals = [self.val[k] for k in keys]
        if len(vals) == 1:
            return vals[0], None
        return sum(vals), None

    def any_row(self, code, seg=None):
        for k, rs in self.by_key.items():
            if k[1] == code and (seg is None or k[0] == seg):
                return rs[0]
        return None


def _elec_ack_key(V, base_segs):
    """전력 원단위 계산에 쓰이는 입력값(구간별 전력, 수소 생산량)을 표준단위로 나열한 문자열.
    연구자 확인 표시는 이 값에 연결되며, 하나라도 바뀌면 무효가 된다."""
    parts = []
    for s in base_segs:
        v, iss = V.get("ELEC", s)
        parts.append(f"{s}·ELEC={'?' if iss else format(v, '.9g')} kWh")
    v, iss = V.get("H2_PROD")
    parts.append(f"H2_PROD={'?' if iss else format(v, '.9g')} kg")
    return "; ".join(parts)


ACK_KEY_FUNCS = {"전력 원단위 이상치": _elec_ack_key}


def _issue_status(issue):
    return LACK if issue in ("missing", "absent") else PENDING


def _issue_text(issue):
    return {"missing": "값 없음", "absent": "행 없음", "dup": "중복 기록 미해결", "unit": "단위 문제 미해결"}[issue]


# ----------------------------------------------------------------------------- 평가
def evaluate(exp, recs, S, EF, history, acks=()):
    V = Values(recs, S)
    checks = OrderedDict((n, (PASS, "")) for n in CHECK_NAMES)
    msgs = []
    base_segs, ref_segs = S.segs("ENERGY_SCOPE_BASE"), S.segs("ENERGY_SCOPE_REF")
    ch4_code = str(S.val("CH4_BASIS")).strip()
    num_code, den_code = [x.strip() for x in str(S.val("CAT_ALLOC")).split("/")]

    # 1) 필수값 (+ 계산에 필요한데 행 자체가 없는 경우)
    lacking = list(V.missing_required)
    needed = [(s, "ELEC") for s in base_segs] + [(s, "HEAT_EXT") for s in base_segs]
    needed += [(None, c) for c in (ch4_code, "H2_PROD", "CAT_TYPE", num_code, den_code)]
    for seg, code in needed:
        if V.get(code, seg)[1] == "absent":
            lacking.append({"seg": seg or "-", "code": code, "name": code, "desc": "양식에 해당 행이 없음", "unit": ""})
    # 계산설정의 필수 행 목록: 입력 파일의 필수/선택 표시와 별도로 행 존재와 공란을 모두 검사
    seen = {(r["seg"], r["code"]) for r in lacking}
    for item in S.items("REQUIRED_ROWS"):
        seg, code = [x.strip() for x in item.split(":")]
        iss = V.get(code, seg)[1]
        if iss in ("absent", "missing") and (seg, code) not in seen:
            seen.add((seg, code))
            if iss == "absent":
                lacking.append({"seg": seg, "code": code, "name": code, "desc": "필수 행이 양식에 없음(행 삭제 의심)",
                                "unit": ""})
            else:
                lacking.append(V.meta[(seg, code)])
    if lacking:
        checks["필수값"] = (LACK, f"{len(lacking)}개 항목 누락")
        for r in lacking:
            std = S.std_unit(r["code"]) or ""
            msgs.append(f"[자료 부족] {r['seg']} 구간 '{r['name']}'({r['code']}) 값을 입력해야 계산할 수 있음 "
                        f"— 측정 구간: {r['desc']} (표준단위 {std})")

    # 2) 중복 기록
    if V.dups:
        parts = []
        for (seg, code), rs in V.dups:
            rows = ", ".join(str(r["row"]) for r in rs)
            parts.append(f"{seg}·{rs[0]['name']}({code}) {len(rs)}회")
            msgs.append(f"[확인 필요] 중복 기록: {seg} 구간 '{rs[0]['name']}'이 {rows}행에 {len(rs)}회 입력됨 "
                        f"— 자동 삭제하지 않음. 원자료 확인 후 1행만 남기고 재검사")
        checks["중복 기록"] = (CHECK, "; ".join(parts))

    # 3) 단위 인식
    if V.unit_bad:
        checks["단위 인식"] = (CHECK, f"{len(V.unit_bad)}개 항목")
        for r, why in V.unit_bad:
            msgs.append(f"[확인 필요] {r['seg']} 구간 '{r['name']}' 단위 '{r['unit']}' — {why}")

    # 3-2) 값 범위·물리 한계 (규칙은 계산설정에서 읽음)
    viol = []
    allowed_segs = S.items("ALLOWED_SEGMENTS")
    nonneg = set(S.items("NONNEG_CODES"))
    ints = set(S.items("INT_MIN1_CODES"))
    for (seg, code), rs in V.by_key.items():
        r = rs[0]
        label = f"{seg}·{r['name']}({code})"
        if allowed_segs and seg not in allowed_segs:
            viol.append((label, f"등록되지 않은 구간 '{seg}' — 집계에서 빠짐"))
            continue
        if V.issue.get((seg, code)) == "unknown":
            viol.append((label, "등록되지 않은 항목코드 — 계산에 쓰이지 않음"))
            continue
        v = V.val.get((seg, code))
        if not isinstance(v, float):
            continue
        if not math.isfinite(v):
            viol.append((label, f"{r['raw']} — 유한한 수가 아님(nan·inf)"))
            continue
        if code in nonneg and v < 0:
            viol.append((label, f"{_fmt(_to_number(r['raw']))} {r['unit']} — 음수가 될 수 없는 항목"))
        rg = S.range_of(code)
        if rg and not (rg[0] <= v <= rg[1]):
            viol.append((label, f"{_fmt(v)} {S.std_unit(code)} — 허용 범위 {rg[0]:g}~{rg[1]:g} 밖"))
        if code in ints and (v < 1 or abs(v - round(v)) > 1e-9):
            viol.append((label, f"{_fmt(v)} — 1 이상의 정수여야 함"))
    for item in S.items("POSITIVE_ROWS"):
        seg, code = [x.strip() for x in item.split(":")]
        v, iss = V.get(code, seg)
        if not iss and isinstance(v, float) and v <= 0:
            viol.append((f"{seg}·{V.meta[(seg, code)]['name']}({code})", f"{_fmt(v)} — 0보다 커야 함"))
    if "SUM_MAX_X" in S.rows:
        xs = [V.get(c, "반응") for c in ("X_CH4", "X_H2")]
        if all(not i and isinstance(v, float) for v, i in xs) and sum(v for v, _ in xs) > S.num("SUM_MAX_X") + 1e-9:
            viol.append(("반응·출구 가스 조성", f"CH4 + H2 몰분율 합 {_fmt(sum(v for v, _ in xs))} — "
                                          f"{S.num('SUM_MAX_X'):g}을 넘음"))
    _in, _ii = V.get("CH4_IN")
    _out, _oi = V.get("CH4_OUT")
    if not (_ii or _oi) and isinstance(_in, float) and isinstance(_out, float) and _out > _in:
        viol.append(("반응·메탄 투입/배출", f"미반응 메탄 {_fmt(_out)} kg이 투입량 {_fmt(_in)} kg보다 많음"))
    _new, _ni = V.get(ch4_code)
    _h2, _hi = V.get("H2_PROD")
    if not (_ni or _hi) and isinstance(_new, float) and isinstance(_h2, float) and _h2 > 0:
        need_min = _h2 * S.num("M_CH4") / (2 * S.num("M_H2")) * (1 - S.num("STOICH_TOL"))
        if _new < need_min:
            viol.append(("반응·메탄 신규 공급량", f"{_fmt(_new)} kg — 수소 {_fmt(_h2)} kg에 필요한 이론 최소량 "
                                             f"{_fmt(need_min)} kg(허용오차 반영)보다 적음"))
    if viol:
        checks["값 범위·물리 한계"] = (CHECK, f"{len(viol)}건: " + "; ".join(l for l, _ in viol)[:60])
        for l, why in viol:
            msgs.append(f"[확인 필요] 값 범위: {l} {why}. 원자료 확인 후 수정·재검사")

    # 4) 배출계수 매칭
    cat_type, cat_issue = V.get("CAT_TYPE")
    ef_needed = {"ELEC": "전력", "HEAT_EXT": "외부 공급열", ch4_code: "원료 메탄"}
    ef_used = {}
    ef_missing = [lab for code, lab in ef_needed.items() if code not in EF]
    for code in ef_needed:
        if code in EF:
            ef_used[code] = EF[code]
    if cat_issue:
        checks["배출계수 매칭"] = (_issue_status(cat_issue), f"촉매 종류 {_issue_text(cat_issue)}")
    else:
        ck = f"{num_code}:{cat_type}"
        if ck in EF:
            ef_used[ck] = EF[ck]
        else:
            ef_missing.append(f"촉매({cat_type})")
    if ef_missing:
        checks["배출계수 매칭"] = (LACK, "없음: " + ", ".join(ef_missing))
        msgs.append(f"[자료 부족] 배출계수 버전 {S.val('EF_VERSION')}에 다음 계수가 없음: {', '.join(ef_missing)}")
    else:
        # 계수 단위(분모가 활동량 표준단위와 맞는지), 음수 계수, 같은 버전 중복 — 자동 환산하지 않음
        bad = []
        for code, e in ef_used.items():
            base = code.split(":")[0]
            want = str(S.val(f"EFU_{base}")).replace(" ", "") if f"EFU_{base}" in S.rows else None
            have = str(e["unit"]).replace(" ", "")
            if want is None or have != want:
                bad.append(f"{e['id']} 단위 '{e['unit']}'(필요: {want or '정의 없음'})")
            if not math.isfinite(e["value"]):
                bad.append(f"{e['id']} 값이 유한한 수가 아님(nan·inf 또는 문자)")
            elif e["value"] < 0:
                bad.append(f"{e['id']} 값 {e['value']:g}(음수)")
            if code in getattr(EF, "dups", set()):
                bad.append(f"항목코드 {code}에 같은 버전 계수가 2개 이상")
        if bad:
            checks["배출계수 매칭"] = (CHECK, "; ".join(bad)[:60])
            for b in bad:
                msgs.append(f"[확인 필요] 배출계수: {b}. 지원하지 않는 단위는 자동 환산하지 않고 계산 보류")

    # 5) 수소 생산량(기능단위)
    h2, h2_issue = V.get("H2_PROD")
    if h2_issue:
        checks["수소 생산량(기능단위)"] = (_issue_status(h2_issue), _issue_text(h2_issue))
    elif h2 <= 0:
        checks["수소 생산량(기능단위)"] = (IMPOSSIBLE, f"수소 생산량 {h2:g} kg")
        msgs.append(f"[계산 불가] 수소 생산량이 {h2:g} kg(0 이하)이므로 기능단위(1 kg-H2)당 배출량을 계산할 수 없음 "
                    f"— 0으로 나누기 방지. 측정 누락·단위 오기인지, 실제로 수소가 생성되지 않았는지 확인")

    # 6) 전력 원단위 (단위 오기 의심)
    seg_kwh, e_issue = {}, None
    for s in base_segs:
        v, iss = V.get("ELEC", s)
        if iss:
            e_issue = e_issue or iss
        else:
            seg_kwh[s] = v
    lo, hi = S.num("ELEC_INTENSITY_MIN"), S.num("ELEC_INTENSITY_MAX")
    if e_issue or h2_issue:
        iss = e_issue or h2_issue
        checks["전력 원단위 이상치"] = (_issue_status(iss), _issue_text(iss))
    elif h2 <= 0:
        checks["전력 원단위 이상치"] = (PENDING, "수소 생산량 0 이하")
    else:
        inten = sum(seg_kwh.values()) / h2
        if lo <= inten <= hi:
            checks["전력 원단위 이상치"] = (PASS, f"{inten:,.1f} kWh/kg-H2")
        else:
            checks["전력 원단위 이상치"] = (CHECK, f"{inten:,.0f} kWh/kg-H2 (범위 {lo:g}~{hi:g} 밖)")
            seg_txt = ", ".join(f"{s} {v:,.1f} kWh" for s, v in seg_kwh.items())
            msgs.append(f"[확인 필요] 전력 원단위 {inten:,.0f} kWh/kg-H2가 범위({lo:g}~{hi:g})를 벗어남 — 단위 오기 의심. "
                        f"구간별: {seg_txt}. 원본 로그와 대조 후 수정·재검사")

    # 7) 화학양론 H
    tol = S.num("STOICH_TOL")
    mch4, mh2, mc = S.num("M_CH4"), S.num("M_H2"), S.num("M_C")
    cin, i1 = V.get("CH4_IN")
    cout, i2 = V.get("CH4_OUT")
    react_issue = i1 or i2
    reacted = None if react_issue else cin - cout
    if react_issue or h2_issue:
        iss = react_issue or h2_issue
        checks["화학양론 일관성(H)"] = (_issue_status(iss), _issue_text(iss))
    elif reacted <= 0:
        checks["화학양론 일관성(H)"] = (CHECK, "반응 메탄량 0 이하")
    else:
        th = reacted * 2 * mh2 / mch4
        ratio = h2 / th
        st = PASS if abs(ratio - 1) <= tol else CHECK
        checks["화학양론 일관성(H)"] = (st, f"기록/이론 {ratio:.1%} (허용 ±{tol:.0%})")
        if st == CHECK:
            msgs.append(f"[확인 필요] 수소 생산량 {h2:.4f} kg이 이론량 {th:.4f} kg의 {ratio:.1%} — "
                        f"가능 원인: 입력 오류 / 미반응 메탄 기록 오류 / 부생 탄화수소 생성")

    # 8) 화학양론 C
    crec, c_issue = V.get("C_REC")
    cacc, a_issue = V.get("C_ACC")
    if react_issue or c_issue:
        iss = react_issue or c_issue
        checks["화학양론 일관성(C)"] = (_issue_status(iss), _issue_text(iss))
    elif a_issue in ("dup", "unit"):
        checks["화학양론 일관성(C)"] = (PENDING, "축적량 " + _issue_text(a_issue))
    elif reacted <= 0:
        checks["화학양론 일관성(C)"] = (CHECK, "반응 메탄량 0 이하")
    else:
        th = reacted * mc / mch4
        rec = crec + (cacc if not a_issue else 0.0)
        ratio = rec / th
        basis = "회수+축적" if not a_issue else "회수량만(축적량 미기록)"
        st = PASS if abs(ratio - 1) <= tol else CHECK
        checks["화학양론 일관성(C)"] = (st, f"{basis}/이론 {ratio:.1%} (허용 ±{tol:.0%})")
        if st == CHECK:
            msgs.append(f"[확인 필요·계산 수행] 고체탄소 {basis} {rec:.4f} kg이 이론량 {th:.4f} kg의 {ratio:.1%} — "
                        f"가능 원인: 입력 오류 / 반응기 내 축적 / 미측정 유출. 질량 배분 참고값 해석에 주의")

    # ---------------- 연구자 확인 표시 반영 (허용된 검사만, 확인 당시 값과 지금 값이 같을 때만)
    allowed = [x.strip() for x in str(S.val("ACK_ALLOWED")).split(",")] if "ACK_ALLOWED" in S.rows else []
    acked = []
    for a in acks:
        if str(a.get("실험ID")).strip() != exp:
            continue
        name = str(a.get("검사")).strip()
        if name not in checks or checks[name][0] != CHECK:
            continue
        if name not in allowed:
            msgs.append(f"[확인 표시 무시] '{name}'은 연구자 확인으로 넘어갈 수 없는 검사임")
            continue
        key_now = ACK_KEY_FUNCS[name](V, base_segs) if name in ACK_KEY_FUNCS else checks[name][1]
        key_then = str(a.get("확인 당시 입력값") or "").strip()
        if key_then != key_now:
            msgs.append(f"[확인 표시 무효] '{name}' 확인 이후 계산 입력값이 바뀜 "
                        f"(확인 당시: {key_then or '기록 없음'} / 현재: {key_now}). 다시 확인 필요")
            continue
        checks[name] = (ACKED, checks[name][1])
        acked.append(name)
        msgs.append(f"[연구자 확인·경고 유지] {name}: {a.get('확인 내용')} ({a.get('확인자')}, {a.get('일시')}) "
                    f"— 계산은 진행하되 이 경고를 영수증에 남김")

    # ---------------- 계산 상태 판정
    hold = []
    for name, (st, _) in checks.items():
        for rid, trig in CHECK_RULE.get(name, []):
            if st == trig and S.rule_holds(rid) and name not in hold:
                hold.append(name)
    if checks["수소 생산량(기능단위)"][0] == IMPOSSIBLE:
        state, reason = STATE_IMPOSSIBLE, "수소 생산량 0 이하 — 기능단위 환산 불가"
    elif hold:
        state, reason = STATE_HOLD, "검사 미해결: " + ", ".join(hold)
    else:
        state, reason = STATE_DONE, ("경고 유지: " + ", ".join(acked) + "(연구자 확인)") if acked else ""
    if state == STATE_DONE:
        # 규칙 설정과 무관하게, 계산에 쓰는 값이 하나라도 확정되지 않았으면 계산하지 않는다(안전장치)
        need = [("ELEC", s) for s in set(base_segs + ref_segs)] + [("HEAT_EXT", s) for s in set(base_segs + ref_segs)]
        need += [(c, None) for c in (ch4_code, "H2_PROD", "CAT_TYPE", num_code, den_code)]
        unresolved = [f"{s or '전체'}·{c}({_issue_text(V.get(c, s)[1])})" for c, s in need if V.get(c, s)[1]]
        if not unresolved and f"{num_code}:{cat_type}" not in EF:
            unresolved.append(f"촉매({cat_type}) 배출계수 없음")
        if any(c not in EF for c in ("ELEC", "HEAT_EXT", ch4_code)):
            unresolved.append("배출계수 없음")
        if unresolved:
            state, reason = STATE_HOLD, "계산에 필요한 값 미확정: " + ", ".join(sorted(unresolved))

    # ---------------- 메타·가정·경계
    temp, _ = V.get("TEMP")
    res = OrderedDict(
        exp=exp, temp=temp, catalyst=cat_type, state=state, state_reason=reason,
        checks=checks, messages=msgs, engine=ENGINE_VERSION, form_version=S.val("FORM_VERSION"),
        ef_version=S.val("EF_VERSION"), fu=f"{S.num('FU'):g} {S.rows['FU']['단위']}",
        label_data=S.val("LABEL_DATA"), label_result=S.val("LABEL_RESULT"),
        boundary_in=[d["값"] for _, d in S.group("BOUNDARY_IN_")],
        boundary_out=[(d["값"], d["설명·근거"]) for _, d in S.group("BOUNDARY_OUT_")],
        history=[h for h in history if str(h.get("실험ID")).strip() == exp],
        ef_used=[{"code": k, **v} for k, v in ef_used.items()],
        items=OrderedDict(), total=None, ref_total=None, ref_items=OrderedDict(), alloc=None,
        alloc_note="", assumptions=[], intermediate=OrderedDict(), trace=[],
        acks=[a for a in acks if str(a.get("실험ID")).strip() == exp],
        ack_keys={n: f(V, base_segs) for n, f in ACK_KEY_FUNCS.items()},
        assumed=[f"{r['seg']}·{r['name']} {_fmt(_to_number(r['raw']))} {r['unit']}" for r in recs
                 if "[가정값]" in str(r.get("note", "")) and not _blank(r["raw"])],
    )
    if state != STATE_DONE:
        res["assumptions"] = _assumption_lines(S, None)
        return res

    # ---------------- 계산
    fu = S.num("FU")
    per = fu / h2

    def energy(segs):
        kwh = sum(V.get("ELEC", s)[0] for s in segs)
        mj = sum(V.get("HEAT_EXT", s)[0] for s in segs)
        return kwh, mj

    kwh_b, mj_b = energy(base_segs)
    kwh_r, mj_r = energy(ref_segs)
    ch4, _ = V.get(ch4_code)
    cat_mass, _ = V.get(num_code)
    cat_uses, _ = V.get(den_code)
    if cat_uses <= 0:
        res["state"], res["state_reason"] = STATE_HOLD, "촉매 총 사용 횟수 0 이하"
        return res
    cat_alloc = cat_mass / cat_uses
    efE, efH, efM = EF["ELEC"], EF["HEAT_EXT"], EF[ch4_code]
    efC = EF[f"{num_code}:{cat_type}"]

    def items_for(kwh, mj):
        return OrderedDict([
            (efE["target"], kwh * efE["value"] * per),
            (efH["target"], mj * efH["value"] * per),
            ("원료 " + efM["target"], ch4 * efM["value"] * per),
            (efC["target"], cat_alloc * efC["value"] * per),
        ])

    res["items"] = items_for(kwh_b, mj_b)
    res["total"] = sum(res["items"].values())
    res["ref_items"] = items_for(kwh_r, mj_r)
    res["ref_total"] = sum(res["ref_items"].values())
    if crec is not None and not c_issue and crec > 0:
        share = h2 / (h2 + crec)
        res["alloc"] = res["total"] * share
        res["alloc_note"] = f"수소 배분율 {share:.2%} = {h2:.4f} ÷ ({h2:.4f} + {crec:.4f})"
        if checks["화학양론 일관성(C)"][0] == CHECK:
            res["alloc_note"] += " · 탄소 점검 '확인 필요' 상태의 회수량 사용"
    else:
        res["alloc_note"] = "고체탄소 회수량 없음 → 산출하지 않음(추정값으로 채우지 않음)"
    res["intermediate"] = OrderedDict([
        ("수소 생산량 (kg)", h2), ("전력 합계-배치 전체 (kWh)", kwh_b), ("전력 합계-반응 구간 (kWh)", kwh_r),
        ("외부 공급열-배치 전체 (MJ)", mj_b), ("메탄 신규 공급량 (kg)", ch4),
        ("촉매 1회 배분량 (kg)", cat_alloc), ("전력 원단위-배치 전체 (kWh/kg-H2)", kwh_b / h2),
    ])
    res["trace"] = _trace(V, S, EF, base_segs, ch4_code, num_code, den_code, cat_type, h2, res["items"],
                          (efE, efH, efM, efC))
    res["assumptions"] = _assumption_lines(S, dict(ch4=ch4, cat_mass=cat_mass, cat_uses=cat_uses,
                                                   cat_alloc=cat_alloc))
    return res


def _fmt(v):
    return f"{v:,.6g}" if isinstance(v, float) else str(v)


def _trace(V, S, EF, base_segs, ch4_code, num_code, den_code, cat_type, h2, items, efs):
    """결과 한 줄마다 입력 행 → 원래 값·단위 → 환산값 → 계수 → 식 → 결과를 연결"""
    efE, efH, efM, efC = efs
    def rows(code, segs=None):
        rs = [r for (seg, c), lst in V.by_key.items() for r in lst if c == code and (segs is None or seg in segs)]
        return sorted(rs, key=lambda r: r["row"])
    def src(rs):
        return "실험기록 " + "·".join(str(r["row"]) for r in rs) + "행"
    def raw(rs):
        return " + ".join(f"{_fmt(_to_number(r['raw']))} {r['unit']}" for r in rs)
    h2r = rows("H2_PROD")
    h2_txt = f"수소 {raw(h2r)} → {_fmt(h2)} kg ({src(h2r)})"
    T = []
    def add(label, rs, std_val, std_unit, ef, formula, val):
        T.append(OrderedDict([("결과 항목", label), ("입력 행", src(rs)), ("원래 값·단위", raw(rs)),
                              ("환산값", f"{_fmt(std_val)} {std_unit}"), ("기능단위 기준", h2_txt),
                              ("계수", f"{ef['id']} {ef['value']:g} {ef['unit']}"), ("식", formula),
                              ("결과 (kgCO2-eq/kg-H2)", round(val, 6))]))
    er = rows("ELEC", base_segs); kwh = sum(V.val[(r['seg'], 'ELEC')] for r in er)
    add(efE["target"], er, kwh, "kWh", efE, f"{_fmt(kwh)} ÷ {_fmt(h2)} × {efE['value']:g}", items[efE["target"]])
    hr = rows("HEAT_EXT", base_segs); mj = sum(V.val[(r['seg'], 'HEAT_EXT')] for r in hr)
    add(efH["target"], hr, mj, "MJ", efH, f"{_fmt(mj)} ÷ {_fmt(h2)} × {efH['value']:g}", items[efH["target"]])
    mr = rows(ch4_code); ch4 = V.get(ch4_code)[0]
    add("원료 " + efM["target"], mr, ch4, "kg", efM, f"{_fmt(ch4)} ÷ {_fmt(h2)} × {efM['value']:g}",
        items["원료 " + efM["target"]])
    cr_ = rows(num_code) + rows(den_code); cm, cu = V.get(num_code)[0], V.get(den_code)[0]
    T.append(OrderedDict([("결과 항목", efC["target"]), ("입력 행", src(cr_)), ("원래 값·단위", raw(cr_).replace(" + ", ", ")),
                          ("환산값", f"{_fmt(cm)} kg ÷ {_fmt(cu)}회 = {_fmt(cm / cu)} kg/회"), ("기능단위 기준", h2_txt),
                          ("계수", f"{efC['id']} {efC['value']:g} {efC['unit']}"),
                          ("식", f"{_fmt(cm)} ÷ {_fmt(cu)} ÷ {_fmt(h2)} × {efC['value']:g}"),
                          ("결과 (kgCO2-eq/kg-H2)", round(items[efC["target"]], 6))]))
    return T


def _assumption_lines(S, n):
    base, ref = S.val("ENERGY_SCOPE_BASE"), S.val("ENERGY_SCOPE_REF")
    lines = [f"에너지 집계: 기본 = {base}(배치 전체), 참고 = {ref} 구간만",
             f"원료 메탄: {S.rows['CH4_BASIS']['설명·근거']}"]
    if n:
        lines[-1] += f" ({n['ch4']:.3f} kg)"
        lines.append(f"촉매 배분: 사용량 {n['cat_mass']:.4f} kg ÷ 총 사용 {n['cat_uses']:g}회(최초 포함) = {n['cat_alloc']:.4f} kg/회")
    else:
        lines.append(f"촉매 배분: {S.rows['CAT_ALLOC']['설명·근거']}")
    lines.append(f"고체탄소: 기본 = {S.val('CARBON_BASE')}, 참고 = {S.val('CARBON_REF')}(회수량 있을 때만)")
    return lines


# ----------------------------------------------------------------------------- 출력
def receipt_rows(res):
    """CSV·텍스트·그림이 공유하는 영수증 내용 (구분, 항목, 값, 단위, 비고)"""
    R = []
    add = lambda *a: R.append(list(a) + [""] * (5 - len(a)))
    add("표시", "자료 상태", res["label_data"])
    add("표시", "결과 성격", res["label_result"])
    add("식별", "실험ID", res["exp"])
    add("식별", "반응온도", "" if res["temp"] is None else f"{res['temp']:g}", "℃")
    add("식별", "촉매", res["catalyst"] or "")
    add("식별", "기능단위", res["fu"])
    add("버전", "계산 엔진", res["engine"])
    add("버전", "입력양식", res["form_version"])
    add("버전", "배출계수", res["ef_version"], "", "시연용 예시값")
    add("상태", "계산 상태", res["state"], "", res["state_reason"])
    for a in res["assumed"]:
        add("가정값 입력", a, "", "", "실측·보고값이 아닌 가정값(입력 행 비고 참조)")
    if res["state"] == STATE_DONE:
        for k, v in res["items"].items():
            add("배출 내역(배치 전체)", k, f"{v:.6f}", "kgCO2-eq/kg-H2",
                f"기여도 {v / res['total']:.2%}" if res["total"] else "기여도 -")
        add("합계", "배치 전체 · 부산물 공제 없음 (기본)", f"{res['total']:.6f}", "kgCO2-eq/kg-H2")
        add("참고값", "반응 구간만 · 부산물 공제 없음", f"{res['ref_total']:.6f}", "kgCO2-eq/kg-H2",
            "승온·중단 구간 에너지 제외")
        add("참고값", "배치 전체 · 질량 배분",
            "" if res["alloc"] is None else f"{res['alloc']:.6f}", "kgCO2-eq/kg-H2", res["alloc_note"])
        for k, v in res["intermediate"].items():
            add("중간값", k, f"{v:.6f}")
    for name, (st, det) in res["checks"].items():
        add("검사", name, st, "", det)
    for m in res["messages"]:
        add("경고·안내", "", m)
    for a in res["assumptions"]:
        add("적용 가정", "", a)
    for b in res["boundary_in"]:
        add("경계 안", "", b)
    for b, why in res["boundary_out"]:
        add("경계 밖", b, "", "", why)
    for e in res["ef_used"]:
        add("사용 배출계수", e["id"], f"{e['value']:g}", e["unit"], f"대상: {e['target']} | 버전: {e['version']} | 출처: {e['source']} | "
            f"기준연도: {e.get('base_year')} | 적용 조건: {e.get('condition')}")
    for t in res["trace"]:
        add("근거 추적", t["결과 항목"], f"{t['결과 (kgCO2-eq/kg-H2)']:.6f}", "kgCO2-eq/kg-H2",
            f"{t['입력 행']} | {t['원래 값·단위']} → {t['환산값']} | {t['기능단위 기준']} | {t['계수']} | {t['식']}")
    for a in res["acks"]:
        add("연구자 확인", a.get("검사"), a.get("확인 당시 판정값"), "", f"{a.get('일시')} · {a.get('확인자')} · {a.get('확인 내용')}")
    for h in res["history"]:
        add("수정 이력", f"{h.get('구간')}·{h.get('항목코드')}",
            f"{h.get('수정 전 값')} {h.get('수정 전 단위')} → {h.get('수정 후 값')} {h.get('수정 후 단위')}", "",
            f"{h.get('일시')} · {h.get('수정자')} · {h.get('사유')}")
    return R


def write_csv(res, path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["구분", "항목", "값", "단위", "비고"])
        w.writerows(receipt_rows(res))


def write_txt(res, path):
    L = []
    L.append(f"※ {res['label_data']}")
    L.append(f"※ {res['label_result']}")
    L.append("=" * 60)
    L.append("탄 소 영 수 증".center(56))
    temp_txt = "" if res["temp"] is None else "{:g}℃".format(res["temp"])
    L.append(f"{res['exp']} | {temp_txt} · 촉매 {res['catalyst'] or '-'}")
    L.append(f"엔진 {res['engine']} · 양식 {res['form_version']} · 배출계수 {res['ef_version']} · 기능단위 {res['fu']}")
    L.append("-" * 60)
    L.append(f"계산 상태: {res['state']}" + (f"  ({res['state_reason']})" if res["state_reason"] else ""))
    if res["assumed"]:
        L.append("※ 가정값 입력(실측·보고값 아님): " + "; ".join(res["assumed"]))
    if res["state"] == STATE_DONE:
        L.append("배출 내역 (kgCO2-eq/kg-H2, 배치 전체)")
        for k, v in res["items"].items():
            L.append(f"  {k:<20}{v:>9.2f}  ({(v / res['total']) if res['total'] else 0:5.1%})")
        L.append(f"  {'합계 · 부산물 공제 없음':<18}{res['total']:>9.2f}")
        L.append("참고값")
        L.append(f"  반응 구간만{'':<12}{res['ref_total']:>9.2f}")
        L.append(f"  질량 배분{'':<14}" + (f"{res['alloc']:>9.2f}" if res["alloc"] is not None else "   미산출")
                 + f"   [{res['alloc_note']}]")
    L.append("-" * 60)
    L.append("검사 항목별 상태")
    for name, (st, det) in res["checks"].items():
        L.append(f"  {name:<18}{st:<8} {det}")
    if res["messages"]:
        L.append("경고·안내")
        for m in res["messages"]:
            L.append("  · " + m)
    if res["history"]:
        L.append("수정 이력")
        for h in res["history"]:
            L.append(f"  · {h.get('구간')}·{h.get('항목코드')}: {h.get('수정 전 값')} {h.get('수정 전 단위')} → "
                     f"{h.get('수정 후 값')} {h.get('수정 후 단위')} ({h.get('사유')}, {h.get('수정자')}, {h.get('일시')})")
    if res["trace"]:
        L.append("결과 한 줄의 근거 (입력 행 → 원래 값 → 환산값 → 계수 → 식 → 결과)")
        for t in res["trace"]:
            L.append(f"  · {t['결과 항목']}: {t['입력 행']} [{t['원래 값·단위']}] → {t['환산값']}; "
                     f"{t['기능단위 기준']}; {t['계수']}; {t['식']} = {t['결과 (kgCO2-eq/kg-H2)']:.4f}")
    L.append("적용 가정")
    L += ["  · " + a for a in res["assumptions"]]
    L.append("경계 밖 항목")
    L += [f"  · {b}" for b, _ in res["boundary_out"]]
    L.append("사용 배출계수 (시연용 예시값)")
    for e in res["ef_used"]:
        L.append(f"  · {e['id']} {e['target']} {e['value']:g} {e['unit']} [{e['version']}]")
        L.append(f"    출처: {e['source']} / 기준연도: {e.get('base_year')} / 적용 조건: {e.get('condition')}")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")


STATUS_COLOR = {PASS: "#2e7d32", CHECK: "#e65100", ACKED: "#b26a00", LACK: "#616161", IMPOSSIBLE: "#b71c1c", PENDING: "#9e9e9e"}
STATE_COLOR = {STATE_DONE: "#2e7d32", STATE_HOLD: "#e65100", STATE_IMPOSSIBLE: "#b71c1c"}


def _font():
    from matplotlib import font_manager as fm
    import matplotlib.pyplot as plt
    for p in ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
              "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"):
        if os.path.exists(p):
            fm.fontManager.addfont(p)
    names = {f.name for f in fm.fontManager.ttflist}
    for n in ("Noto Sans CJK KR", "Noto Sans CJK JP", "NanumGothic", "Malgun Gothic", "AppleGothic"):
        if n in names:
            plt.rcParams["font.family"] = n
            break
    plt.rcParams["axes.unicode_minus"] = False


def write_png(res, path):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch
    _font()

    W = 6.4
    ops = []  # (kind, payload, height_in)
    wrap = lambda t, n: textwrap.wrap(t, n) or [""]

    def text(t, size=8.6, w="normal", c="#222", indent=0.0, n=62):
        for ln in wrap(t, n):
            ops.append(("text", (ln, size, w, c, indent), size / 72 * 1.45))

    def kv(k, v, w="normal", size=8.8, c="#222", vc=None):
        ops.append(("kv", (k, v, w, size, c, vc or c), size / 72 * 1.55))

    def rule():
        ops.append(("rule", None, 0.10))

    def gap(h=0.06):
        ops.append(("gap", None, h))

    ops.append(("title", "탄 소 영 수 증", 0.42))
    t = "" if res["temp"] is None else f"{res['temp']:g}℃ · "
    ops.append(("center", (f"{res['exp']}  |  {t}{res['catalyst'] or '-'} 촉매", 9, "#555"), 0.22))
    ops.append(("center", (f"엔진 {res['engine']} · 양식 {res['form_version']} · 배출계수 {res['ef_version']}"
                           f" · 기능단위 {res['fu']}", 7.6, "#777"), 0.22))
    rule()
    kv("계산 상태", res["state"], w="bold", vc=STATE_COLOR[res["state"]])
    if res["state_reason"]:
        text(res["state_reason"], size=7.8, c="#666", indent=0.1, n=70)
    if res["assumed"]:
        text("※ 가정값 입력(실측·보고값 아님): " + "; ".join(res["assumed"]), size=8.0, w="bold", c="#b71c1c", n=60)
    rule()
    if res["state"] == STATE_DONE:
        text("배출 내역 · 배치 전체 (kgCO2-eq/kg-H2)", size=9, w="bold")
        for k, v in res["items"].items():
            kv(k, f"{v:6.2f}   ({(v / res['total']) if res['total'] else 0:5.1%})")
        gap(0.03)
        kv("합계 · 부산물 공제 없음 (기본)", f"{res['total']:6.2f}", w="bold", size=10)
        gap(0.05)
        text("참고값", size=9, w="bold")
        kv("반응 구간만 · 부산물 공제 없음", f"{res['ref_total']:6.2f}", size=8.4)
        kv("배치 전체 · 질량 배분", "미산출" if res["alloc"] is None else f"{res['alloc']:6.2f}", size=8.4)
        text(res["alloc_note"], size=7.4, c="#777", indent=0.1, n=76)
        rule()
    text("검사 항목별 상태", size=9, w="bold")
    for name, (st, det) in res["checks"].items():
        ops.append(("status", (name, st, det), 0.205))
    if res["messages"]:
        rule()
        text("경고·안내", size=9, w="bold")
        for m in res["messages"]:
            text("· " + m, size=7.6, c="#8a3b00" if "확인" in m or "불가" in m else "#444", n=70)
    if res["history"]:
        rule()
        text("수정 이력", size=9, w="bold")
        for h in res["history"]:
            text(f"· {h.get('구간')}·{h.get('항목')}: {h.get('수정 전 값')} {h.get('수정 전 단위')} → "
                 f"{h.get('수정 후 값')} {h.get('수정 후 단위')} / {h.get('사유')} / {h.get('수정자')} / {h.get('일시')}",
                 size=7.6, c="#444", n=70)
    rule()
    text("적용 가정", size=8.4, w="bold")
    for a in res["assumptions"]:
        text("· " + a, size=7.4, c="#555", n=74)
    text("경계 밖 (계산 제외)", size=8.4, w="bold")
    text("· " + " / ".join(b for b, _ in res["boundary_out"]), size=7.4, c="#555", n=74)
    text("사용 배출계수 (시연용 예시값, 공식 계수 아님)", size=8.4, w="bold")
    text("· " + " / ".join(f"{e['id']} {e['value']:g} {e['unit']}" for e in res["ef_used"])
         + f"  [{res['ef_version']}]", size=7.4, c="#555", n=74)
    text("· 계수 출처·기준연도·적용 조건은 CSV에 기록", size=7.4, c="#555", n=74)

    top_band, bottom_band = 0.55, 0.15
    H = top_band + sum(o[2] for o in ops) + 0.35 + bottom_band
    fig = plt.figure(figsize=(W, H), dpi=200)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, W); ax.set_ylim(H, 0); ax.axis("off")
    fig.patch.set_facecolor("white")
    ax.text(W / 2, 0.16, res["label_data"], ha="center", va="top", fontsize=8.6, color="#b71c1c", fontweight="bold")
    ax.text(W / 2, 0.34, res["label_result"], ha="center", va="top", fontsize=8.0, color="#b71c1c")
    box_top = top_band
    ax.add_patch(FancyBboxPatch((0.2, box_top), W - 0.4, H - box_top - bottom_band, boxstyle="round,pad=0.02",
                                fc="#fffdf7", ec="#555", lw=1.1, ls=(0, (4, 3))))
    L, Rx = 0.45, W - 0.45
    y = box_top + 0.18
    for kind, p, h in ops:
        if kind == "title":
            ax.text(W / 2, y, p, ha="center", va="top", fontsize=15, fontweight="bold", color="#222")
        elif kind == "center":
            ax.text(W / 2, y, p[0], ha="center", va="top", fontsize=p[1], color=p[2])
        elif kind == "rule":
            ax.plot([L, Rx], [y + 0.05, y + 0.05], color="#999", lw=0.7, ls=(0, (2, 2)))
        elif kind == "text":
            ln, size, w, c, ind = p
            ax.text(L + ind, y, ln, ha="left", va="top", fontsize=size, fontweight=w, color=c)
        elif kind == "kv":
            k, v, w, size, c, vc = p
            ax.text(L, y, k, ha="left", va="top", fontsize=size, fontweight=w, color=c)
            ax.text(Rx, y, v, ha="right", va="top", fontsize=size, fontweight=w, color=vc)
        elif kind == "status":
            name, st, det = p
            ax.text(L, y, name, ha="left", va="top", fontsize=8.2, color="#222")
            ax.text(L + 1.75, y + 0.01, det[:44], ha="left", va="top", fontsize=6.9, color="#777")
            ax.text(Rx, y, f" {st} ", ha="right", va="top", fontsize=7.6, color="white", fontweight="bold",
                    bbox=dict(boxstyle="round,pad=0.22", fc=STATUS_COLOR[st], ec="none"))
        y += h
    fig.savefig(path, dpi=200, facecolor="white")
    plt.close(fig)


def _json_ready(res):
    out = dict(res)
    out["checks"] = {k: {"status": s, "detail": d} for k, (s, d) in res["checks"].items()}
    return out


def run(input_path, out_dir, sheet="실험기록", png=True):
    """입력 파일 1개를 처리하고 실험별 결과(dict) 목록을 돌려준다."""
    wb = load_workbook(input_path, data_only=True)
    S = Settings(wb)
    EF = load_ef(wb, str(S.val("EF_VERSION")).strip())
    recs = load_records(wb, sheet)
    if not recs:
        raise FormError(f"'{sheet}' 시트에 실험ID가 입력된 행이 없습니다.")
    hist = load_history(wb)
    acks = load_acks(wb)
    os.makedirs(out_dir, exist_ok=True)
    by_exp = OrderedDict()
    for r in recs:
        by_exp.setdefault(r["exp"], []).append(r)
    results = []
    for exp, rs in by_exp.items():
        res = evaluate(exp, rs, S, EF, hist, acks)
        base = os.path.join(out_dir, f"{exp}_탄소영수증")
        write_csv(res, base + ".csv")
        write_txt(res, base + ".txt")
        if png:
            write_png(res, base + ".png")
        with open(base + ".json", "w", encoding="utf-8") as f:
            json.dump(_json_ready(res), f, ensure_ascii=False, indent=1, default=str)
        results.append(res)
    return results


def main():
    ap = argparse.ArgumentParser(description="탄소영수증 시범 구현 (가상 데이터·예시 배출계수)")
    ap.add_argument("input", help="표준 입력양식 Excel 파일")
    ap.add_argument("-o", "--out", default="출력", help="출력 폴더 (기본: 출력)")
    ap.add_argument("--sheet", default="실험기록", help="실험 기록 시트 이름 (기본: 실험기록)")
    a = ap.parse_args()
    try:
        results = run(a.input, a.out, a.sheet)
    except FormError as e:
        print(f"[양식 오류] {e}", file=sys.stderr)
        sys.exit(2)
    for r in results:
        tot = f"{r['total']:.2f} kgCO2-eq/kg-H2" if r["total"] is not None else "-"
        print(f"{r['exp']}: {r['state']}  합계 {tot}  {r['state_reason']}")
    print(f"출력 폴더: {os.path.abspath(a.out)}")


if __name__ == "__main__":
    main()
