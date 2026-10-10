"""보고서 그림 생성(이미 출력된 영수증을 잘라 붙이는 작업만 함, 계산은 하지 않음)

그림 1: 단위 오기 사례의 '오류 발견 → 수정 → 재계산'과 결과 한 줄의 근거(전력) 확대
그림 3: 공개 논문 보고값만 입력한 경우의 영수증 윗부분
run_validation.py, make_paper_case.py 실행 뒤에 실행한다.
"""
import os

from PIL import Image, ImageDraw, ImageFont

R = "결과"
BOLD = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
REG = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"


def font(path, size):
    return ImageFont.truetype(path, size, index=1)  # KR


def rule_rows(img):
    """영수증의 점선 구분선 y 좌표"""
    px = img.load()
    w, h = img.size
    ys = []
    for y in range(h):
        n = 0
        for x in range(100, min(1180, w), 2):
            r, g, b = px[x, y]
            if abs(r - 153) < 12 and abs(g - 153) < 12 and abs(b - 153) < 12:
                n += 1
        if n > 150 and (not ys or y - ys[-1] > 2):
            ys.append(y)
    return ys


def box(d, xy, text_lines, fill, outline, f_title, f_body, title=None):
    d.rounded_rectangle(xy, radius=18, fill=fill, outline=outline, width=3)
    x, y = xy[0] + 28, xy[1] + 22
    if title:
        d.text((x, y), title, font=f_title, fill="#222")
        y += 52
    for ln in text_lines:
        d.text((x, y), ln, font=f_body, fill="#333")
        y += 40


def figure1():
    err = Image.open(os.path.join(R, "영수증", "사례2_단위오기", "EXP-800_탄소영수증.png")).convert("RGB")
    fix = Image.open(os.path.join(R, "영수증", "사례2_단위오기_수정후", "EXP-800_탄소영수증.png")).convert("RGB")
    er, fr = rule_rows(err), rule_rows(fix)
    err_c = err.crop((0, 0, err.width, er[3] + 4))                     # 상태·검사·경고까지
    top = fix.crop((0, 0, fix.width, fr[2] + 4))                         # 상태·배출 내역·참고값
    hist = fix.crop((0, fr[3], fix.width, fr[4] + 4))                    # 수정 이력
    fix_c = Image.new("RGB", (fix.width, top.height + hist.height), "white")
    fix_c.paste(top, (0, 0)); fix_c.paste(hist, (0, top.height))

    pad, gap = 40, 140
    W = pad * 2 + err_c.width + gap + fix_c.width
    step_h, trace_h = 250, 330
    H = 30 + step_h + 40 + max(err_c.height, fix_c.height) + 40 + trace_h + pad
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    fT, fB, fS = font(BOLD, 40), font(REG, 30), font(BOLD, 34)

    # 단계 상자 3개
    y0 = 30
    bw = (W - pad * 2 - 2 * 60) // 3
    steps = [
        ("① 입력 오류", ["반응 구간 전력 5,000 Wh를", "5,000 kWh로 잘못 입력"], "#fdecea", "#c62828"),
        ("② 시스템 대응", ["전력 원단위 20,006 kWh/kg-H2", "(범위 5~100 밖) → '확인 필요'", "계산 보류, 원본 로그 대조 안내"],
         "#fff4e5", "#e65100"),
        ("③ 수정 후", ["5,000 Wh로 수정, 수정 이력 1건 보존", "재검사 통과 → 재계산 14.14", "kgCO2-eq/kg-H2 (정상 사례와 같음)"],
         "#e8f5e9", "#2e7d32"),
    ]
    for i, (t, lines, fl, ol) in enumerate(steps):
        x = pad + i * (bw + 60)
        box(d, (x, y0, x + bw, y0 + step_h), lines, fl, ol, fT, fB, title=t)
        if i < 2:
            ax = x + bw + 8
            d.polygon([(ax, y0 + step_h // 2 - 22), (ax + 44, y0 + step_h // 2), (ax, y0 + step_h // 2 + 22)],
                      fill="#555")

    # 영수증 두 장
    y1 = y0 + step_h + 40
    img.paste(err_c, (pad, y1))
    x2 = pad + err_c.width + gap
    img.paste(fix_c, (x2, y1))
    mid = y1 + 380
    d.line([(pad + err_c.width + 15, mid), (x2 - 25, mid)], fill="#555", width=6)
    d.polygon([(x2 - 25, mid - 18), (x2 + 5, mid), (x2 - 25, mid + 18)], fill="#555")
    d.text((pad + err_c.width + 6, mid + 26), "수정·", font=fS, fill="#555")
    d.text((pad + err_c.width + 6, mid + 66), "재검사", font=fS, fill="#555")

    # 근거 추적 확대
    y2 = y1 + max(err_c.height, fix_c.height) + 40
    d.rounded_rectangle((pad, y2, W - pad, y2 + trace_h), radius=18, fill="#f5f7fb", outline="#1f3864", width=3)
    d.text((pad + 28, y2 + 20), "결과 한 줄의 근거 — 전력(계통) 11.70 kgCO2-eq/kg-H2 (CSV 근거 추적 출력에서 확대)",
           font=fT, fill="#1f3864")
    rows = [
        ("입력 행", "실험기록 3·12·22행: 승온 1,200 Wh + 반응 5,000 Wh + 중단 300 Wh"),
        ("환산값", "6.5 kWh  ·  수소 생산량 실험기록 14행: 250 g → 0.25 kg"),
        ("계수", "EF-ELEC-01  0.45 kgCO2-eq/kWh (시연용 예시값, 버전 EF-DEMO-v0.1)"),
        ("식 → 결과", "6.5 ÷ 0.25 × 0.45 = 11.70"),
    ]
    yy = y2 + 90
    for k, v in rows:
        d.text((pad + 40, yy), k, font=fS, fill="#333")
        d.text((pad + 320, yy), v, font=font(REG, 34), fill="#222")
        yy += 56
    out = os.path.join(R, "보고서그림1_오류수정재계산.png")
    img.save(out, dpi=(200, 200))
    return out, img.size


def figure3():
    c = Image.open(os.path.join(R, "논문적용", "영수증_1차_보고값만", "LIT-AF2016-Fe30_탄소영수증.png")).convert("RGB")
    rr = rule_rows(c)
    cut = rr[2] + 167  # 경고·안내 첫 두 항목까지
    crop = c.crop((0, 0, c.width, cut))
    f = font(REG, 26)
    out = Image.new("RGB", (c.width, cut + 110), "white")
    out.paste(crop, (0, 0))
    d = ImageDraw.Draw(out)
    d.text((90, cut + 8), "· … 나머지 7개 누락 항목(반응·중단 전력, 외부 공급열 2구간, 메탄 신규 공급량,", font=f, fill="#555")
    d.text((90, cut + 44), "  수소 생산량, 촉매 총 사용 횟수) 안내는 생략. 전체 영수증은 첨부 5", font=f, fill="#555")
    p = os.path.join(R, "보고서그림3_논문보고값_영수증.png")
    out.save(p, dpi=(200, 200))
    return p, out.size


if __name__ == "__main__":
    print(figure1())
    print(figure3())
