#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fin_thickness_calc.py 의 CLI 계산기를 라이브 엑셀 수식 워크북으로 변환.

매크로 없음(.xlsx). 모든 결과는 입력 셀을 참조하는 네이티브 수식으로 계산되어,
차압/온도/피치를 바꾸면 자동 재계산된다. 좌굴 최소두께(반복해)는 후보 두께 스캔
테이블 + MINIFS 로 대신한다.

셀 맵 (계산기 시트, 값은 C열):
  입력   C5 T · C6 ΔP_t · C7 ΔP_c · C8 p · C9 p_adj · C10 h · C11 K ·
         C12 f(스테이) · C13 공차 · C14 w · C15 P_min · C16 t_sel
  상수   C18 ν · C19 DF
  물성   C21 S(T) · C22 Sy(T) · C23 E(T) · C24 온도유효성
  결과   C26 인장 · C27 좌굴 · C28 요구(유효) · C29 요구(공칭) ·
         C30 두께판정 · C31 피치판정 · C32 높이판정
  파팅시트 C34 p̄ · C35 t_fin_eff · C36 스팬 · C37 M_d · C38 t_ps · C39 판정
  검증   C41 t_eff · C42 fσ_t · C43 인장판정 · C44 σ_c · C45 σ_cr ·
         C46 허용 · C47 압축판정 · C48 종합   (G45 = σ_cr,E 헬퍼)
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

OUT = "/home/user/PFHE/fin_strength_calc/PFHE_fin_calc.xlsx"
CG = "계산기"

S_TABLE = [(40, 22.8), (65, 22.8), (93, 22.6), (121, 21.4),
           (149, 18.6), (177, 14.5), (204, 11.0)]
SY_TABLE = [(40, 34.5), (93, 33.1), (149, 29.6), (204, 24.1)]
E_TABLE = [(25, 68900), (93, 66200), (149, 63400), (204, 60000)]
SCAN_LO, SCAN_HI, SCAN_STEP = 0.010, 1.000, 0.005

ARIAL = "Arial"
f_title = Font(name=ARIAL, size=14, bold=True, color="FFFFFF")
f_sub = Font(name=ARIAL, size=9, italic=True, color="F2F2F2")
f_hdr = Font(name=ARIAL, size=11, bold=True, color="FFFFFF")
f_lbl = Font(name=ARIAL, size=10)
f_lbl_b = Font(name=ARIAL, size=10, bold=True)
f_in = Font(name=ARIAL, size=10, bold=True, color="0000FF")
f_out = Font(name=ARIAL, size=10)
f_out_b = Font(name=ARIAL, size=10, bold=True)
f_note = Font(name=ARIAL, size=8, italic=True, color="555555")
f_warn = Font(name=ARIAL, size=10, bold=True, color="C00000")
fill_title = PatternFill("solid", fgColor="1F4E78")
fill_hdr = PatternFill("solid", fgColor="2E75B6")
fill_in = PatternFill("solid", fgColor="FFF2CC")
fill_res = PatternFill("solid", fgColor="E2EFDA")
fill_prop = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
ctr = Alignment(horizontal="center", vertical="center")
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
right = Alignment(horizontal="right", vertical="center")

wb = openpyxl.Workbook()

# ---------------- 물성DB ----------------
ws_m = wb.active
ws_m.title = "물성DB"
ws_m.sheet_view.showGridLines = False
ws_m["A1"] = "재료 물성 데이터 — AL 3003-O (ASME Sec. II-D 대표값)"
ws_m["A1"].font = Font(name=ARIAL, size=12, bold=True)
ws_m["A2"] = "⚠️ 대표값 — 실제 설계 전 적용 코드 연도판 값으로 교체할 것. 온도 오름차순 유지."
ws_m["A2"].font = f_warn

def write_table(ws, sr, sc, title, pairs, unit):
    c0, c1 = get_column_letter(sc), get_column_letter(sc + 1)
    ws.cell(sr, sc, title).font = f_lbl_b
    for k, txt in enumerate(("온도(°C)", unit)):
        c = ws.cell(sr + 1, sc + k, txt)
        c.font = f_hdr; c.fill = fill_hdr; c.alignment = ctr
    for i, (t, v) in enumerate(pairs):
        r = sr + 2 + i
        for k, val in enumerate((t, v)):
            c = ws.cell(r, sc + k, val)
            c.font = f_out; c.alignment = ctr; c.border = box
    first, lastr = sr + 2, sr + 1 + len(pairs)
    return f"물성DB!${c0}${first}:${c0}${lastr}", f"물성DB!${c1}${first}:${c1}${lastr}"

T_S, V_S = write_table(ws_m, 4, 1, "허용응력 S(T)", S_TABLE, "S (MPa)")
T_SY, V_SY = write_table(ws_m, 4, 4, "항복강도 Sy(T)", SY_TABLE, "Sy (MPa)")
T_E, V_E = write_table(ws_m, 4, 7, "탄성계수 E(T)", E_TABLE, "E (MPa)")
for col in "ABDEGH":
    ws_m.column_dimensions[col].width = 11
for name, ref in [("T_S", T_S), ("V_S", V_S), ("T_SY", T_SY), ("V_SY", V_SY),
                  ("T_E", T_E), ("V_E", V_E)]:
    wb.defined_names.add(DefinedName(name, attr_text=ref))

def interp(tc, TT, VV):
    m = f"MATCH({tc},{TT},1)"
    return (f"IF({tc}<=MIN({TT}),INDEX({VV},1),"
            f"IF({tc}>=MAX({TT}),INDEX({VV},COUNT({TT})),"
            f"INDEX({VV},{m})+({tc}-INDEX({TT},{m}))"
            f"/(INDEX({TT},{m}+1)-INDEX({TT},{m}))"
            f"*(INDEX({VV},{m}+1)-INDEX({VV},{m}))))")

# ---------------- 좌굴스캔 ----------------
ws_s = wb.create_sheet("좌굴스캔")
ws_s.sheet_view.showGridLines = False
for j, h in enumerate(["t 후보(mm)", "σ_c(MPa)", "σ_cr,E(MPa)",
                       "σ_cr Johnson(MPa)", "허용(MPa)", "OK(1/0)"], 1):
    c = ws_s.cell(1, j, h); c.font = f_hdr; c.fill = fill_hdr; c.alignment = ctr
ws_s["H1"] = "핀 좌굴 최소두께 = ΔP_c>0 일 때 OK=1 인 최소 t 후보 (계산기!C27, MINIFS)"
ws_s["H1"].font = f_note
n = int(round((SCAN_HI - SCAN_LO) / SCAN_STEP)) + 1
for i in range(n):
    r = i + 2
    ws_s.cell(r, 1, round(SCAN_LO + i * SCAN_STEP, 4)).number_format = "0.000"
    ws_s.cell(r, 2, f"={CG}!$C$7*{CG}!$C$8/A{r}")
    ws_s.cell(r, 3, f"=PI()^2*{CG}!$C$23/(12*(1-{CG}!$C$18^2))"
                    f"*(A{r}/({CG}!$C$11*{CG}!$C$10))^2")
    ws_s.cell(r, 4, f"=IF(C{r}<={CG}!$C$22/2,C{r},{CG}!$C$22*(1-{CG}!$C$22/(4*C{r})))")
    ws_s.cell(r, 5, f"=MIN(D{r}/{CG}!$C$19,{CG}!$C$21)")
    ws_s.cell(r, 6, f"=IF(B{r}<=E{r},1,0)")
    for j in range(2, 6):
        ws_s.cell(r, j).number_format = "0.00"
lastr = n + 1
wb.defined_names.add(DefinedName("BucklingT", attr_text=f"좌굴스캔!$A$2:$A${lastr}"))
wb.defined_names.add(DefinedName("BucklingOK", attr_text=f"좌굴스캔!$F$2:$F${lastr}"))
for col, w in zip("ABCDEF", (11, 11, 13, 17, 11, 9)):
    ws_s.column_dimensions[col].width = w

# ---------------- 계산기 ----------------
ws = wb.create_sheet("계산기")
wb.move_sheet("계산기", -(len(wb.sheetnames) - 1))
ws.sheet_view.showGridLines = False
for col, w in zip("ABCDE", (2, 31, 15, 8, 48)):
    ws.column_dimensions[col].width = w

def put(row, label, value=None, unit=None, note=None, *, kind="", num="0.000"):
    b = ws.cell(row, 2, label); b.font = f_lbl; b.alignment = left
    c = ws.cell(row, 3)
    if value is not None:
        c.value = value
    c.number_format = num; c.alignment = right; c.border = box
    styles = {"in": (f_in, fill_in), "res": (f_out_b, fill_res),
              "prop": (f_out, fill_prop)}
    if kind in styles:
        c.font, c.fill = styles[kind]
    else:
        c.font = f_out
    if unit:
        u = ws.cell(row, 4, unit); u.font = f_note; u.alignment = ctr
    if note:
        nn = ws.cell(row, 5, note); nn.font = f_note; nn.alignment = left
    return c

def section(row, text):
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    c = ws.cell(row, 2, text); c.font = f_hdr; c.fill = fill_hdr
    c.alignment = Alignment(horizontal="left", vertical="center")

ws.merge_cells("B1:E1")
ws["B1"] = "PFHE Plain 핀 적정두께 계산기 (ALPEMA 2024 기반 유도)"
ws["B1"].font = f_title; ws["B1"].fill = fill_title
ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 24
ws.merge_cells("B2:E2")
ws["B2"] = ("예비설계용 · 단위 MPa/mm · 파랑=입력 셀 · 인증설계는 "
            "ALPEMA 5.15.1.1(승인 계산법)/5.15.1.2(파열시험)에 따름")
ws["B2"].font = Font(name=ARIAL, size=9, italic=True, color="333333")
ws["B2"].alignment = Alignment(horizontal="center")
ws.merge_cells("B3:E3")
ws["B3"] = ('=IF($C$5>204,"⚠️ 설계온도 204°C 초과 — 3003 재질한계(ALPEMA Table 6-1) '
            '벗어남. 결과 무효!","")')
ws["B3"].font = f_warn; ws["B3"].alignment = Alignment(horizontal="center")

section(4, "① 입력 (파랑 셀을 편집)")
put(5, "설계온도 T", 65, "°C", "≤ 204 (3003 재질한계)", kind="in", num="0.0")
put(6, "핀 인장 유발 차압 ΔP_t", 5.0, "MPa", "통상 = 층 설계압력(MAWP)", kind="in")
put(7, "핀 압축 유발 차압 ΔP_c", 0.6, "MPa", "인접층 고압/자층 감압 시 (없으면 0)", kind="in")
put(8, "자층 핀 피치 p", 1.4, "mm", "ALPEMA 1.0–4.5", kind="in")
put(9, "인접층 핀 피치 p_adj", 0.0, "mm", "0=자층과 동일. 파팅시트 스팬용", kind="in")
put(10, "핀 높이 h", 6.5, "mm", "ALPEMA 2.0–12", kind="in", num="0.0")
put(11, "좌굴 유효길이계수 K", 0.7, "-", "기본 0.7, 구속 불확실 시 1.0", kind="in", num="0.00")
put(12, "인장 스테이 계수 f", 1.0, "-", "UG-50(b) 준용 시 1.10", kind="in", num="0.00")
put(13, "제작 공차계수", 1.0, "-", "0<f≤1 (예 0.9=공칭의 90%만 유효)", kind="in", num="0.00")
put(14, "파팅시트 차압 w=|P_A−P_B|", 5.0, "MPa", "층간 순차압", kind="in")
put(15, "균형가압 P_min", 5.0, "MPa", "양면 동시가압 낮은쪽(핀 어긋남 항)", kind="in")
put(16, "선정 핀 두께(공칭) t_sel", 0.30, "mm", "검증용. 0이면 검증 생략", kind="in")

section(17, "② 상수 및 온도별 물성 (자동)")
put(18, "포아송비 ν", 0.33, "-", "알루미늄", kind="prop", num="0.00")
put(19, "좌굴 설계계수 DF", 3.0, "-", "ASME UG-28 준용", kind="prop", num="0.0")
put(21, "허용응력 S(T)", f"={interp('$C$5', 'T_S', 'V_S')}", "MPa",
    "3003-O, 온도 선형보간", kind="prop", num="0.00")
put(22, "항복강도 Sy(T)", f"={interp('$C$5', 'T_SY', 'V_SY')}", "MPa",
    "Johnson 좌굴 천이용", kind="prop", num="0.00")
put(23, "탄성계수 E(T)", f"={interp('$C$5', 'T_E', 'V_E')}", "MPa",
    "좌굴 임계응력용", kind="prop", num="0")
put(24, "온도 유효성", '=IF($C$5>204,"온도초과(>204°C)","OK")', "-",
    "204°C 초과 시 결과 무효", kind="prop", num="General")

section(25, "③ 결과 — 핀 최소두께")
put(26, "[1] 핀 인장 최소두께 (식2)", "=IF($C$6<=0,0,$C$12*$C$6*$C$8/($C$21+$C$12*$C$6))",
    "mm", "t ≥ f·ΔP_t·p/(S+f·ΔP_t)", kind="res")
put(27, "[2] 핀 좌굴 최소두께 (식3~5)",
    "=IF($C$7<=0,0,IF(COUNTIF(BucklingOK,1)=0,999,_xlfn.MINIFS(BucklingT,BucklingOK,1)))",
    "mm", "Johnson 천이·전피치 하중폭 (좌굴스캔 시트)", kind="res")
ws.cell(27, 5).value = ('=IF(AND($C$7>0,COUNTIF(BucklingOK,1)=0),'
                        '"⚠️ 좌굴 소요두께 1.0mm 초과 — 표준 핀 불가",'
                        '"Johnson 천이·전피치 하중폭 (좌굴스캔 시트)")')
put(28, "[3] 요구 핀두께 (유효)", "=MAX($C$26,$C$27)", "mm",
    "max(인장, 좌굴)", kind="res")
put(29, "    요구 핀두께 (공칭)", '=IF($C$5>204,"온도초과",$C$28/$C$13)', "mm",
    "유효 / 공차계수", kind="res")
put(30, "  ALPEMA 두께범위 판정",
    '=IF($C$5>204,"—",IF(AND($C$29>=0.15,$C$29<=0.7),"OK 범위내","⚠️ 범위밖(0.15-0.7)"))',
    "-", "제작 가능 범위 0.15–0.7 mm", kind="prop", num="General")
put(31, "  핀 피치 판정",
    '=IF(AND($C$8>=1,$C$8<=4.5),"OK 범위내","⚠️ 범위밖(1.0-4.5)")', "-",
    "ALPEMA 5.11.4", kind="prop", num="General")
put(32, "  핀 높이 판정",
    '=IF(AND($C$10>=2,$C$10<=12),"OK 범위내","⚠️ 범위밖(2-12)")', "-",
    "ALPEMA 5.11.4", kind="prop", num="General")

section(33, "④ 파팅시트 두께 (식6~7)")
put(34, "지배 피치 p̄", "=IF($C$9>0,MAX($C$8,$C$9),$C$8)", "mm",
    "max(자층, 인접층 피치)", kind="prop")
put(35, "스팬 공제 핀두께 t_fin,eff",
    "=IF($C$9>$C$8,0,IF($C$16>0,$C$16*$C$13,$C$28))", "mm",
    "인접층 지배 시 0(보수) — 문서 §6.1", kind="prop")
put(36, "스팬 s", "=$C$34-$C$35", "mm", "s = p̄ − t_fin,eff", kind="prop")
put(37, "설계 모멘트 M_d", "=$C$14*$C$36^2/12+$C$15*($C$34-$C$35)*$C$34/8",
    "N·mm/mm", "w·s²/12 + P_min·(p̄−t)·p̄/8", kind="prop", num="0.0000")
put(38, "파팅시트 최소두께 t_ps",
    '=IF($C$36<=0,"스팬오류",SQRT(4*$C$37/$C$21))', "mm",
    "σ_b ≤ 1.5·S", kind="res")
put(39, "  ALPEMA 시트범위 판정",
    '=IF(NOT(ISNUMBER($C$38)),"—",IF($C$38>2,"⚠️ 범위밖(>2.0)",'
    'IF($C$38<0.8,"소요<0.8 → 표준 0.8 적용","OK 범위내")))', "-",
    "ALPEMA 5.11.1: 0.8–2.0 mm", kind="prop", num="General")

section(40, "⑤ 선정 두께 검증 (t_sel 입력 시)")
put(41, "유효 핀두께 t_eff", '=IF($C$16>0,$C$16*$C$13,"")', "mm",
    "공칭 × 공차계수", kind="prop")
put(42, "인장 f·σ_t",
    '=IF($C$16>0,IF($C$16*$C$13>=$C$8,"두께≥피치",'
    'IF($C$6>0,$C$12*$C$6*($C$8-$C$41)/$C$41,0)),"")', "MPa",
    "vs S(T)", kind="prop", num="0.00")
put(43, "  인장 판정",
    '=IF($C$16>0,IF(NOT(ISNUMBER($C$42)),"—",'
    'IF($C$42<=$C$21,"합격 ✓","불합격 ✗")),"")', "-", "", kind="res",
    num="General")
put(44, "압축 σ_c",
    '=IF(AND($C$16>0,$C$7>0),IF($C$16*$C$13>=$C$8,"두께≥피치",'
    '$C$7*$C$8/$C$41),"")', "MPa", "ΔP_c·p/t_eff", kind="prop", num="0.00")
# G45 = σ_cr,E 헬퍼 (탄성 임계응력)
ws.cell(45, 7, '=IF(AND($C$16>0,$C$7>0,$C$16*$C$13<$C$8),'
               'PI()^2*$C$23/(12*(1-$C$18^2))*($C$41/($C$11*$C$10))^2,"")')
ws.cell(45, 7).font = f_note
ws.cell(45, 6, "σ_cr,E 헬퍼 →").font = f_note
put(45, "압축 σ_cr (Johnson)",
    '=IF(ISNUMBER($G$45),IF($G$45<=$C$22/2,$G$45,'
    '$C$22*(1-$C$22/(4*$G$45))),"")', "MPa", "비탄성 천이", kind="prop",
    num="0.00")
put(46, "압축 허용 min(σ_cr/DF,S)",
    '=IF(ISNUMBER($C$45),MIN($C$45/$C$19,$C$21),"")', "MPa", "", kind="prop",
    num="0.00")
put(47, "  압축 판정",
    '=IF(AND($C$16>0,$C$7>0),IF(NOT(ISNUMBER($C$44)),"—",'
    'IF($C$44<=$C$46,"합격 ✓","불합격 ✗")),"")', "-", "", kind="res",
    num="General")
put(48, "■ 종합 판정",
    '=IF($C$5>204,"온도초과",IF($C$16>0,IF(OR($C$43="불합격 ✗",$C$47="불합격 ✗"),'
    '"■ 종합 불합격","■ 종합 합격"),"(선정두께 미입력 — 검증 생략)"))', "-",
    "", kind="res", num="General")
ws.cell(48, 3).font = Font(name=ARIAL, size=11, bold=True, color="1F4E78")

for r in list(range(5, 17)) + list(range(21, 33)) + list(range(34, 40)) \
        + list(range(41, 49)):
    ws.row_dimensions[r].height = 15

# ---------------- 설명 시트 ----------------
ws_n = wb.create_sheet("설명")
ws_n.sheet_view.showGridLines = False
ws_n.column_dimensions["A"].width = 100
notes = [
    ("PFHE Plain 핀 적정두께 계산기 — 사용 안내", 12, True, "1F4E78"),
    ("", 10, False, None),
    ("■ 사용법", 11, True, "2E75B6"),
    ("1. [계산기] 시트의 파랑(노랑배경) 셀만 편집합니다. 나머지는 자동 계산 수식입니다.", 10, False, None),
    ("2. 단위는 압력 MPa, 길이 mm 입니다 (1 MPa = 10 bar ≈ 10.2 kgf/cm²).", 10, False, None),
    ("3. [1][2]는 최소 소요두께, [3]은 채택 기준(둘 중 큰 값)입니다.", 10, False, None),
    ("4. t_sel(선정 두께)을 입력하면 ⑤ 검증에서 합/부를 판정합니다. 0이면 검증을 건너뜁니다.", 10, False, None),
    ("5. ΔP_c(압축)가 0이면 좌굴 검토는 생략됩니다.", 10, False, None),
    ("", 10, False, None),
    ("■ 계산 근거 (수식)", 11, True, "2E75B6"),
    ("[1] 핀 인장:    t ≥ f·ΔP_t·p / (S(T) + f·ΔP_t)", 10, False, None),
    ("[2] 핀 좌굴:    ΔP_c·p/t ≤ min(σ_cr/DF, S), σ_cr는 오일러식 + Johnson 비탄성 천이", 10, False, None),
    ("               → [좌굴스캔] 시트가 후보 두께를 훑어 조건을 만족하는 최소 t를 MINIFS로 반환", 10, False, None),
    ("[3] 파팅시트:  t_ps ≥ √(4·M_d/S),  M_d = w·s²/12 + P_min·(p̄−t)·p̄/8", 10, False, None),
    ("온도 반영:     S(T)/Sy(T)/E(T)를 [물성DB] 표에서 온도 선형보간 (204°C 초과 무효)", 10, False, None),
    ("", 10, False, None),
    ("■ 시트 구성", 11, True, "2E75B6"),
    ("계산기   — 입력/결과 (메인)", 10, False, None),
    ("물성DB   — S(T)/Sy(T)/E(T) 표 (코드 연도판으로 교체 대상)", 10, False, None),
    ("좌굴스캔 — 좌굴 최소두께 반복해용 스캔 테이블 (수정 불필요)", 10, False, None),
    ("", 10, False, None),
    ("■ 한계 및 주의", 11, True, "C00000"),
    ("· 예비설계용. 정적 압력하중 전용(열응력·피로는 ALPEMA 8.1 운전제한으로 별도 관리).", 10, False, None),
    ("· S/Sy/E 표는 대표값 — 적용 코드 연도판으로 반드시 검증/교체할 것.", 10, False, None),
    ("· 149°C 이상 장기운전은 크리프 별도 검토. Serrated/천공핀·헤더·노즐은 범위 밖.", 10, False, None),
    ("· 인증 설계는 ALPEMA 5.15.1.1(승인 계산법)/5.15.1.2(파열시험)으로 확정.", 10, False, None),
    ("· 상세 유도는 fin_thickness_derivation.md, 원본 로직은 fin_thickness_calc.py 참조.", 10, False, None),
]
for i, (txt, sz, bold, color) in enumerate(notes, 1):
    c = ws_n.cell(i, 1, txt)
    c.font = Font(name=ARIAL, size=sz, bold=bold, color=color or "000000")
    c.alignment = left

# 열 때 전체 재계산 강제 → 어떤 뷰어에서도 값이 즉시 표시됨
wb.calculation.fullCalcOnLoad = True

wb.save(OUT)
print(f"saved: {OUT}")
print(f"scan rows: {n}")
