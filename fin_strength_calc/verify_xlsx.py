#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""엑셀 워크북을 formulas 라이브러리로 평가해 fin_thickness_calc.py 와 다중 케이스 대조.

LibreOffice 없이 Excel 수식을 파싱·평가한다. 입력 셀을 케이스별로 오버라이드하여
파이썬 CLI 로직과 결과가 일치하는지 검증한다.
"""
import warnings
import formulas
import fin_thickness_calc as fc

warnings.filterwarnings("ignore")
XLSX = "/home/user/PFHE/fin_strength_calc/PFHE_fin_calc.xlsx"

xl = formulas.ExcelModel().loads(XLSX).finish()

# 입력 셀 주소 (계산기 시트)
CELL = dict(T="C5", dpt="C6", dpc="C7", p="C8", padj="C9", h="C10", K="C11",
            f="C12", tol="C13", w="C14", pmin="C15", tsel="C16")
# 결과 셀
RES = dict(S="C21", Sy="C22", E="C23", tempok="C24", ten="C26", buc="C27",
           reqE="C28", reqN="C29", pbar="C34", tfe="C35", span="C36",
           Md="C37", tps="C38", teff="C41", sigt="C42", tv="C43", sigc="C44",
           scr="C45", allow="C46", cv="C47", overall="C48")

# 모델 키 접두 파악
sample = next(iter(xl.cells)) if hasattr(xl, "cells") else None
sol0 = xl.calculate()
prefix = None
for k in sol0:
    if "계산기" in k and k.upper().rstrip().endswith("!C21"):
        prefix = k[: k.rfind("!") + 1]
        break
assert prefix, "계산기 시트 키 접두를 찾지 못함"

def key(addr):
    return f"{prefix}{addr}"

def val(sol, addr):
    v = sol[key(addr)]
    v = v.value if hasattr(v, "value") else v
    try:
        return v[0][0]
    except (TypeError, IndexError):
        return v

def py_ref(T, dpt, dpc, p, padj, h, K, f, tol, w, pmin, tsel):
    """파이썬 CLI 로직으로 기대값 계산."""
    s = fc.allowable_stress(T); sy = fc.yield_strength(T); e = fc.youngs_modulus(T)
    ten = fc.fin_min_thickness_tension(dpt, p, s, f)
    buc = fc.fin_min_thickness_compression(dpc, p, h, e, sy, s, K)
    reqE = max(ten, buc); reqN = reqE / tol
    pbar = max(p, padj) if padj > 0 else p
    tfe = 0.0 if padj > p else (tsel * tol if tsel else reqE)
    span = pbar - tfe
    md = fc.sheet_design_moment(w, span, pbar, tfe, pmin)
    tps = fc.sheet_min_thickness(md, s)
    out = dict(S=s, Sy=sy, E=e, ten=ten, reqE=reqE, reqN=reqN, pbar=pbar,
               span=span, Md=md, tps=tps)
    if tsel:
        teff = tsel * tol
        out["teff"] = teff
        out["sigt"] = fc.fin_stress_tension(dpt * f, p, teff) if dpt > 0 else 0.0
        if dpc > 0:
            out["sigc"] = fc.fin_stress_compression(dpc, p, teff)
            out["scr"] = fc.buckling_critical(teff, h, e, sy, K)
            out["allow"] = fc.compression_allowable(teff, h, e, sy, s, K)
    return out

CASES = [
    ("기본",       dict(T=65, dpt=5.0, dpc=0.6, p=1.4, padj=0, h=6.5, K=0.7, f=1.0, tol=1.0, w=5.0, pmin=5.0, tsel=0.30)),
    ("온도보간107", dict(T=107, dpt=8.0, dpc=0.0, p=2.0, padj=0, h=6.5, K=0.7, f=1.0, tol=1.0, w=0.0, pmin=0.0, tsel=0.0)),
    ("고온204",     dict(T=204, dpt=5.0, dpc=0.0, p=1.4, padj=0, h=6.5, K=0.7, f=1.0, tol=1.0, w=0.0, pmin=0.0, tsel=0.0)),
    ("스테이1.1",   dict(T=65, dpt=5.0, dpc=0.0, p=1.4, padj=0, h=6.5, K=0.7, f=1.10, tol=1.0, w=0.0, pmin=0.0, tsel=0.0)),
    ("공차0.9검증", dict(T=65, dpt=5.0, dpc=0.6, p=1.4, padj=0, h=6.5, K=0.7, f=1.0, tol=0.9, w=0.0, pmin=0.0, tsel=0.26)),
    ("인접피치지배", dict(T=65, dpt=5.0, dpc=0.0, p=1.4, padj=4.0, h=6.5, K=0.7, f=1.0, tol=1.0, w=5.0, pmin=0.0, tsel=0.30)),
    ("압축지배",    dict(T=93, dpt=0.5, dpc=3.0, p=2.5, padj=0, h=11.0, K=1.0, f=1.0, tol=1.0, w=0.0, pmin=0.0, tsel=0.0)),
]

# span/tfe 는 좌굴지배(tsel=0) 시 reqE(스캔 0.005 해상도)가 전파되므로 동일 허용오차
TOL = {"E": 5, "scr": 0.15, "allow": 0.05, "sigt": 0.02, "sigc": 0.02,
       "buc": 0.006, "reqE": 0.006, "reqN": 0.007, "span": 0.006, "tfe": 0.006}
allpass = True
for name, inp in CASES:
    overrides = {key(CELL[k]): [[v]] for k, v in inp.items()}
    sol = xl.calculate(inputs=overrides)
    ref = py_ref(**inp)
    print(f"\n=== {name}: T={inp['T']} ΔPt={inp['dpt']} ΔPc={inp['dpc']} "
          f"p={inp['p']} tol={inp['tol']} tsel={inp['tsel']} ===")
    for tag, pv in ref.items():
        xv = val(sol, RES[tag])
        tol_ = TOL.get(tag, 0.001)
        try:
            ok = abs(float(xv) - float(pv)) <= tol_
        except (TypeError, ValueError):
            ok = False
        allpass &= ok
        mark = "✓" if ok else "✗ 불일치"
        xs = f"{float(xv):.4f}" if isinstance(xv, (int, float)) else str(xv)
        print(f"  {tag:<6} 엑셀={xs:>12}  파이썬={pv:>12.4f}  {mark}")
    # 온도초과 배너/판정
    if inp["T"] > 204:
        banner = val(sol, "C29")
        print(f"  C29(공칭) 온도초과표시: {banner!r}")

print("\n" + "=" * 60)
print("전체 결과:", "PASS ✓  (엑셀 수식 = 파이썬 CLI 로직 일치)"
      if allpass else "FAIL ✗  — 위 불일치 항목 확인")
