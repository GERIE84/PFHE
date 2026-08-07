#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""강도계산서 시트 수식을 평가해 오류 유무와 계산기 연동을 확인."""
import warnings
import formulas

warnings.filterwarnings("ignore")
xl = formulas.ExcelModel().loads("PFHE_fin_calc.xlsx").finish()
sol = xl.calculate()

ERR_TOKENS = ("#NAME", "#VALUE", "#REF", "#DIV", "#N/A", "#NUM", "#NULL")
rep, calc, errs = {}, {}, []
for k, v in sol.items():
    val = v.value if hasattr(v, "value") else v
    try:
        val = val[0][0]
    except (TypeError, IndexError):
        pass
    addr = k.split("!")[-1].replace("$", "")
    target = rep if "강도계산서" in k else (calc if "계산기" in k else None)
    if target is None:
        continue
    target[addr] = val
    if any(tok in str(val).upper() for tok in ERR_TOKENS):
        errs.append((("RPT" if target is rep else "CAL"), addr, str(val)[:40]))

print(f"강도계산서 평가 셀: {len(rep)}   계산기 평가 셀: {len(calc)}")
print("오류 셀:", errs if errs else "없음 ✓")

print("\n--- §2 설계조건 (E열=값) ---")
for r in range(12, 24):
    print(f"  E{r} = {rep.get(f'E{r}')!r:<24} G{r} = {str(rep.get(f'G{r}'))[:52]!r}")

print("\n--- §3 재료 ---")
for r in range(27, 33):
    print(f"  E{r} = {rep.get(f'E{r}')!r}")

print("\n--- §4 계산결과 ---")
for r in range(36, 42):
    print(f"  E{r} = {rep.get(f'E{r}')!r:<22} G{r} = {str(rep.get(f'G{r}'))[:50]!r}")

print("\n--- §5 검증 (E=응력 F=허용 G=SF H=판정) ---")
for r in range(45, 49):
    print(f"  {r}: E={rep.get(f'E{r}')!r:<10} F={rep.get(f'F{r}')!r:<10} "
          f"G={rep.get(f'G{r}')!r:<10} H={rep.get(f'H{r}')!r}")

print("\n--- §6 결론 ---")
for r in range(51, 57):
    v = rep.get(f"B{r}")
    if v not in (None, ""):
        print(f"  B{r}: {v}")

print("\n--- 계산기 연동 대조 ---")
pairs = [("E12", "C5"), ("E13", "C6"), ("E27", None), ("E28", "C21"),
         ("E36", "C26"), ("E37", "C27"), ("E38", "C28"), ("E39", "C29"),
         ("E40", "C38"), ("E45", "C42"), ("F45", "C21"), ("E46", "C44"),
         ("F46", "C46")]
ok = True
for ra, ca in pairs:
    if ca is None:
        continue
    rv, cv = rep.get(ra), calc.get(ca)
    try:
        same = abs(float(rv) - float(cv)) < 1e-9
    except (TypeError, ValueError):
        same = str(rv) == str(cv)
    ok &= same
    print(f"  강도계산서!{ra}={rv!r:<22} 계산기!{ca}={cv!r:<22} "
          f"{'✓' if same else '✗'}")

# 안전율 수동 검산
sf_t = calc.get("C21") / calc.get("C42") if calc.get("C42") else None
sf_c = calc.get("C46") / calc.get("C44") if calc.get("C44") else None
for tag, addr, expect in [("인장 SF", "G45", sf_t), ("압축 SF", "G46", sf_c)]:
    got = rep.get(addr)
    same = expect is not None and abs(float(got) - expect) < 1e-9
    ok &= same
    print(f"  {tag}: 계산서={got!r} 검산={expect!r} {'✓' if same else '✗'}")

print("\n" + "=" * 58)
print("결과:", "PASS ✓ (오류 없음 · 계산기 연동 일치)"
      if ok and not errs else "FAIL ✗")
