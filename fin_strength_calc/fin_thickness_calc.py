#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PFHE Plain 핀 적정두께 계산기 (Rev.1).

차압(ΔP)과 설계온도(T)로부터 Plain 핀의 최소 요구두께와 파팅시트 최소두께를
계산한다. 유도 근거는 같은 폴더의 fin_thickness_derivation.md 참조.

  - 핀 인장:        t >= f·ΔP_t·p / (S(T) + f·ΔP_t)   (f: 스테이 계수)   ... 식 (2)/(2'')
  - 핀 압축·좌굴:   ΔP_c·p/t <= min(σ_cr/3, S(T)),
                    σ_cr: 탄성 σ_cr,E = π²E/(12(1-ν²))·(t/(K·h))² 에
                    Johnson 포물선 천이(σ_cr,E > Sy/2 구간) 적용        ... 식 (3)~(5)
  - 파팅시트 굽힘:  M_d = w·s²/12 + P_min·(p̄-t_f)·p̄/8,
                    t_ps >= sqrt(4·M_d/S(T))                            ... 식 (6)~(7)

허용응력 S(T)/항복 Sy(T)/탄성계수 E(T): SB-209 3003-O, ASME Sec.VIII Div.1
(Sec.II-D) 대표값. ⚠️ 실제 설계 전 적용 코드 연도판 값으로 교체할 것
(--s-table 옵션 또는 테이블 상수 수정). 본 계산은 예비설계용이며, 인증 설계는
ALPEMA 5.15.1.1(승인 계산법) 또는 5.15.1.2(파열시험)에 따른다.

사용 예:
    python3 fin_thickness_calc.py --dp-t 5.0 --temp 65 --pitch 1.4 --height 6.5
    python3 fin_thickness_calc.py --dp-t 5.0 --dp-c 0.6 --temp 93 \\
        --pitch 2.0 --height 9.5 --t-fin 0.4 --w-sheet 5.0 --p-min 5.0
    python3 fin_thickness_calc.py --selftest
"""

import argparse
import json
import math
import sys

NU = 0.33          # 알루미늄 포아송비
DF_BUCKLING = 3.0  # 좌굴 설계계수 (ASME Div.1 UG-28 철학 준용)
K_DEFAULT = 0.7    # 유효길이계수: 고정-힌지 이론값 준용, sway 억제 전제 (문서 §5.2)

# ASME Sec.II-D Table 1B — SB-209 3003-O 허용응력 대표값 [(온도 °C, S MPa)]
# 상온 22.8 MPa = US 단위표 3.3 ksi 반올림 (이론치 ⅔Sy=23.0 대비 보수적)
# ⚠️ 대표값: 적용 코드 연도판으로 검증/교체 필요
ALLOWABLE_S_TABLE = [
    (40.0, 22.8),
    (65.0, 22.8),
    (93.0, 22.6),
    (121.0, 21.4),
    (149.0, 18.6),
    (177.0, 14.5),
    (204.0, 11.0),
]

# ASME Sec.II-D Table Y-1 계열 — 3003-O 최소 항복강도 대표값 [(°C, MPa)]
YIELD_SY_TABLE = [
    (40.0, 34.5),
    (93.0, 33.1),
    (149.0, 29.6),
    (204.0, 24.1),
]

# ASME Sec.II-D TM — 알루미늄 탄성계수 [(온도 °C, E MPa)]
E_TABLE = [
    (25.0, 68900.0),
    (93.0, 66200.0),
    (149.0, 63400.0),
    (204.0, 60000.0),
]

# 재질 온도 한계: ALPEMA Table 6-1, SB-209 3003 (ASME) 최대 적용 설계온도 [°C]
# 사용자 s-table 교체와 무관하게 main/report 에서 강제된다.
T_MAX = 204.0

# ALPEMA 5.11 표준 제작 범위 [mm]
FIN_T_RANGE = (0.15, 0.7)
FIN_PITCH_RANGE = (1.0, 4.5)
FIN_H_RANGE = (2.0, 12.0)
SHEET_T_RANGE = (0.8, 2.0)


def _validate_table(table, name):
    if not table:
        raise ValueError(f"{name}: 테이블이 비어 있음")
    temps = [row[0] for row in table]
    if temps != sorted(temps) or len(set(temps)) != len(temps):
        raise ValueError(f"{name}: 온도가 오름차순·중복없음 이어야 함: {temps}")
    if any(row[1] <= 0 for row in table):
        raise ValueError(f"{name}: 값은 양수여야 함")


def _interp(table, x, name):
    """테이블 선형 보간. 하한 미만은 첫 값 사용, 상한 초과는 오류."""
    if not (isinstance(x, (int, float)) and math.isfinite(x)):
        raise ValueError(f"{name}: 온도 입력이 유한한 수가 아님: {x}")
    if x > table[-1][0]:
        raise ValueError(
            f"{name}: {x} °C 는 테이블 상한 {table[-1][0]} °C 초과")
    if x <= table[0][0]:
        return table[0][1]
    for (x0, y0), (x1, y1) in zip(table, table[1:]):
        if x0 <= x <= x1:
            return y0 + (y1 - y0) * (x - x0) / (x1 - x0)
    raise AssertionError("unreachable")


def allowable_stress(temp_c, table=None):
    """설계온도[°C] → 허용응력 S(T) [MPa]."""
    return _interp(table or ALLOWABLE_S_TABLE, temp_c, "S(T)")


def yield_strength(temp_c):
    """설계온도[°C] → 최소 항복강도 Sy(T) [MPa] (Johnson 천이용)."""
    return _interp(YIELD_SY_TABLE, temp_c, "Sy(T)")


def youngs_modulus(temp_c):
    """설계온도[°C] → 탄성계수 E(T) [MPa]."""
    return _interp(E_TABLE, temp_c, "E(T)")


def _check_positive(**kwargs):
    for name, v in kwargs.items():
        if not (isinstance(v, (int, float)) and math.isfinite(v) and v > 0):
            raise ValueError(f"{name} 는 양의 유한값이어야 함: {v}")


def _check_nonneg(**kwargs):
    for name, v in kwargs.items():
        if not (isinstance(v, (int, float)) and math.isfinite(v) and v >= 0):
            raise ValueError(f"{name} 는 0 이상 유한값이어야 함: {v}")


def fin_stress_tension(dp_t, pitch, t):
    """식 (1): 핀 인장응력 σ_t = ΔP_t(p-t)/t [MPa]."""
    _check_nonneg(dp_t=dp_t)
    _check_positive(pitch=pitch)
    if not 0 < t < pitch:
        raise ValueError(f"핀 두께 t={t} 는 0 < t < 피치({pitch}) 이어야 함")
    return dp_t * (pitch - t) / t


def fin_min_thickness_tension(dp_t, pitch, s_allow, stay_factor=1.0):
    """식 (2)/(2''): 인장 기준 최소 핀 두께 [mm]. stay_factor=1.10 → UG-50(b) 준용."""
    _check_nonneg(dp_t=dp_t)
    _check_positive(pitch=pitch, s_allow=s_allow, stay_factor=stay_factor)
    if dp_t == 0:
        return 0.0
    f = stay_factor
    return f * dp_t * pitch / (s_allow + f * dp_t)


def fin_stress_compression(dp_c, pitch, t):
    """식 (3): 핀 압축응력 σ_c = ΔP_c·p/t [MPa] (전피치 하중폭, 문서 §5.1)."""
    _check_nonneg(dp_c=dp_c)
    _check_positive(pitch=pitch, t=t)
    return dp_c * pitch / t


def buckling_critical_elastic(t, height, e_mod, k=K_DEFAULT):
    """식 (4): 탄성 판 기둥 임계응력 σ_cr,E [MPa]."""
    _check_positive(t=t, height=height, e_mod=e_mod, k=k)
    return math.pi ** 2 * e_mod / (12.0 * (1.0 - NU ** 2)) * (t / (k * height)) ** 2


def buckling_critical(t, height, e_mod, sy, k=K_DEFAULT):
    """식 (4'): Johnson 포물선 천이 적용 임계응력 σ_cr [MPa]."""
    scr_e = buckling_critical_elastic(t, height, e_mod, k)
    if scr_e <= sy / 2.0:
        return scr_e
    return sy * (1.0 - sy / (4.0 * scr_e))


def compression_allowable(t, height, e_mod, sy, s_allow, k=K_DEFAULT):
    """식 (5) 우변: min(σ_cr/DF, S) [MPa]."""
    return min(buckling_critical(t, height, e_mod, sy, k) / DF_BUCKLING, s_allow)


def fin_min_thickness_compression(dp_c, pitch, height, e_mod, sy, s_allow,
                                  k=K_DEFAULT, tol=1e-6):
    """식 (3)&(5)를 만족하는 최소 핀 두께 [mm] (이분법).

    g(t) = ΔP_c·p/t - min(σ_cr(t)/DF, S) 는 t에 강단조감소(σ_c 감소·허용 증가)
    이므로 근이 유일하다. g(hi)>0 은 σ_c(hi)→ΔP_c 수준의 유한값이므로 피치가
    지나치게 작지 않는 한 발생하지 않는다.
    """
    _check_nonneg(dp_c=dp_c)
    if dp_c == 0:
        return 0.0
    _check_positive(pitch=pitch, height=height)

    def g(t):
        return dp_c * pitch / t - compression_allowable(
            t, height, e_mod, sy, s_allow, k)

    lo, hi = tol, pitch * (1 - 1e-9)
    if g(hi) > 0:
        raise ValueError(
            "피치 이내에서 압축 기준을 만족하는 두께가 없음 (ΔP_c 대비 기하 재검토 필요)")
    if g(lo) <= 0:
        return 0.0
    for _ in range(200):
        mid = 0.5 * (lo + hi)
        if g(mid) > 0:
            lo = mid
        else:
            hi = mid
        if hi - lo < tol:
            break
    return hi


def sheet_design_moment(w, span, pitch_bar, t_fin_eff, p_min):
    """식 (6): M_d = w·s²/12 + P_min·(p̄-t_f)·p̄/8 [N·mm/mm]."""
    _check_nonneg(w=w, p_min=p_min)
    _check_positive(span=span, pitch_bar=pitch_bar)
    if not 0 <= t_fin_eff < pitch_bar:
        raise ValueError(f"t_fin_eff={t_fin_eff} 는 0 ≤ t < p̄({pitch_bar}) 이어야 함")
    return w * span ** 2 / 12.0 + p_min * (pitch_bar - t_fin_eff) * pitch_bar / 8.0


def sheet_min_thickness(m_design, s_allow):
    """식 (7): t_ps >= sqrt(4·M_d/S) [mm] (σ_b ≤ 1.5S)."""
    _check_nonneg(m_design=m_design)
    _check_positive(s_allow=s_allow)
    return math.sqrt(4.0 * m_design / s_allow)


def _range_note(value, lo_hi, label):
    lo, hi = lo_hi
    if value < lo or value > hi:
        return f"  ⚠️ {label} {value:.3g} mm 는 ALPEMA 표준범위 {lo}–{hi} mm 밖"
    return f"  ✓ {label} {value:.3g} mm — ALPEMA 표준범위 {lo}–{hi} mm 내"


def report(args):
    if args.temp > T_MAX:
        raise ValueError(
            f"설계온도 {args.temp} °C 는 3003 재질 한계 {T_MAX} °C 초과 "
            "(ALPEMA Table 6-1) — 사용자 S 테이블과 무관하게 허용되지 않음")
    s = allowable_stress(args.temp, args.s_table)
    sy = yield_strength(args.temp)
    e = youngs_modulus(args.temp)
    s_label = "[사용자 제공 S 테이블]" if args.s_table else "[3003-O, ASME Div.1 대표값]"
    lines = []
    add = lines.append
    add("=" * 64)
    add("PFHE Plain 핀 적정두께 계산 (예비설계용, Rev.1)")
    add("=" * 64)
    add(f"입력: ΔP_t={args.dp_t} MPa, ΔP_c={args.dp_c} MPa, T={args.temp} °C")
    add(f"      피치 p={args.pitch} mm, 핀 높이 h={args.height} mm, K={args.k}, "
        f"스테이계수 f={args.stay_factor}")
    add(f"물성: S(T)={s:.2f}, Sy(T)={sy:.2f}, E(T)={e:.0f} MPa  {s_label}")
    add("-" * 64)

    t_tension = fin_min_thickness_tension(args.dp_t, args.pitch, s, args.stay_factor)
    add(f"[1] 핀 인장  최소두께 (식 2)  : t ≥ {t_tension:.4f} mm"
        + (f"  (UG-50 계수 {args.stay_factor} 반영)" if args.stay_factor != 1.0 else ""))

    t_comp = fin_min_thickness_compression(
        args.dp_c, args.pitch, args.height, e, sy, s, args.k)
    if args.dp_c > 0:
        add(f"[2] 핀 좌굴  최소두께 (식 3~5): t ≥ {t_comp:.4f} mm  (Johnson 천이 적용)")
    else:
        add("[2] 핀 좌굴  : ΔP_c=0 → 해당 없음")

    t_req_eff = max(t_tension, t_comp)          # 유효(최소 보증) 두께 기준
    if args.tol_factor < 1.0:
        t_req_nom = t_req_eff / args.tol_factor  # 공칭두께 환산
        add(f"[3] 요구두께 : 유효 t ≥ {t_req_eff:.4f} mm → 공차계수 "
            f"{args.tol_factor} 반영 공칭 t ≥ {t_req_nom:.4f} mm")
    else:
        t_req_nom = t_req_eff
        add(f"[3] 요구두께 : t_req = max(인장, 좌굴) = {t_req_nom:.4f} mm")
    add(_range_note(t_req_nom, FIN_T_RANGE, "요구 핀 두께(공칭)"))
    add(_range_note(args.pitch, FIN_PITCH_RANGE, "핀 피치"))
    add(_range_note(args.height, FIN_H_RANGE, "핀 높이"))

    if args.t_fin:
        t_eff = args.t_fin * args.tol_factor     # 합부 판정은 유효두께 기준 (문서 §8-5)
        add("-" * 64)
        add(f"[검증] 선정 핀 두께 공칭 t={args.t_fin} mm"
            + (f" → 유효 t={t_eff:.4f} mm (공차계수 {args.tol_factor})"
               if args.tol_factor < 1.0 else "") + ":")
        if t_eff >= args.pitch:
            raise ValueError(f"선정 핀 두께 {args.t_fin} mm 가 피치 {args.pitch} mm 이상")
        st = fin_stress_tension(args.dp_t * args.stay_factor, args.pitch, t_eff) \
            if args.dp_t > 0 else 0.0
        ok_t = st <= s
        add(f"  인장 σ_t={st:.2f} MPa vs S={s:.2f} MPa → {'합격 ✓' if ok_t else '불합격 ✗'}")
        if args.dp_c > 0:
            sc = fin_stress_compression(args.dp_c, args.pitch, t_eff)
            sa = compression_allowable(t_eff, args.height, e, sy, s, args.k)
            scr = buckling_critical(t_eff, args.height, e, sy, args.k)
            ok_c = sc <= sa
            add(f"  압축 σ_c={sc:.2f} MPa vs min(σ_cr/{DF_BUCKLING:.0f}, S)="
                f"{sa:.2f} MPa (σ_cr={scr:.1f}, Johnson) → {'합격 ✓' if ok_c else '불합격 ✗'}")

    if args.w_sheet > 0 or args.p_min > 0:
        pitch_bar = max(args.pitch, args.pitch_adj or 0.0)
        t_fin_eff = (args.t_fin * args.tol_factor) if args.t_fin else t_req_eff
        span = pitch_bar - t_fin_eff
        if span <= 0:
            raise ValueError(f"파팅시트 스팬이 0 이하 (p̄={pitch_bar}, t_eff={t_fin_eff})")
        m_d = sheet_design_moment(args.w_sheet, span, pitch_bar, t_fin_eff, args.p_min)
        tps = sheet_min_thickness(m_d, s)
        add("-" * 64)
        add(f"[4] 파팅시트 (식 6~7): p̄=max(자층, 인접층 피치)={pitch_bar} mm, "
            f"스팬 s={span:.3f} mm")
        add(f"    차압 w={args.w_sheet} MPa, 균형가압 P_min={args.p_min} MPa "
            f"→ M_d={m_d:.4f} N·mm/mm")
        add(f"    t_ps ≥ {tps:.4f} mm")
        add(_range_note(tps, SHEET_T_RANGE, "파팅시트 소요두께"))
        if tps < SHEET_T_RANGE[0]:
            add(f"    → ALPEMA 표준 최소 {SHEET_T_RANGE[0]} mm 적용 권장")
        if not args.pitch_adj:
            add("    ※ 인접층 피치 미입력(--pitch-adj): 자층 피치만 사용 — 인접층이 더")
            add("      성기면 비보수적이므로 반드시 양면 중 성긴 피치를 반영할 것 (문서 §6.1)")

    add("=" * 64)
    add("주의: S/Sy/E 테이블은 대표값 — 적용 코드 연도판으로 검증할 것.")
    add("인증 설계는 ALPEMA 5.15.1.1(승인 계산법)/5.15.1.2(파열시험)에 따름.")
    return "\n".join(lines)


def selftest():
    """유도문서 예제 1~3 재현 + 경계조건 검사."""
    # --- 물성 보간 ---
    assert abs(allowable_stress(65) - 22.8) < 1e-9
    assert abs(allowable_stress(20) - 22.8) < 1e-9          # 하한 미만 → 첫 값
    assert abs(allowable_stress(107) - 22.0) < 0.01          # 93~121 중간
    assert abs(yield_strength(93) - 33.1) < 1e-9
    assert abs(youngs_modulus(93) - 66200) < 1e-6
    for bad in (205, float("nan")):
        try:
            allowable_stress(bad)
            raise AssertionError(f"{bad} 가 허용됨")
        except ValueError:
            pass

    # --- 예제 1: 인장 ---
    s65 = allowable_stress(65)
    t1 = fin_min_thickness_tension(5.0, 1.4, s65)
    assert abs(t1 - 0.2518) < 0.001, t1
    assert abs(fin_stress_tension(5.0, 1.4, 0.30) - 18.333) < 0.01
    assert abs(fin_stress_tension(5.0, 1.4, t1) - s65) < 1e-9   # 자기일관성
    # UG-50 스테이 계수형 (식 2''): f=1.1
    t1s = fin_min_thickness_tension(5.0, 1.4, s65, 1.10)
    assert t1s > t1 and abs(t1s - 1.1 * 5.0 * 1.4 / (s65 + 1.1 * 5.0)) < 1e-12
    # 입력 검증
    for bad_call in (lambda: fin_min_thickness_tension(-1, 1.4, s65),
                     lambda: fin_min_thickness_tension(float("nan"), 1.4, s65),
                     lambda: fin_min_thickness_tension(5.0, -1, s65),
                     lambda: fin_stress_tension(5.0, 1.4, 1.5)):
        try:
            bad_call()
            raise AssertionError("잘못된 입력이 통과됨")
        except ValueError:
            pass

    # --- 예제 2: 압축·좌굴 (전피치 하중폭 + Johnson) ---
    e93, sy93, s93 = youngs_modulus(93), yield_strength(93), allowable_stress(93)
    sc = fin_stress_compression(0.6, 2.0, 0.40)
    assert abs(sc - 3.0) < 1e-9, sc                          # σ_c = ΔP·p/t
    scr_e = buckling_critical_elastic(0.40, 9.5, e93, 0.7)
    assert abs(scr_e - 221.0) < 2.0, scr_e
    scr = buckling_critical(0.40, 9.5, e93, sy93, 0.7)       # Johnson
    assert abs(scr - 31.86) < 0.1, scr
    assert scr < sy93                                        # Johnson 은 Sy 를 넘지 않음
    sa = compression_allowable(0.40, 9.5, e93, sy93, s93, 0.7)
    assert abs(sa - scr / 3.0) < 1e-9 and sc <= sa
    # 탄성 영역 연속성: σ_cr,E ≤ Sy/2 이면 탄성값 그대로
    t_slender = 0.05
    assert buckling_critical_elastic(t_slender, 9.5, e93, 0.7) <= sy93 / 2
    assert abs(buckling_critical(t_slender, 9.5, e93, sy93, 0.7)
               - buckling_critical_elastic(t_slender, 9.5, e93, 0.7)) < 1e-9
    # 좌굴 최소두께: 근에서 g(t)=0
    tc = fin_min_thickness_compression(0.6, 2.0, 9.5, e93, sy93, s93, 0.7)
    lhs = fin_stress_compression(0.6, 2.0, tc)
    rhs = compression_allowable(tc, 9.5, e93, sy93, s93, 0.7)
    assert abs(lhs - rhs) < 0.05, (lhs, rhs)

    # --- 예제 3: 파팅시트 (케이스 A/B) ---
    m_a = sheet_design_moment(5.0, 1.10, 1.4, 0.30, 0.0)     # 단독가압
    assert abs(m_a - 5.0 * 1.10 ** 2 / 12) < 1e-9
    tps_a = sheet_min_thickness(m_a, s65)
    assert abs(tps_a - 0.2975) < 0.001, tps_a
    sigma_b = 6 * m_a / tps_a ** 2                            # σ_b == 1.5S 자기일관성
    assert abs(sigma_b - 1.5 * s65) < 1e-6
    m_b = sheet_design_moment(0.0, 1.10, 1.4, 0.30, 5.0)     # 균형가압+어긋남
    assert abs(m_b - 0.9625) < 1e-9
    tps_b = sheet_min_thickness(m_b, s65)
    assert abs(tps_b - 0.4109) < 0.001, tps_b
    assert tps_b > tps_a                                      # 케이스 B 지배 확인

    # --- 극한 거동 통일: ΔP→0 이면 두께→0 ---
    assert fin_min_thickness_tension(0.0, 1.4, s65) == 0.0
    assert fin_min_thickness_compression(0.0, 1.4, 6.5, e93, sy93, s65) == 0.0
    print("selftest: OK (모든 검증 통과)")


def _parse_s_table(text):
    try:
        rows = json.loads(text)
        table = [(float(a), float(b)) for a, b in rows]
    except (ValueError, TypeError) as exc:
        raise ValueError(f"--s-table JSON 형식 오류: {exc}")
    table.sort(key=lambda r: r[0])
    _validate_table(table, "--s-table")
    return table


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="PFHE Plain 핀 적정두께 계산기 (유도: fin_thickness_derivation.md)")
    ap.add_argument("--dp-t", type=float, default=0.0,
                    help="핀 인장 유발 차압 ΔP_t [MPa] (통상 = 층 설계압력)")
    ap.add_argument("--dp-c", type=float, default=0.0,
                    help="핀 압축 유발 차압 ΔP_c [MPa] (인접층 고압/자층 감압)")
    ap.add_argument("--temp", type=float, default=40.0, help="설계온도 [°C] (≤ 204)")
    ap.add_argument("--pitch", type=float, help="자층 핀 피치 p [mm]")
    ap.add_argument("--pitch-adj", type=float, default=0.0,
                    help="인접층 핀 피치 [mm] — 파팅시트 스팬은 max(자층, 인접층) 기준")
    ap.add_argument("--height", type=float, default=6.5, help="핀 높이 h [mm]")
    ap.add_argument("--k", type=float, default=K_DEFAULT,
                    help=f"좌굴 유효길이계수 K (기본 {K_DEFAULT}, 구속 불확실 시 1.0)")
    ap.add_argument("--t-fin", type=float, default=0.0,
                    help="선정 핀 두께(공칭) [mm] — 지정 시 유효두께 기준 합부 검증")
    ap.add_argument("--w-sheet", type=float, default=0.0,
                    help="파팅시트 양면 순 차압 w=|P_A-P_B| [MPa]")
    ap.add_argument("--p-min", type=float, default=0.0,
                    help="양면 동시가압 시 낮은 쪽 압력 P_min [MPa] — 핀 어긋남 교번하중 항")
    ap.add_argument("--stay-factor", type=float, default=1.0,
                    help="인장 스테이 계수 f (UG-50(b) 준용 시 1.10, 기본 1.0)")
    ap.add_argument("--tol-factor", type=float, default=1.0,
                    help="제작 공차 계수 (0<f≤1, 예: 0.9 → 공칭두께의 90%%만 유효 가정)")
    ap.add_argument("--s-table", type=str, default=None,
                    help='허용응력 테이블 교체용 JSON: [[T°C, S MPa], ...] (온도 오름차순)')
    ap.add_argument("--selftest", action="store_true", help="내부 검증 실행")
    args = ap.parse_args(argv)

    if args.selftest:
        selftest()
        return 0
    if args.pitch is None:
        ap.error("--pitch 는 필수입니다 (--selftest 제외)")
    # 입력 검증 (NaN 포함: 비교식이 False 가 되도록 not-형식 사용)
    if not (args.pitch > 0 and math.isfinite(args.pitch)):
        ap.error(f"--pitch 는 양의 유한값이어야 합니다: {args.pitch}")
    if not (args.height > 0 and math.isfinite(args.height)):
        ap.error(f"--height 는 양의 유한값이어야 합니다: {args.height}")
    if not (args.k > 0 and math.isfinite(args.k)):
        ap.error(f"--k 는 양의 유한값이어야 합니다: {args.k}")
    if not (0.0 < args.tol_factor <= 1.0):
        ap.error(f"--tol-factor 는 0 < f ≤ 1 이어야 합니다: {args.tol_factor}")
    if not (args.stay_factor >= 1.0 and math.isfinite(args.stay_factor)):
        ap.error(f"--stay-factor 는 1.0 이상이어야 합니다: {args.stay_factor}")
    for name in ("dp_t", "dp_c", "w_sheet", "p_min", "pitch_adj", "t_fin"):
        v = getattr(args, name)
        if not (v >= 0 and math.isfinite(v)):
            ap.error(f"--{name.replace('_', '-')} 는 0 이상 유한값이어야 합니다: {v}")
    if args.t_fin and not args.t_fin < args.pitch:
        ap.error(f"--t-fin({args.t_fin}) 은 피치({args.pitch})보다 작아야 합니다")
    if args.dp_t <= 0 and args.dp_c <= 0:
        ap.error("--dp-t 또는 --dp-c 중 하나는 0보다 커야 합니다")
    try:
        args.s_table = _parse_s_table(args.s_table) if args.s_table else None
        print(report(args))
    except ValueError as exc:
        print(f"오류: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
