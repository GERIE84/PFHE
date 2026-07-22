#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PFHE Plain 핀 적정두께 계산기.

차압(ΔP)과 설계온도(T)로부터 Plain 핀의 최소 요구두께와 파팅시트 최소두께를
계산한다. 유도 근거는 같은 폴더의 fin_thickness_derivation.md 참조.

  - 핀 인장:        t >= ΔP_t · p / (S(T) + ΔP_t)                    ... 식 (2)
  - 핀 압축·좌굴:   ΔP_c(p-t)/t <= min(σ_cr/3, S(T)),
                    σ_cr = π²E(T)/(12(1-ν²)) · (t/(K·h))²            ... 식 (3)~(5)
  - 파팅시트 굽힘:  t_ps >= s · sqrt(w / (3·S(T)))                    ... 식 (7)

허용응력 S(T): SB-209 3003-O, ASME Sec.VIII Div.1 (Sec.II-D Table 1B) 대표값.
⚠️ 실제 설계 전 적용 코드 연도판 값으로 교체할 것 (--s-table 옵션 또는
ALLOWABLE_S_TABLE 수정). 본 계산은 예비설계용이며, 인증 설계는 ALPEMA
5.15.1.1(승인 계산법) 또는 5.15.1.2(파열시험)에 따른다.

사용 예:
    python3 fin_thickness_calc.py --dp-t 5.0 --temp 65 --pitch 1.4 --height 6.5
    python3 fin_thickness_calc.py --dp-t 5.0 --dp-c 0.6 --temp 93 \\
        --pitch 2.0 --height 9.5 --t-fin 0.4 --w-sheet 5.0
    python3 fin_thickness_calc.py --selftest
"""

import argparse
import json
import sys

NU = 0.33          # 알루미늄 포아송비
DF_BUCKLING = 3.0  # 좌굴 설계계수 (ASME Div.1 UG-28 철학 준용)
K_DEFAULT = 0.7    # 유효길이계수 (브레이즈 구속, 실무 권장)

# ASME Sec.II-D Table 1B — SB-209 3003-O 허용응력 대표값 [(온도 °C, S MPa)]
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

# ASME Sec.II-D TM — 알루미늄 탄성계수 [(온도 °C, E MPa)]
E_TABLE = [
    (25.0, 68900.0),
    (93.0, 66200.0),
    (149.0, 63400.0),
    (204.0, 60000.0),
]

T_MAX = 204.0  # ALPEMA Table 6-1: 3003 최대 적용 설계온도(ASME) [°C]

# ALPEMA 5.11 표준 제작 범위 [mm]
FIN_T_RANGE = (0.15, 0.7)
FIN_PITCH_RANGE = (1.0, 4.5)
FIN_H_RANGE = (2.0, 12.0)
SHEET_T_RANGE = (0.8, 2.0)


def _interp(table, x, name):
    """테이블 선형 보간. 하한 미만은 첫 값 사용, 상한 초과는 오류."""
    if x > table[-1][0]:
        raise ValueError(
            f"{name}: {x} °C 는 테이블 상한 {table[-1][0]} °C (ALPEMA/ASME 한계) 초과")
    if x <= table[0][0]:
        return table[0][1]
    for (x0, y0), (x1, y1) in zip(table, table[1:]):
        if x0 <= x <= x1:
            return y0 + (y1 - y0) * (x - x0) / (x1 - x0)
    raise AssertionError("unreachable")


def allowable_stress(temp_c, table=None):
    """설계온도[°C] → 허용응력 S(T) [MPa]."""
    return _interp(table or ALLOWABLE_S_TABLE, temp_c, "S(T)")


def youngs_modulus(temp_c):
    """설계온도[°C] → 탄성계수 E(T) [MPa]."""
    return _interp(E_TABLE, temp_c, "E(T)")


def fin_stress_tension(dp_t, pitch, t):
    """식 (1): 핀 인장응력 σ_t = ΔP_t(p-t)/t [MPa]."""
    if not 0 < t < pitch:
        raise ValueError(f"핀 두께 t={t} 는 0 < t < 피치({pitch}) 이어야 함")
    return dp_t * (pitch - t) / t


def fin_min_thickness_tension(dp_t, pitch, s_allow):
    """식 (2): 인장 기준 최소 핀 두께 [mm]."""
    return dp_t * pitch / (s_allow + dp_t)


def buckling_critical_stress(t, height, e_mod, k=K_DEFAULT):
    """식 (4): 판 기둥 좌굴 임계응력 σ_cr [MPa]."""
    import math
    return math.pi ** 2 * e_mod / (12.0 * (1.0 - NU ** 2)) * (t / (k * height)) ** 2


def compression_allowable(t, height, e_mod, s_allow, k=K_DEFAULT):
    """식 (5) 우변: min(σ_cr/DF, S) [MPa]."""
    return min(buckling_critical_stress(t, height, e_mod, k) / DF_BUCKLING, s_allow)


def fin_min_thickness_compression(dp_c, pitch, height, e_mod, s_allow,
                                  k=K_DEFAULT, tol=1e-6):
    """식 (3)&(5)를 만족하는 최소 핀 두께 [mm] (이분법).

    g(t) = ΔP_c(p-t)/t - min(σ_cr(t)/DF, S) 는 t 증가에 단조감소하므로
    g(t)=0 의 근을 [tol, pitch) 에서 탐색한다.
    """
    if dp_c <= 0:
        return 0.0

    def g(t):
        return dp_c * (pitch - t) / t - compression_allowable(
            t, height, e_mod, s_allow, k)

    lo, hi = tol, pitch * (1 - 1e-9)
    if g(hi) > 0:
        raise ValueError("피치 이내에서 압축 기준을 만족하는 두께가 없음 (기하 재검토 필요)")
    if g(lo) <= 0:
        return lo
    for _ in range(200):
        mid = 0.5 * (lo + hi)
        if g(mid) > 0:
            lo = mid
        else:
            hi = mid
        if hi - lo < tol:
            break
    return hi


def sheet_min_thickness(span, w, s_allow):
    """식 (7): 파팅시트 최소두께 t_ps >= s·sqrt(w/(3S)) [mm]."""
    import math
    if w <= 0:
        return 0.0
    return span * math.sqrt(w / (3.0 * s_allow))


def _range_note(value, lo_hi, label):
    lo, hi = lo_hi
    if value < lo or value > hi:
        return f"  ⚠️ {label} {value:.3g} mm 는 ALPEMA 표준범위 {lo}–{hi} mm 밖"
    return f"  ✓ {label} {value:.3g} mm — ALPEMA 표준범위 {lo}–{hi} mm 내"


def report(args):
    s = allowable_stress(args.temp, args.s_table)
    e = youngs_modulus(args.temp)
    lines = []
    add = lines.append
    add("=" * 64)
    add("PFHE Plain 핀 적정두께 계산 (예비설계용)")
    add("=" * 64)
    add(f"입력: ΔP_t={args.dp_t} MPa, ΔP_c={args.dp_c} MPa, T={args.temp} °C")
    add(f"      피치 p={args.pitch} mm, 핀 높이 h={args.height} mm, K={args.k}")
    add(f"물성: S(T)={s:.2f} MPa, E(T)={e:.0f} MPa  [3003-O, ASME Div.1 대표값]")
    add("-" * 64)

    t_tension = fin_min_thickness_tension(args.dp_t, args.pitch, s)
    add(f"[1] 핀 인장  최소두께 (식 2)  : t ≥ {t_tension:.4f} mm")

    t_comp = fin_min_thickness_compression(
        args.dp_c, args.pitch, args.height, e, s, args.k)
    if args.dp_c > 0:
        add(f"[2] 핀 좌굴  최소두께 (식 3~5): t ≥ {t_comp:.4f} mm")
    else:
        add("[2] 핀 좌굴  : ΔP_c=0 → 해당 없음")

    t_req = max(t_tension, t_comp)
    if args.tol_factor < 1.0:
        t_req_nom = t_req / args.tol_factor
        add(f"[3] 요구두께 : t_req = {t_req:.4f} mm "
            f"(공차계수 {args.tol_factor} 반영 시 공칭 t ≥ {t_req_nom:.4f} mm)")
        t_req = t_req_nom
    else:
        add(f"[3] 요구두께 : t_req = max(인장, 좌굴) = {t_req:.4f} mm")
    add(_range_note(t_req, FIN_T_RANGE, "요구 핀 두께"))
    add(_range_note(args.pitch, FIN_PITCH_RANGE, "핀 피치"))
    add(_range_note(args.height, FIN_H_RANGE, "핀 높이"))

    if args.t_fin:
        st = fin_stress_tension(args.dp_t, args.pitch, args.t_fin)
        ok_t = st <= s
        add("-" * 64)
        add(f"[검증] 선정 핀 두께 t={args.t_fin} mm:")
        add(f"  인장 σ_t={st:.2f} MPa vs S={s:.2f} MPa → {'합격 ✓' if ok_t else '불합격 ✗'}")
        if args.dp_c > 0:
            sc = args.dp_c * (args.pitch - args.t_fin) / args.t_fin
            sa = compression_allowable(args.t_fin, args.height, e, s, args.k)
            scr = buckling_critical_stress(args.t_fin, args.height, e, args.k)
            ok_c = sc <= sa
            add(f"  압축 σ_c={sc:.2f} MPa vs min(σ_cr/{DF_BUCKLING:.0f}, S)="
                f"{sa:.2f} MPa (σ_cr={scr:.1f}) → {'합격 ✓' if ok_c else '불합격 ✗'}")

    if args.w_sheet > 0:
        t_fin_for_span = args.t_fin if args.t_fin else t_req
        span = args.pitch - t_fin_for_span
        tps = sheet_min_thickness(span, args.w_sheet, s)
        add("-" * 64)
        add(f"[4] 파팅시트 (식 7): 스팬 s={span:.3f} mm, w={args.w_sheet} MPa")
        add(f"    t_ps ≥ {tps:.4f} mm")
        add(_range_note(max(tps, SHEET_T_RANGE[0]), SHEET_T_RANGE, "파팅시트 두께(하한 적용)"))

    add("=" * 64)
    add("주의: S(T) 테이블은 대표값 — 적용 코드 연도판으로 검증할 것.")
    add("인증 설계는 ALPEMA 5.15.1.1(승인 계산법)/5.15.1.2(파열시험)에 따름.")
    return "\n".join(lines)


def selftest():
    """유도문서 예제 1~3 재현 + 경계조건 검사."""
    import math
    # S(T), E(T) 보간
    assert abs(allowable_stress(65) - 22.8) < 1e-9
    assert abs(allowable_stress(20) - 22.8) < 1e-9          # 하한 미만 → 첫 값
    assert abs(allowable_stress(107) - 22.0) < 0.01          # 93~121 중간
    assert abs(youngs_modulus(93) - 66200) < 1e-6
    try:
        allowable_stress(205)
        raise AssertionError("204 °C 초과가 허용됨")
    except ValueError:
        pass

    # 예제 1: 인장
    t1 = fin_min_thickness_tension(5.0, 1.4, allowable_stress(65))
    assert abs(t1 - 0.2518) < 0.001, t1
    assert abs(fin_stress_tension(5.0, 1.4, 0.30) - 18.333) < 0.01

    # 식 (2) 자기일관성: t=t_min 에서 σ_t == S
    s65 = allowable_stress(65)
    assert abs(fin_stress_tension(5.0, 1.4, t1) - s65) < 1e-9

    # 예제 2: 좌굴
    e93 = youngs_modulus(93)
    scr = buckling_critical_stress(0.40, 9.5, e93, 0.7)
    assert abs(scr - 221.0) < 2.0, scr
    sc = 0.6 * (2.0 - 0.40) / 0.40
    assert abs(sc - 2.4) < 1e-9
    assert sc <= compression_allowable(0.40, 9.5, e93, allowable_stress(93), 0.7)

    # 좌굴 최소두께: 근에서 g(t)=0 성립 확인
    tc = fin_min_thickness_compression(0.6, 2.0, 9.5, e93, allowable_stress(93), 0.7)
    lhs = 0.6 * (2.0 - tc) / tc
    rhs = compression_allowable(tc, 9.5, e93, allowable_stress(93), 0.7)
    assert abs(lhs - rhs) < 0.01, (lhs, rhs)

    # 예제 3: 파팅시트
    tps = sheet_min_thickness(1.10, 5.0, s65)
    assert abs(tps - 0.2975) < 0.001, tps
    # 식 (7) 자기일관성: t_ps=t_min 에서 σ_b == 1.5S
    sigma_b = 5.0 * 1.10 ** 2 / (2.0 * tps ** 2)
    assert abs(sigma_b - 1.5 * s65) < 1e-6

    # 극한: ΔP→0 이면 두께→0
    assert fin_min_thickness_tension(0.0, 1.4, s65) == 0.0
    assert fin_min_thickness_compression(0.0, 1.4, 6.5, e93, s65) == 0.0
    print("selftest: OK (모든 검증 통과)")


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="PFHE Plain 핀 적정두께 계산기 (유도: fin_thickness_derivation.md)")
    ap.add_argument("--dp-t", type=float, default=0.0,
                    help="핀 인장 유발 차압 ΔP_t [MPa] (통상 = 층 설계압력)")
    ap.add_argument("--dp-c", type=float, default=0.0,
                    help="핀 압축 유발 차압 ΔP_c [MPa] (인접층 고압/자층 감압)")
    ap.add_argument("--temp", type=float, default=40.0, help="설계온도 [°C]")
    ap.add_argument("--pitch", type=float, help="핀 피치 p [mm]")
    ap.add_argument("--height", type=float, default=6.5, help="핀 높이 h [mm]")
    ap.add_argument("--k", type=float, default=K_DEFAULT,
                    help=f"좌굴 유효길이계수 K (기본 {K_DEFAULT}, 보수적 1.0)")
    ap.add_argument("--t-fin", type=float, default=0.0,
                    help="선정 핀 두께 [mm] — 지정 시 합부 검증 수행")
    ap.add_argument("--w-sheet", type=float, default=0.0,
                    help="파팅시트 양면 순 차압 w [MPa] — 지정 시 식(7) 계산")
    ap.add_argument("--tol-factor", type=float, default=1.0,
                    help="제작 공차 계수 (예: 0.9 → 공칭두께의 90%%만 유효 가정)")
    ap.add_argument("--s-table", type=str, default=None,
                    help='허용응력 테이블 교체용 JSON: [[T°C, S MPa], ...]')
    ap.add_argument("--selftest", action="store_true", help="내부 검증 실행")
    args = ap.parse_args(argv)

    if args.selftest:
        selftest()
        return 0
    if args.pitch is None:
        ap.error("--pitch 는 필수입니다 (--selftest 제외)")
    if args.dp_t <= 0 and args.dp_c <= 0:
        ap.error("--dp-t 또는 --dp-c 중 하나는 0보다 커야 합니다")
    args.s_table = ([(float(a), float(b)) for a, b in json.loads(args.s_table)]
                    if args.s_table else None)
    print(report(args))
    return 0


if __name__ == "__main__":
    sys.exit(main())
