"""
engine_ddc.py - Mesin Kalkulasi Logistik DDC (Pure Python - Tanpa C++ / Pandas)
===============================================================================
Mengimplementasikan seluruh aturan bisnis Project Logistic DDC menggunakan
murni Python standar sehingga kompatibel 100% dengan Python 3.11 s.d. 3.15+ di Windows.
"""

import math
from db_excel import PLANTS, EXPEDITIONS, parse_date_str, format_date_d_mmm_yy

MONTH_NAMES_ID = {
    1: "JANUARY", 2: "FEBRUARY", 3: "MARCH", 4: "APRIL",
    5: "MAY", 6: "JUNE", 7: "JULY", 8: "AUGUST",
    9: "SEPTEMBER", 10: "OCTOBER", 11: "NOVEMBER", 12: "DECEMBER"
}


def get_last_3_months(target_month, target_year):
    """Menentukan 3 bulan terakhir sebelum bulan berjalan."""
    periods = []
    m, y = int(target_month), int(target_year)
    for _ in range(3):
        m -= 1
        if m < 1:
            m = 12
            y -= 1
        periods.append((y, m))
    periods.reverse()
    return periods


def calculate_monthly_quota(
    sl_records,
    target_month=10,
    target_year=2026,
    cutoff_date=None,
    start_date=None,
    end_date=None,
    gudang_filter="ALL",
):
    """
    Menghitung Kuota Bulanan (%) berdasarkan pencapaian aktual SL dalam rentang tanggal
    (mendukung 2 kolom kalender: start_date s.d. end_date, atau rentang 3 bulan terakhir dengan cut-off).
    """
    clean_start = parse_date_str(start_date) if start_date else None
    clean_end = parse_date_str(end_date) if end_date else (parse_date_str(cutoff_date) if cutoff_date else None)
    g_filter = (gudang_filter or "ALL").strip().upper()

    default_periods = get_last_3_months(target_month, target_year)
    default_period_set = set(default_periods)

    if clean_start or clean_end:
        if not clean_start and clean_end:
            first_y, first_m = default_periods[0]
            clean_start = f"{first_y:04d}-{first_m:02d}-01"
        filtered = []
        all_gudang_set = set()
        for r in sl_records:
            all_gudang_set.add(r["GUDANG"])
            t = r["TANGGAL_STR"]
            if clean_start and t < clean_start:
                continue
            if clean_end and t > clean_end:
                continue
            filtered.append(r)

        unique_periods = sorted(list({(int(r["YEAR"]), int(r["MONTH"])) for r in filtered}))
        periods = unique_periods if unique_periods else default_periods
    else:
        periods = default_periods
        filtered = []
        all_gudang_set = set()
        for r in sl_records:
            all_gudang_set.add(r["GUDANG"])
            if (int(r["YEAR"]), int(r["MONTH"])) not in default_period_set:
                continue
            filtered.append(r)

    # Filter sesuai gudang_filter untuk tabel utama
    main_records = [
        r for r in filtered if (g_filter == "ALL" or r["GUDANG"] == g_filter)
    ]

    unit_matrix = {x: {p: 0 for p in PLANTS} for x in EXPEDITIONS}
    for r in main_records:
        x = r["XPDC"]
        p = r["PLANT"]
        if x in unit_matrix and p in unit_matrix[x]:
            unit_matrix[x][p] += int(r["QTY"])

    plant_totals = {
        p: sum(unit_matrix[x][p] for x in EXPEDITIONS) for p in PLANTS
    }
    total_3m_units = sum(plant_totals.values())

    for x in EXPEDITIONS:
        gt = sum(unit_matrix[x][p] for p in PLANTS)
        unit_matrix[x]["GRAND_TOTAL"] = gt
        unit_matrix[x]["PCT_ALL_PLANT"] = (
            (gt / total_3m_units) if total_3m_units > 0 else 0.0
        )

    pct_matrix = {x: {} for x in EXPEDITIONS}
    for x in EXPEDITIONS:
        for p in PLANTS:
            p_tot = plant_totals[p]
            pct_matrix[x][p] = (
                (unit_matrix[x][p] / p_tot) if p_tot > 0 else 0.0
            )
        pct_matrix[x]["TARGET_PCT"] = unit_matrix[x]["PCT_ALL_PLANT"]

    # Rincian per Gudang (KARAWANG, BANDUNG, CIREBON)
    gudang_breakdown = {}
    for g in sorted(all_gudang_set):
        g_units = {x: 0 for x in EXPEDITIONS}
        for r in filtered:
            if r["GUDANG"] == g and r["XPDC"] in g_units:
                g_units[r["XPDC"]] += int(r["QTY"])
        g_tot = sum(g_units.values())
        gudang_breakdown[g] = {
            "units": g_units,
            "total_units": g_tot,
            "pct": {
                x: round(g_units[x] / g_tot, 6) if g_tot > 0 else 0.0
                for x in EXPEDITIONS
            },
        }

    start_disp = format_date_d_mmm_yy(clean_start) if clean_start else "-"
    end_disp = format_date_d_mmm_yy(clean_end) if clean_end else "-"
    date_range_display = f"{start_disp} s.d. {end_disp}" if (clean_start and clean_end) else (f"s.d. {end_disp}" if clean_end else "Akhir Bulan Penuh")

    return {
        "target_month": int(target_month),
        "target_month_name": MONTH_NAMES_ID.get(int(target_month), str(target_month)),
        "target_year": int(target_year),
        "periods": [f"{MONTH_NAMES_ID.get(m, m)} {y}" for (y, m) in periods],
        "start_date": clean_start,
        "end_date": clean_end,
        "cutoff_date": clean_end or "Akhir Bulan Penuh",
        "date_range_display": date_range_display,
        "total_3m_units": total_3m_units,
        "plant_totals": plant_totals,
        "unit_matrix": unit_matrix,
        "pct_matrix": pct_matrix,
        "gudang_breakdown": gudang_breakdown,
    }


def _round_preserving_sum(float_vals_dict, target_int_sum):
    """Membulatkan dictionary angka desimal menjadi integer dengan tetap menjaga jumlah total = target_int_sum."""
    keys = list(float_vals_dict.keys())
    floors = {k: int(math.floor(max(0.0, float_vals_dict[k]))) for k in keys}
    rem = int(target_int_sum) - sum(floors.values())
    if rem > 0:
        fracs = sorted(keys, key=lambda k: float_vals_dict[k] - floors[k], reverse=True)
        for i in range(min(rem, len(fracs))):
            floors[fracs[i]] += 1
    elif rem < 0:
        fracs = sorted(keys, key=lambda k: float_vals_dict[k] - floors[k])
        for k in fracs:
            while rem < 0 and floors[k] > 0:
                floors[k] -= 1
                rem += 1
    return floors


def _build_stat_rows(alloc_dict, total_do, target_pct, tolerance_pct):
    rows = []
    for x in EXPEDITIONS:
        r = {"XPDC": x}
        gt = 0
        for p in PLANTS:
            val = int(alloc_dict[x][p])
            r[p] = val
            gt += val
        pct = (gt / total_do) if total_do > 0 else 0.0
        t_pct = float(target_pct[x])
        t_units = t_pct * total_do
        gap_u = gt - t_units
        diff_pct = (pct - t_pct) * 100.0
        r["GRAND_TOTAL"] = gt
        r["PCT"] = round(pct * 100.0, 2)
        r["TARGET_UNITS"] = round(t_units, 1)
        r["TARGET_PCT"] = round(t_pct * 100.0, 2)
        r["GAP_UNITS"] = round(gap_u, 1)
        r["DIFF_PCT"] = round(diff_pct, 2)
        r["WITHIN_TOL"] = abs(diff_pct) <= tolerance_pct
        rows.append(r)
    return rows


def solve_fm8_allocation(do_fm8, target_units, curr_alloc):
    """
    Kalkulasi Alokasi FM8 (Plant 8) sesuai batasan armada:
    1. Ekspedisi RJTM, LBK, WSS, PUNINAR (Khusus Trailer):
       - Minimal alokasi 54 unit dan maksimal 64 unit (1 rit trailer)
       - Jika lebih dari itu dikali 2: rentang 108 s.d. 128 unit (2 rit trailer)
       - Jika lebih dari itu dikali 3: rentang 162 s.d. 192 unit (3 rit trailer), dst.
       - Atau 0 unit jika tidak dialokasikan di FM8 pada hari itu.
    2. Ekspedisi TUNAS MUDA (Pilihan Engkel atau Trailer):
       - 44 unit (1 rit engkel), jika lebih dari itu dikali 2 (88 unit, 2 rit engkel), 132 unit, dst.
       - Atau 54 s.d. 64 unit (1 rit trailer), jika lebih dari itu dikali 2 (108 s.d. 128 unit, 2 rit trailer), dst.
       - Atau 0 unit jika tidak dialokasikan di FM8 pada hari itu.
    """
    if do_fm8 <= 0:
        return {x: 0 for x in EXPEDITIONS}

    if do_fm8 <= 44:
        alloc = {x: 0 for x in EXPEDITIONS}
        alloc["TUNAS MUDA"] = do_fm8
        return alloc

    rem_need = {x: max(0.0, target_units[x] - curr_alloc[x]) for x in EXPEDITIONS}
    others = ["RJTM", "LBK", "WSS", "PUNINAR"]

    tm_options = []
    # Engkel: 44, 88, 132...
    for m in [1, 2, 3]:
        v = m * 44
        if v <= do_fm8:
            tm_options.append(("engkel", m, v, v))
    # Trailer: 1 rit (54..64), 2 rit (108..128), etc.
    for k in [1, 2, 3]:
        min_v = k * 54
        max_v = k * 64
        if min_v <= do_fm8:
            tm_options.append(("trailer", k, min_v, min(do_fm8, max_v)))

    if do_fm8 >= 54:
        tm_options.append(("none", 0, 0, 0))

    best_alloc = None
    best_score = float("inf")
    max_k = 3 if do_fm8 > 400 else 2

    for tm_type, tm_trips, tm_min, tm_max in tm_options:
        for k_rj in range(0, max_k + 1):
            for k_lbk in range(0, max_k + 1):
                for k_wss in range(0, max_k + 1):
                    for k_pun in range(0, max_k + 1):
                        trips_other = {"RJTM": k_rj, "LBK": k_lbk, "WSS": k_wss, "PUNINAR": k_pun}
                        tot_k_other = sum(trips_other.values())

                        min_other = tot_k_other * 54
                        max_other = tot_k_other * 64

                        if tm_type == "engkel":
                            tm_u = tm_min
                            rem = do_fm8 - tm_u
                            if min_other <= rem <= max_other:
                                if tot_k_other > 0:
                                    base = {x: trips_other[x] * 54 for x in others}
                                    slack = rem - min_other
                                    alloc_other = dict(base)
                                    for x in sorted(others, key=lambda o: rem_need[o], reverse=True):
                                        if trips_other[x] > 0 and slack > 0:
                                            add = min(slack, trips_other[x] * 10)
                                            alloc_other[x] += add
                                            slack -= add
                                else:
                                    if rem != 0:
                                        continue
                                    alloc_other = {x: 0 for x in others}

                                alloc_total = dict(alloc_other)
                                alloc_total["TUNAS MUDA"] = tm_u

                                active_count = (1 if tm_trips > 0 else 0) + sum(1 for k in trips_other.values() if k > 0)
                                penalty = 0
                                if do_fm8 > 270:
                                    if active_count < 5:
                                        penalty += (5 - active_count) * 100000
                                elif do_fm8 > 100:
                                    if active_count < 2:
                                        penalty += 50000

                                ideal_ratio = do_fm8 / sum(rem_need.values()) if sum(rem_need.values()) > 0 else 0.2
                                var_score = sum((alloc_total[x] - rem_need[x] * ideal_ratio) ** 2 for x in EXPEDITIONS)

                                score = penalty + var_score
                                if score < best_score:
                                    best_score = score
                                    best_alloc = alloc_total

                        elif tm_type == "trailer":
                            tot_k = tm_trips + tot_k_other
                            if tot_k == 0:
                                continue
                            min_all = tot_k * 54
                            max_all = tot_k * 64
                            if min_all <= do_fm8 <= max_all:
                                all_trips = dict(trips_other)
                                all_trips["TUNAS MUDA"] = tm_trips
                                base = {x: all_trips[x] * 54 for x in EXPEDITIONS}
                                slack = do_fm8 - min_all
                                alloc_total = dict(base)
                                for x in sorted(EXPEDITIONS, key=lambda o: rem_need[o], reverse=True):
                                    if all_trips[x] > 0 and slack > 0:
                                        add = min(slack, all_trips[x] * 10)
                                        alloc_total[x] += add
                                        slack -= add

                                active_count = sum(1 for k in all_trips.values() if k > 0)
                                penalty = 0
                                if do_fm8 > 270:
                                    if active_count < 5:
                                        penalty += (5 - active_count) * 100000
                                elif do_fm8 > 100:
                                    if active_count < 2:
                                        penalty += 50000

                                ideal_ratio = do_fm8 / sum(rem_need.values()) if sum(rem_need.values()) > 0 else 0.2
                                var_score = sum((alloc_total[x] - rem_need[x] * ideal_ratio) ** 2 for x in EXPEDITIONS)

                                score = penalty + var_score
                                if score < best_score:
                                    best_score = score
                                    best_alloc = alloc_total

    if best_alloc is None:
        if do_fm8 <= 44:
            best_alloc = {x: 0 for x in EXPEDITIONS}
            best_alloc["TUNAS MUDA"] = do_fm8
        else:
            best_alloc = _round_preserving_sum({x: do_fm8 / 5.0 for x in EXPEDITIONS}, do_fm8)

    return best_alloc


def calculate_and_balance_daily_quota(
    do_dict,
    pct_matrix,
    rolling_winner="TUNAS MUDA",
    mode="operational_trailer",
    tolerance_pct=0.10,
):
    """
    Menghitung Kuota Harian (Unit) dari DO Harian dan Kuota Bulanan (%).
    """
    do_clean = {p: int(do_dict.get(p, 0)) for p in PLANTS}
    total_do = sum(do_clean.values())
    target_pct = {x: float(pct_matrix[x]["TARGET_PCT"]) for x in EXPEDITIONS}

    # 1. KALKULASI AWAL (RAW)
    raw_float = {
        x: {p: float(do_clean[p]) * float(pct_matrix[x][p]) for p in PLANTS}
        for x in EXPEDITIONS
    }
    raw_int = {x: {p: 0 for p in PLANTS} for x in EXPEDITIONS}
    for p in PLANTS:
        col_dict = {x: raw_float[x][p] for x in EXPEDITIONS}
        rounded_col = _round_preserving_sum(col_dict, do_clean[p])
        for x in EXPEDITIONS:
            raw_int[x][p] = rounded_col[x]

    raw_rows = _build_stat_rows(raw_int, total_do, target_pct, tolerance_pct)
    all_within_tolerance = all(r["WITHIN_TOL"] for r in raw_rows)

    # 2. MESIN AUTO-BALANCING
    bal_int = {x: {p: 0 for p in PLANTS} for x in EXPEDITIONS}

    if total_do == 0 or (all_within_tolerance and mode == "excel_formula"):
        for x in EXPEDITIONS:
            for p in PLANTS:
                bal_int[x][p] = raw_int[x][p]

    elif mode == "excel_formula":
        qty_calc = {x: target_pct[x] * total_do for x in EXPEDITIONS}
        j_sum = {x: sum(raw_float[x][p] for p in PLANTS) for x in EXPEDITIONS}
        k_pct = {x: j_sum[x] / total_do for x in EXPEDITIONS}
        n_gap = {x: j_sum[x] - qty_calc[x] for x in EXPEDITIONS}

        o_adj_fm6 = {}
        p_adj_fm8 = {}
        new_fm6 = {}
        new_fm8 = {}
        for x in ["TUNAS MUDA", "RJTM"]:
            e_val = raw_float[x]["FM6"]
            h_val = raw_float[x]["FM8"]
            cond = (k_pct[x] > target_pct[x]) and (e_val < 20.0)
            o_adj_fm6[x] = e_val if cond else 0.0
            p_adj_fm8[x] = (n_gap[x] - e_val) if cond else n_gap[x]
            new_fm6[x] = (
                (e_val - p_adj_fm8[x] - o_adj_fm6[x])
                if (p_adj_fm8[x] < 0)
                else (e_val - o_adj_fm6[x])
            )
            new_fm8[x] = (h_val - p_adj_fm8[x]) if (p_adj_fm8[x] > 0) else h_val

        for x in ["LBK", "WSS", "PUNINAR"]:
            o_adj_fm6[x] = 0.0 if n_gap[x] < 0 else n_gap[x]

        o33_sum = sum(o_adj_fm6.values())
        neg_gaps = sorted(
            [(x, abs(n_gap[x]) if n_gap[x] < 0 else 0.0) for x in ["LBK", "WSS", "PUNINAR"]],
            key=lambda item: item[1],
        )
        rem_o33 = o33_sum
        q_alloc = {}
        for x, abs_g in neg_gaps:
            alloc = min(abs_g, max(0.0, rem_o33))
            q_alloc[x] = alloc
            rem_o33 -= alloc

        for x in ["LBK", "WSS", "PUNINAR"]:
            e_val = raw_float[x]["FM6"]
            new_fm6[x] = max(0.0, e_val - o_adj_fm6[x] + q_alloc[x])
            new_fm8[x] = abs(n_gap[x] + q_alloc[x]) if n_gap[x] < 0 else 0.0

        bal_float = {
            x: {p: raw_float[x][p] for p in PLANTS} for x in EXPEDITIONS
        }
        for x in EXPEDITIONS:
            bal_float[x]["FM6"] = max(0.0, new_fm6[x])
            bal_float[x]["FM8"] = max(0.0, new_fm8[x])

        for p in PLANTS:
            col_sum = sum(bal_float[x][p] for x in EXPEDITIONS)
            if col_sum > 0 and do_clean[p] > 0:
                scaled = {
                    x: bal_float[x][p] / col_sum * do_clean[p] for x in EXPEDITIONS
                }
            else:
                scaled = {x: 0.0 for x in EXPEDITIONS}
            rounded_col = _round_preserving_sum(scaled, do_clean[p])
            for x in EXPEDITIONS:
                bal_int[x][p] = rounded_col[x]

    else:
        # MODE OPERATIONAL TRAILER & ROLLING RULES
        target_units_int = _round_preserving_sum(
            {x: target_pct[x] * total_do for x in EXPEDITIONS}, total_do
        )

        winner = (
            rolling_winner
            if rolling_winner in ["TUNAS MUDA", "RJTM"]
            else "TUNAS MUDA"
        )
        bal_int[winner]["FM1"] = do_clean["FM1"]
        bal_int[winner]["FM2"] = do_clean["FM2"]

        if do_clean["FM9"] > 0:
            fm9_weights = {
                "TUNAS MUDA": pct_matrix["TUNAS MUDA"]["FM9"] or 0.5017,
                "RJTM": pct_matrix["RJTM"]["FM9"] or 0.4983,
            }
            w_sum = sum(fm9_weights.values())
            fm9_alloc = _round_preserving_sum(
                {x: (fm9_weights.get(x, 0.0) / w_sum) * do_clean["FM9"] for x in EXPEDITIONS},
                do_clean["FM9"],
            )
            for x in EXPEDITIONS:
                bal_int[x]["FM9"] = fm9_alloc[x]

        if do_clean["FM7"] > 0:
            fm7_weights = {
                "TUNAS MUDA": pct_matrix["TUNAS MUDA"]["FM7"] or 0.578,
                "RJTM": pct_matrix["RJTM"]["FM7"] or 0.422,
            }
            w_sum = sum(fm7_weights.values())
            fm7_alloc = _round_preserving_sum(
                {x: (fm7_weights.get(x, 0.0) / w_sum) * do_clean["FM7"] for x in EXPEDITIONS},
                do_clean["FM7"],
            )
            for x in EXPEDITIONS:
                bal_int[x]["FM7"] = fm7_alloc[x]

        if do_clean["FM6A"] > 0:
            g_b = ["LBK", "WSS", "PUNINAR"]
            w_6a = {x: pct_matrix[x]["FM6A"] or (1.0 / 3.0) for x in g_b}
            w_sum = sum(w_6a.values())
            fm6a_alloc = _round_preserving_sum(
                {x: (w_6a.get(x, 0.0) / w_sum) * do_clean["FM6A"] for x in EXPEDITIONS},
                do_clean["FM6A"],
            )
            for x in EXPEDITIONS:
                bal_int[x]["FM6A"] = fm6a_alloc[x]

        # ATURAN KHUSUS FM8:
        # RJTM, LBK, WSS, PUNINAR: Minimal 54 unit dan maksimal 64 unit (1 rit trailer); jika lebih dari itu dikali 2 (108 s.d. 128 unit, 2 rit trailer), dst.
        # TUNAS MUDA: 44 unit (1 rit engkel), jika lebih dari itu dikali 2 (88 unit); atau 54 sampai 64 unit (1 rit trailer), jika lebih dari itu dikali 2 (108-128 unit).
        do_fm8 = do_clean["FM8"]
        if do_fm8 > 0:
            curr_temp = {
                x: sum(bal_int[x][p] for p in ["FM1", "FM2", "FM6A", "FM7", "FM9"])
                for x in EXPEDITIONS
            }
            fm8_alloc = solve_fm8_allocation(do_fm8, target_units_int, curr_temp)
            for x in EXPEDITIONS:
                bal_int[x]["FM8"] = fm8_alloc[x]

        # ATURAN KHUSUS FM6:
        # Jika Plant 6 perlu balancing akhir dari Ekspedisi LBK, WSS, PUNINAR:
        # Tambahkan ekspedisi TUNAS MUDA dan RJTM dengan kapasitas minimal 54 per truk
        # maksimal 64 per truk dikali alokasi DO harian.
        do_fm6 = do_clean["FM6"]
        current_alloc = {
            x: sum(bal_int[x][p] for p in ["FM1", "FM2", "FM6A", "FM7", "FM8", "FM9"])
            for x in EXPEDITIONS
        }
        rem_need = {
            x: max(0, target_units_int[x] - current_alloc[x])
            for x in EXPEDITIONS
        }
        tot_rem = sum(rem_need.values())

        if do_fm6 > 0:
            if tot_rem > 0:
                fm6_alloc = _round_preserving_sum(
                    {x: (rem_need[x] / tot_rem) * do_fm6 for x in EXPEDITIONS},
                    do_fm6,
                )
            else:
                fm6_alloc = _round_preserving_sum(
                    {x: target_pct[x] * do_fm6 for x in EXPEDITIONS},
                    do_fm6,
                )
            for x in EXPEDITIONS:
                bal_int[x]["FM6"] = fm6_alloc[x]

    balanced_rows = _build_stat_rows(bal_int, total_do, target_pct, tolerance_pct)

    trailer_checks = {}
    for x in EXPEDITIONS:
        trailer_checks[x] = {}
        for p in PLANTS:
            u = int(bal_int[x][p])
            if p == "FM8":
                if u == 0:
                    trailer_checks[x][p] = {"trailers": 0, "valid": True, "type": "none", "note": "-"}
                elif x == "TUNAS MUDA":
                    # TUNAS MUDA: 44 unit (1 engkel), 88 unit (2 engkel), atau trailer (54-64, 108-128, dst.)
                    if u % 44 == 0 and u > 0:
                        m = u // 44
                        trailer_checks[x][p] = {
                            "trailers": m,
                            "valid": True,
                            "type": "engkel",
                            "note": f"{m} Engkel ({u} u)",
                        }
                    else:
                        min_k = math.ceil(u / 64.0)
                        max_k = math.floor(u / 54.0)
                        if min_k <= max_k and min_k >= 1:
                            trailer_checks[x][p] = {
                                "trailers": min_k,
                                "valid": True,
                                "type": "trailer",
                                "note": f"{min_k} Truk (~{u/min_k:.0f} u)",
                            }
                        else:
                            est_k = max(min_k, round(u / 59.0))
                            is_ok = (54 <= u / est_k <= 64) if est_k > 0 else False
                            trailer_checks[x][p] = {
                                "trailers": est_k,
                                "valid": is_ok,
                                "type": "trailer",
                                "note": f"{est_k} Truk (~{u/est_k:.0f} u)",
                            }
                else:
                    # RJTM, LBK, WSS, PUNINAR: Trailer min 54 s.d. 64 u/truk, kelipatan 2 rit (108-128), dst.
                    min_k = math.ceil(u / 64.0)
                    max_k = math.floor(u / 54.0)
                    if min_k <= max_k and min_k >= 1:
                        trailer_checks[x][p] = {
                            "trailers": min_k,
                            "valid": True,
                            "type": "trailer",
                            "note": f"{min_k} Truk (~{u/min_k:.0f} u)",
                        }
                    else:
                        est_k = max(min_k, round(u / 59.0))
                        is_ok = (54 <= u / est_k <= 64) if est_k > 0 else False
                        trailer_checks[x][p] = {
                            "trailers": est_k,
                            "valid": is_ok,
                            "type": "trailer",
                            "note": f"{est_k} Truk (~{u/est_k:.0f} u)",
                        }
            else:
                if u == 0:
                    trailer_checks[x][p] = {"trailers": 0, "valid": True, "type": "none", "note": "-"}
                else:
                    min_k = math.ceil(u / 64.0)
                    max_k = math.floor(u / 54.0)
                    if min_k <= max_k and min_k >= 1:
                        trailer_checks[x][p] = {
                            "trailers": min_k,
                            "valid": True,
                            "type": "trailer",
                            "note": f"{min_k} Truk (~{u/min_k:.0f} u)",
                        }
                    else:
                        est_k = max(min_k, round(u / 59.0))
                        is_ok = (54 <= u / est_k <= 64) if est_k > 0 else False
                        trailer_checks[x][p] = {
                            "trailers": est_k,
                            "valid": is_ok,
                            "type": "trailer",
                            "note": f"{est_k} Truk (~{u/est_k:.0f} u)",
                        }

    return {
        "do_per_plant": do_clean,
        "total_do": total_do,
        "rolling_winner": rolling_winner,
        "mode": mode,
        "tolerance_pct": tolerance_pct,
        "all_within_tolerance_raw": all_within_tolerance,
        "raw_rows": raw_rows,
        "balanced_rows": balanced_rows,
        "trailer_checks": trailer_checks,
    }


def calculate_summary_and_performance(
    sl_records,
    do_records,
    alloc_records,
    monthly_quota_pct,
    report_date_str="2026-10-31",
    unfilled_h_minus_1=0,
    planned_monthly_target=None,
):
    """Menghitung Output 3.A (DISTRIBUTION SUMMARY) & 3.B (ANALISA PERFORMANCE)."""
    rep_str = parse_date_str(report_date_str) or "2026-10-31"
    rep_year = int(rep_str[:4])
    rep_month = int(rep_str[5:7])

    sl_month = [
        r for r in sl_records
        if int(r["YEAR"]) == rep_year and int(r["MONTH"]) == rep_month
    ]
    sl_upto = [r for r in sl_month if r["TANGGAL_STR"] <= rep_str]
    sl_daily = [r for r in sl_month if r["TANGGAL_STR"] == rep_str]

    do_month = [
        r for r in do_records
        if int(r["TAHUN"]) == rep_year and int(r["BULAN"]) == rep_month
    ]
    do_upto = [r for r in do_month if r["TANGGAL_STR"] <= rep_str]
    do_daily_total = sum(int(r["QTY"]) for r in do_month if r["TANGGAL_STR"] == rep_str)

    alloc_month = [
        r for r in alloc_records
        if int(r["TAHUN"]) == rep_year and int(r["BULAN"]) == rep_month
    ]
    alloc_upto = [r for r in alloc_month if r["TANGGAL"] <= rep_str]
    alloc_daily = [r for r in alloc_month if r["TANGGAL"] == rep_str]

    total_target_do_month = sum(int(r["QTY"]) for r in do_month)
    total_target_do_upto = sum(int(r["QTY"]) for r in do_upto)
    sisa_do_month = max(0, total_target_do_month - total_target_do_upto)

    dist_rows = []
    perf_rows = []

    tot_daily_tgt = 0
    tot_daily_act = 0
    tot_mtd_tgt_full = 0
    tot_mtd_act = 0
    tot_upto_tgt = 0

    has_plan = bool(planned_monthly_target and isinstance(planned_monthly_target, dict) and any(int(v) > 0 for v in planned_monthly_target.values()))

    for x in EXPEDITIONS:
        q_pct = float(monthly_quota_pct.get(x, 0.0))
        d_tgt = sum(int(r["GRAND_TOTAL"]) for r in alloc_daily if r["XPDC"] == x)
        d_act = sum(int(r["QTY"]) for r in sl_daily if r["XPDC"] == x)

        if d_act == 0 and d_tgt == 0:
            d_ach = 0.0
        elif d_act > 0 and d_tgt == 0:
            d_ach = 1.0
        else:
            d_ach = d_act / d_tgt

        upto_tgt = sum(int(r["GRAND_TOTAL"]) for r in alloc_upto if r["XPDC"] == x)
        if upto_tgt == 0 and total_target_do_upto > 0:
            upto_tgt = int(round(total_target_do_upto * q_pct))

        sisa_eksp = int(round(sisa_do_month * q_pct))
        
        # Jika ada Target Perencanaan Bulanan yang disimpan, tampilkan di kolom TARGET BULAN
        if has_plan and x in planned_monthly_target and int(planned_monthly_target[x]) > 0:
            mtd_full_tgt = int(planned_monthly_target[x])
        else:
            mtd_full_tgt = upto_tgt + sisa_eksp

        mtd_act = sum(int(r["QTY"]) for r in sl_upto if r["XPDC"] == x)
        mtd_ach = (mtd_act / mtd_full_tgt) if mtd_full_tgt > 0 else 0.0

        tot_daily_tgt += d_tgt
        tot_daily_act += d_act
        tot_mtd_tgt_full += mtd_full_tgt
        tot_mtd_act += mtd_act
        tot_upto_tgt += upto_tgt

        dist_rows.append({
            "XPDC": x,
            "DAILY_TARGET": d_tgt,
            "DAILY_ACTUAL": d_act,
            "DAILY_ACH_PCT": round(d_ach * 100, 2),
            "MTD_TARGET": mtd_full_tgt,
            "MTD_ACTUAL": mtd_act,
            "MTD_ACH_PCT": round(mtd_ach * 100, 2),
            "QUOTA_PCT": round(q_pct * 100, 2),
            "TARGET_UPTO": upto_tgt,
            "SISA_TARGET": max(0, mtd_full_tgt - upto_tgt),
        })

        gap_upto = mtd_act - upto_tgt
        ach_upto = (mtd_act / upto_tgt) if upto_tgt > 0 else 0.0
        perf_rows.append({
            "XPDC": x,
            "TARGET_UPTO": upto_tgt,
            "ACTUAL_UPTO": mtd_act,
            "GAP": gap_upto,
            "ACH_PCT": round(ach_upto * 100, 2),
        })

    for r in dist_rows:
        r["KONTRIBUSI_PCT"] = (
            round((r["MTD_ACTUAL"] / tot_mtd_act * 100), 2) if tot_mtd_act > 0 else 0.0
        )
    for r in perf_rows:
        r["SHARE_TARGET_PCT"] = (
            round((r["TARGET_UPTO"] / tot_upto_tgt * 100), 2) if tot_upto_tgt > 0 else 0.0
        )

    total_do_box = int(unfilled_h_minus_1) + do_daily_total
    unfilled_h = max(0, total_do_box - tot_daily_act)

    return {
        "report_date": rep_str,
        "month_name": MONTH_NAMES_ID.get(rep_month, str(rep_month)),
        "year": rep_year,
        "has_planned_target": has_plan,
        "unfilled_box": {
            "unfilled_h_minus_1": int(unfilled_h_minus_1),
            "do_ahm": do_daily_total,
            "total_do": total_do_box,
            "act_dist": tot_daily_act,
            "unfilled_h": unfilled_h,
        },
        "distribution_summary": dist_rows,
        "distribution_totals": {
            "DAILY_TARGET": tot_daily_tgt,
            "DAILY_ACTUAL": tot_daily_act,
            "DAILY_ACH_PCT": round((tot_daily_act / tot_daily_tgt * 100), 2)
            if tot_daily_tgt > 0
            else (100.0 if tot_daily_act > 0 else 0.0),
            "MTD_TARGET": tot_mtd_tgt_full,
            "MTD_ACTUAL": tot_mtd_act,
            "MTD_ACH_PCT": round((tot_mtd_act / tot_mtd_tgt_full * 100), 2)
            if tot_mtd_tgt_full > 0
            else 0.0,
            "TARGET_UPTO": tot_upto_tgt,
            "SISA_TARGET": max(0, tot_mtd_tgt_full - tot_upto_tgt),
        },
        "performance_analysis": perf_rows,
        "performance_totals": {
            "TARGET_UPTO": tot_upto_tgt,
            "ACTUAL_UPTO": tot_mtd_act,
            "GAP": tot_mtd_act - tot_upto_tgt,
            "ACH_PCT": round((tot_mtd_act / tot_upto_tgt * 100), 2)
            if tot_upto_tgt > 0
            else 0.0,
        },
    }
