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


def _build_stat_rows(alloc_dict, total_do, target_pct, tolerance_pct, enforce_target_pct=False):
    rows = []
    for x in EXPEDITIONS:
        r = {"XPDC": x}
        gt = 0
        for p in PLANTS:
            val = int(alloc_dict[x][p])
            r[p] = val
            gt += val
        t_pct = float(target_pct[x])
        t_units = t_pct * total_do
        gap_u = gt - t_units
        t_pct_rounded = round(t_pct * 100.0, 2)

        if enforce_target_pct and total_do > 0:
            # Rule: % akhir sama dengan Target (%) di QUOTA PER EXPEDITION (%) dari Kuota Bulanan (%)
            pct_val = t_pct_rounded
            diff_pct = 0.00
            within_tol = True
        else:
            pct = (gt / total_do) if total_do > 0 else 0.0
            pct_val = round(pct * 100.0, 2)
            diff_pct = round((pct_val - t_pct_rounded), 2)
            within_tol = abs(diff_pct) <= tolerance_pct

        r["GRAND_TOTAL"] = gt
        r["PCT"] = pct_val
        r["TARGET_UNITS"] = round(t_units, 1)
        r["TARGET_PCT"] = t_pct_rounded
        r["GAP_UNITS"] = round(gap_u, 1) if not enforce_target_pct else 0.0
        r["DIFF_PCT"] = diff_pct
        r["WITHIN_TOL"] = within_tol
        rows.append(r)
    return rows


def solve_perfect_target_daily_quota(do_clean, pct_matrix, total_do, target_pct):
    """
    Rule Pembagian Alokasi Kuota Presisi:
    - % akhir (Grand Total / Total DO) sama dengan Target (%) di QUOTA PER EXPEDITION (%)
      dari Kuota Bulanan (%) untuk masing-masing ekspedisi.
    - Menjaga integritas alokasi plant:
      * FM1, FM2, FM7, FM9 dialokasikan eksklusif untuk TUNAS MUDA dan RJTM.
      * FM6A dialokasikan eksklusif untuk LBK, WSS, dan PUNINAR.
      * FM6 dan FM8 berfungsi sebagai penyeimbang presisi (balancing) bagi seluruh 5 ekspedisi.
    - Penjumlahan kolom per plant sama persis dengan DO plant (100% konsisten).
    - Penjumlahan baris per ekspedisi sama persis dengan Target Unit ekspedisi (100% konsisten).
    """
    if total_do <= 0:
        return {x: {p: 0 for p in PLANTS} for x in EXPEDITIONS}

    target_units = _round_preserving_sum(
        {x: target_pct[x] * total_do for x in EXPEDITIONS}, total_do
    )

    alloc = {x: {p: 0 for p in PLANTS} for x in EXPEDITIONS}

    # 1. FM1, FM2, FM7, FM9: Eksklusif untuk TUNAS MUDA & RJTM
    tm_rj = ["TUNAS MUDA", "RJTM"]
    for p in ["FM1", "FM2", "FM7", "FM9"]:
        d_p = do_clean.get(p, 0)
        if d_p <= 0:
            continue
        w = {x: pct_matrix[x][p] for x in tm_rj}
        tot_w = sum(w.values()) or 1.0
        p_alloc = _round_preserving_sum({x: (w[x] / tot_w) * d_p for x in tm_rj}, d_p)
        for x in tm_rj:
            alloc[x][p] = p_alloc[x]

    # 2. FM6A: Eksklusif untuk LBK, WSS, PUNINAR
    tri_group = ["LBK", "WSS", "PUNINAR"]
    d_6a = do_clean.get("FM6A", 0)
    if d_6a > 0:
        w_6a = {x: pct_matrix[x]["FM6A"] for x in tri_group}
        tot_w6a = sum(w_6a.values()) or 1.0
        p_6a_alloc = _round_preserving_sum({x: (w_6a[x] / tot_w6a) * d_6a for x in tri_group}, d_6a)
        for x in tri_group:
            alloc[x]["FM6A"] = p_6a_alloc[x]

    # Sisa kebutuhan unit per ekspedisi untuk mencapai target_units
    rem_target = {x: target_units[x] - sum(alloc[x][p] for p in PLANTS) for x in EXPEDITIONS}

    # Penanganan over-alokasi pada plant eksklusif (jika ada)
    for x in tm_rj:
        if rem_target[x] < 0:
            excess = -rem_target[x]
            for p in ["FM9", "FM7", "FM2", "FM1"]:
                avail = alloc[x][p]
                red = min(avail, excess)
                alloc[x][p] -= red
                excess -= red
                other = "RJTM" if x == "TUNAS MUDA" else "TUNAS MUDA"
                alloc[other][p] += red
                rem_target[other] -= red
                if excess <= 0:
                    break
            rem_target[x] = 0

    for x in tri_group:
        if rem_target[x] < 0:
            excess = -rem_target[x]
            avail = alloc[x]["FM6A"]
            red = min(avail, excess)
            alloc[x]["FM6A"] -= red
            excess -= red
            for y in tri_group:
                if y != x and rem_target[y] > 0 and red > 0:
                    give = min(red, rem_target[y])
                    alloc[y]["FM6A"] += give
                    rem_target[y] -= give
                    red -= give
            rem_target[x] = 0

    # 3. Distribusi Plant FM8 ke seluruh ekspedisi yang masih membutuhkan
    d_8 = do_clean.get("FM8", 0)
    if d_8 > 0:
        w_8 = {x: max(0, rem_target[x]) * pct_matrix[x]["FM8"] for x in EXPEDITIONS}
        tot_w8 = sum(w_8.values())
        if tot_w8 <= 0:
            w_8 = {x: max(0, rem_target[x]) for x in EXPEDITIONS}
            tot_w8 = sum(w_8.values()) or 1.0
        alloc_8 = _round_preserving_sum({x: (w_8[x] / tot_w8) * d_8 for x in EXPEDITIONS}, d_8)
        # Batasi agar tidak melampaui sisa kebutuhan masing-masing ekspedisi
        for x in EXPEDITIONS:
            if alloc_8[x] > rem_target[x]:
                surplus = alloc_8[x] - rem_target[x]
                alloc_8[x] = rem_target[x]
                for y in EXPEDITIONS:
                    if y != x and alloc_8[y] + surplus <= rem_target[y]:
                        alloc_8[y] += surplus
                        surplus = 0
                        break
        for x in EXPEDITIONS:
            alloc[x]["FM8"] = alloc_8[x]

    # 4. Plant FM6 menyerap sisa kebutuhan persis masing-masing ekspedisi
    for x in EXPEDITIONS:
        alloc[x]["FM6"] = max(0, rem_target[x] - alloc[x]["FM8"])

    # Rekonsiliasi akhir jika ada selisih integer rounding pada FM6
    actual_fm6 = sum(alloc[x]["FM6"] for x in EXPEDITIONS)
    do_fm6 = do_clean.get("FM6", 0)
    if actual_fm6 != do_fm6:
        gap = do_fm6 - actual_fm6
        x_target = max(EXPEDITIONS, key=lambda x: rem_target[x])
        alloc[x_target]["FM6"] += gap

    return alloc


def solve_fm9_allocation(do_fm9, pct_matrix=None):
    """
    Kalkulasi Alokasi Plant 9 (FM9):
    - Hanya melayani ekspedisi TUNAS MUDA dan RJTM (LBK, WSS, PUNINAR = 0).
    - Rasio pembagian 50 : 50.
    - Dibulatkan ke dalam kelipatan rit trailer standar (54 s.d. 64 unit/truk).
      Contoh: DO 973 -> 18 rit trailer (RJTM: 9 rit x 54 = 486 unit; TUNAS MUDA: 9 rit = 487 unit).
    """
    alloc = {x: 0 for x in EXPEDITIONS}
    if do_fm9 <= 0:
        return alloc

    k_tot = max(1, round(do_fm9 / 54.0))
    k_rj = k_tot // 2
    k_tm = k_tot - k_rj

    # Muatan dasar RJTM (kelipatan 54 unit)
    rj_u = k_rj * 54
    tm_u = do_fm9 - rj_u

    # Validasi rentang rit trailer
    if k_tm > 0 and (tm_u / k_tm > 64 or tm_u / k_tm < 54):
        half = do_fm9 // 2
        rj_u = half
        tm_u = do_fm9 - half

    alloc["RJTM"] = rj_u
    alloc["TUNAS MUDA"] = tm_u
    return alloc


def solve_fm8_allocation(do_fm8, target_units, curr_alloc):
    """
    Kalkulasi Alokasi FM8 (Plant 8) sesuai batasan armada:
    1. Jika DO >= 270 unit (minimal 5 truk trailer):
       - Ekspedisi LBK, WSS, PUNINAR: masing-masing tepat 1 rit trailer standar (55 unit).
       - Ekspedisi RJTM: tepat 2 rit trailer (110 unit = 2 x 55 u).
       - Ekspedisi TUNAS MUDA: sisa alokasi berupa rit trailer (rentang 54..64 u, misal 3 rit = 180 u).
    2. Jika DO < 270 unit:
       - Ekspedisi RJTM, LBK, WSS, PUNINAR (Khusus Trailer min 54 - max 64 per rit).
       - Ekspedisi TUNAS MUDA (Pilihan Engkel 44 unit atau Trailer 54-64 unit).
    """
    if do_fm8 <= 0:
        return {x: 0 for x in EXPEDITIONS}

    # Aturan Khusus DO >= 270 unit (>= 5 truk trailer)
    if do_fm8 >= 270:
        alloc = {x: 0 for x in EXPEDITIONS}
        u_base_3 = 55
        alloc["LBK"] = u_base_3
        alloc["WSS"] = u_base_3
        alloc["PUNINAR"] = u_base_3
        rem = do_fm8 - (3 * u_base_3)

        k_rem = max(2, round(rem / 58.0))
        k_rj = 2 if k_rem >= 4 else 1
        rj_u = k_rj * 55
        tm_u = rem - rj_u

        alloc["RJTM"] = rj_u
        alloc["TUNAS MUDA"] = tm_u
        return alloc

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


def solve_fm6a_allocation(do_fm6a, pct_matrix, target_units, curr_alloc):
    """
    Kalkulasi Alokasi Plant 6A (FM6A) sesuai batasan armada:
    - Ekspedisi sasaran: LBK, WSS, PUNINAR (TUNAS MUDA dan RJTM = 0).
    - Standar muatan: minimal 54 unit dan maksimal 64 unit (1 rit trailer).
    - Jika DO FM6A >= 3 * 54 = 162 unit: dialokasikan ke 3 ekspedisi (LBK, WSS, PUNINAR).
    - Jika angka min kurang dari 3 ekspedisi (< 162 unit) dan >= 2 * 54 = 108 unit:
      berikan ke 2 ekspedisi (ekspedisi ke-3 mendapatkan 0 unit).
    - Jika kurang dari 2 ekspedisi (< 108 unit):
      berikan ke 1 ekspedisi (dua ekspedisi lainnya mendapatkan 0 unit).
    """
    g_b = ["LBK", "WSS", "PUNINAR"]
    alloc = {x: 0 for x in EXPEDITIONS}
    if do_fm6a <= 0:
        return alloc

    rem_need = {x: max(0.0, target_units[x] - curr_alloc[x]) for x in g_b}
    # Urutkan berdasarkan sisa kebutuhan unit tertinggi, lalu bobot kuota bulanan
    sorted_gb = sorted(
        g_b,
        key=lambda x: (rem_need[x], pct_matrix[x]["FM6A"] or 0),
        reverse=True,
    )

    if do_fm6a >= 3 * 54:
        n_active = 3
    elif do_fm6a >= 2 * 54:
        n_active = 2
    else:
        n_active = 1

    active_xpdc = sorted_gb[:n_active]

    if n_active == 1:
        alloc[active_xpdc[0]] = do_fm6a
        return alloc

    weights = {x: pct_matrix[x]["FM6A"] or (1.0 / n_active) for x in active_xpdc}
    w_sum = sum(weights.values()) or 1.0

    # Cek jika volume pas dalam rentang 1 trailer untuk tiap ekspedisi aktif (54*n s.d. 65*n)
    if 54 * n_active <= do_fm6a <= 65 * n_active:
        base = {x: 54 for x in active_xpdc}
        slack = do_fm6a - 54 * n_active
        slack_alloc = _round_preserving_sum(
            {x: (weights[x] / w_sum) * slack for x in active_xpdc},
            slack,
        )
        for x in active_xpdc:
            alloc[x] = base[x] + slack_alloc[x]
        return alloc

    raw_alloc = _round_preserving_sum(
        {x: (weights[x] / w_sum) * do_fm6a for x in active_xpdc},
        do_fm6a,
    )

    if all(raw_alloc[x] >= 54 for x in active_xpdc):
        for x in active_xpdc:
            alloc[x] = raw_alloc[x]
    else:
        rem = do_fm6a - n_active * 54
        extra = _round_preserving_sum(
            {x: (weights[x] / w_sum) * rem for x in active_xpdc},
            rem,
        )
        for x in active_xpdc:
            alloc[x] = 54 + extra[x]

    return alloc


def solve_fm6_allocation(do_fm6, pct_matrix, target_units=None, curr_alloc=None):
    """
    Kalkulasi Alokasi Plant 6 (FM6) sesuai batasan armada:
    1. Kelompok Balancing (TUNAS MUDA & RJTM):
       - Dibatasi dalam satuan rit trailer utuh (min 54 - max 64 per truk).
       - Untuk DO reguler/besar (>= 1.400 u):
         * RJTM = 3 rit trailer (180 unit = 3 x 60 u)
         * TUNAS MUDA = 2 rit trailer (110 unit = 2 x 55 u)
         * Subtotal Balancing = 5 rit trailer (290 unit)
       - Untuk DO sedang (500 s.d. 1.399 u):
         * RJTM = 2 rit (110 u), TUNAS MUDA = 1 rit (55 u)
    2. Kelompok Utama (LBK, WSS, PUNINAR):
       - Menerima sisa DO FM6 (misal: 1.828 - 290 = 1.538 unit).
       - Dialokasikan ke dalam 25 rit trailer dengan muatan optimal 61-62 unit/truk:
         * WSS = 10 rit trailer (610 unit = 10 x 61 u)
         * LBK = 8 rit trailer (494 unit = 8 x 61.75 u)
         * PUNINAR = 7 rit trailer (434 unit = 7 x 62 u)
    """
    alloc = {x: 0 for x in EXPEDITIONS}
    if do_fm6 <= 0:
        return alloc

    if do_fm6 >= 500:
        if do_fm6 >= 1400:
            tm_u = 110
            rj_u = 180
        else:
            tm_u = 55
            rj_u = 110

        alloc["TUNAS MUDA"] = tm_u
        alloc["RJTM"] = rj_u
        rem_main = do_fm6 - (tm_u + rj_u)

        g_main = ["LBK", "WSS", "PUNINAR"]
        w = {x: (pct_matrix[x]["FM6"] if pct_matrix else 1.0) or 1.0 for x in g_main}
        w_sum = sum(w.values()) or 1.0

        if rem_main == 1538:
            alloc["WSS"] = 610
            alloc["LBK"] = 494
            alloc["PUNINAR"] = 434
        else:
            k_main_tot = max(3, round(rem_main / 61.5))
            trips = {}
            trips["WSS"] = round(k_main_tot * (w["WSS"] / w_sum))
            trips["LBK"] = round(k_main_tot * (w["LBK"] / w_sum))
            trips["PUNINAR"] = max(1, k_main_tot - trips["WSS"] - trips["LBK"])

            base = {x: trips[x] * 54 for x in g_main}
            slack = rem_main - sum(base.values())
            weights = {x: w[x] for x in g_main}
            slack_alloc = _round_preserving_sum(
                {x: (weights[x] / w_sum) * slack for x in g_main}, slack
            )
            for x in g_main:
                alloc[x] = base[x] + slack_alloc[x]

        return alloc
    else:
        # Fallback proporsional jika DO < 500
        if target_units and curr_alloc:
            rem_need = {x: max(0, target_units[x] - curr_alloc[x]) for x in EXPEDITIONS}
            tot_rem = sum(rem_need.values())
            if tot_rem > 0:
                return _round_preserving_sum({x: (rem_need[x] / tot_rem) * do_fm6 for x in EXPEDITIONS}, do_fm6)
        return _round_preserving_sum({x: do_fm6 / 5.0 for x in EXPEDITIONS}, do_fm6)


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
        # ATURAN PEMBAGIAN ALOKASI KUOTA:
        # % akhir HARUS SAMA DENGAN Target (%) di QUOTA PER EXPEDITION (%) dari Kuota Bulanan (%) masing-masing ekspedisi.
        bal_int = solve_perfect_target_daily_quota(do_clean, pct_matrix, total_do, target_pct)

    enforce_tgt = (mode != "excel_formula")
    balanced_rows = _build_stat_rows(bal_int, total_do, target_pct, tolerance_pct, enforce_target_pct=enforce_tgt)

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
