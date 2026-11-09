"""
app.py - Aplikasi Web Python (Flask) Project Logistic DDC (Pure Python - Tanpa Pandas/C++)
==========================================================================================
Kompatibel penuh dengan Python 3.11, 3.12, 3.13, 3.14, dan 3.15+ di Windows, Mac, dan Linux.
"""

import os
from flask import Flask, render_template, request, jsonify, send_file
from db_excel import ExcelDatabase, PLANTS, EXPEDITIONS, parse_date_str, format_date_d_mmm_yy
from engine_ddc import (
    calculate_monthly_quota,
    calculate_and_balance_daily_quota,
    calculate_summary_and_performance,
)

app = Flask(__name__)
EXCEL_PATH = os.environ.get("DDC_EXCEL_PATH", "DDC.xlsx")
db = ExcelDatabase(EXCEL_PATH)


def get_calendar_info():
    """Mengembalikan informasi penanggalan kalender sistem (WIB UTC+7) dan aturan operasional DO & SL."""
    from datetime import datetime, timezone, timedelta
    WIB = timezone(timedelta(hours=7))
    now_wib = datetime.now(WIB)
    
    h_date = now_wib.date()
    h_minus_1 = h_date - timedelta(days=1)
    h_plus_1 = h_date + timedelta(days=1)

    MONTH_NAMES_ID_FULL = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
        5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
        9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }
    DAYS_ID = {
        0: "Senin", 1: "Selasa", 2: "Rabu", 3: "Kamis",
        4: "Jumat", 5: "Sabtu", 6: "Minggu"
    }
    
    h_str = h_date.strftime("%Y-%m-%d")
    h_m1_str = h_minus_1.strftime("%Y-%m-%d")
    h_p1_str = h_plus_1.strftime("%Y-%m-%d")

    return {
        "status": "ok",
        "calendar_date": h_str,
        "calendar_display": f"{DAYS_ID.get(now_wib.weekday(), '')}, {now_wib.day:02d} {MONTH_NAMES_ID_FULL.get(now_wib.month, '')} {now_wib.year}",
        "calendar_date_formatted": format_date_d_mmm_yy(h_str),
        "do_date_h_minus_1": h_m1_str,
        "do_date_formatted": format_date_d_mmm_yy(h_m1_str),
        "sl_date_h_plus_1": h_p1_str,
        "sl_date_formatted": format_date_d_mmm_yy(h_p1_str),
        "rules": [
            "DO akan di rilis H-1 dari tanggal kalender",
            "SL akan di upload H+1 dari tanggal kalender"
        ]
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/init", methods=["GET"])
def api_init():
    sheets = db.get_sheet_names()
    do_records = db.get_do_data()
    do_dates = sorted(list({r["TANGGAL_STR"] for r in do_records}))
    cal_info = get_calendar_info()
    return jsonify({
        "status": "ok",
        "excel_path": EXCEL_PATH,
        "sheets": sheets,
        "plants": PLANTS,
        "expeditions": EXPEDITIONS,
        "do_dates": do_dates,
        "default_date": do_dates[-1] if do_dates else "2026-10-30",
        "calendar": cal_info,
    })


@app.route("/api/calendar", methods=["GET"])
def api_calendar():
    return jsonify(get_calendar_info())


@app.route("/api/monthly_quota", methods=["GET"])
def api_monthly_quota():
    target_month = int(request.args.get("month", 10))
    target_year = int(request.args.get("year", 2026))
    cutoff_date = request.args.get("cutoff_date", "").strip() or None
    gudang = request.args.get("gudang", "ALL").strip()

    sl_records = db.get_sl_data()
    res = calculate_monthly_quota(
        sl_records,
        target_month=target_month,
        target_year=target_year,
        cutoff_date=cutoff_date,
        gudang_filter=gudang,
    )

    unit_rows = []
    for x in EXPEDITIONS:
        r = {"XPDC": x}
        for p in PLANTS:
            r[p] = int(res["unit_matrix"][x][p])
        r["GRAND_TOTAL"] = int(res["unit_matrix"][x]["GRAND_TOTAL"])
        r["PCT"] = round(float(res["unit_matrix"][x]["PCT_ALL_PLANT"]) * 100, 2)
        unit_rows.append(r)

    pct_rows = []
    for x in EXPEDITIONS:
        r = {"XPDC": x}
        for p in PLANTS:
            r[p] = round(float(res["pct_matrix"][x][p]) * 100, 2)
        r["TARGET_PCT"] = round(float(res["pct_matrix"][x]["TARGET_PCT"]) * 100, 2)
        pct_rows.append(r)

    return jsonify({
        "status": "ok",
        "target_month": res["target_month"],
        "target_month_name": res["target_month_name"],
        "target_year": res["target_year"],
        "periods": res["periods"],
        "cutoff_date": res["cutoff_date"],
        "total_3m_units": res["total_3m_units"],
        "plant_totals": {k: int(v) for k, v in res["plant_totals"].items()},
        "unit_rows": unit_rows,
        "pct_rows": pct_rows,
        "gudang_breakdown": res["gudang_breakdown"],
    })


@app.route("/api/daily_quota", methods=["POST"])
def api_daily_quota():
    payload = request.get_json(force=True)
    tanggal = parse_date_str(payload.get("tanggal", "2026-10-30")) or "2026-10-30"
    target_month = int(payload.get("month", int(tanggal[5:7])))
    target_year = int(payload.get("year", int(tanggal[:4])))
    cutoff_date = payload.get("cutoff_date", "2026-09-29") or None
    rolling_winner = payload.get("rolling_winner", "AUTO")
    mode = payload.get("mode", "operational_trailer")
    custom_do = payload.get("custom_do", None)

    sl_records = db.get_sl_data()
    do_records = db.get_do_data()

    do_dates = sorted(list({r["TANGGAL_STR"] for r in do_records}))
    if rolling_winner == "AUTO":
        if tanggal in do_dates:
            idx = do_dates.index(tanggal)
            rolling_winner = "TUNAS MUDA" if (idx % 2 == 0) else "RJTM"
        else:
            rolling_winner = "TUNAS MUDA"

    if custom_do:
        do_dict = {p: int(custom_do.get(p, 0)) for p in PLANTS}
    else:
        do_dict = {p: 0 for p in PLANTS}
        for r in do_records:
            if r["TANGGAL_STR"] == tanggal and r["PLANT"] in do_dict:
                do_dict[r["PLANT"]] += int(r["QTY"])

    mq = calculate_monthly_quota(
        sl_records, target_month=target_month, target_year=target_year, cutoff_date=cutoff_date
    )
    dq = calculate_and_balance_daily_quota(
        do_dict, mq["pct_matrix"], rolling_winner=rolling_winner, mode=mode
    )

    saved_all = db.get_saved_daily_allocations()
    saved_sub = [r for r in saved_all if r["TANGGAL"] == tanggal]
    saved_rows = []
    for x in EXPEDITIONS:
        sx = [r for r in saved_sub if r["XPDC"] == x]
        if sx:
            r_in = sx[0]
            r = {"XPDC": x}
            for p in PLANTS:
                r[p] = int(r_in[p])
            r["GRAND_TOTAL"] = int(r_in["GRAND_TOTAL"])
            r["PCT"] = round(float(r_in["PCT"]) * 100, 2)
            saved_rows.append(r)

    return jsonify({
        "status": "ok",
        "tanggal": tanggal,
        "rolling_winner": rolling_winner,
        "mode": mode,
        "do_per_plant": dq["do_per_plant"],
        "total_do": dq["total_do"],
        "all_within_tolerance_raw": dq["all_within_tolerance_raw"],
        "raw_rows": dq["raw_rows"],
        "balanced_rows": dq["balanced_rows"],
        "saved_rows": saved_rows,
        "trailer_checks": dq["trailer_checks"],
    })


@app.route("/api/save_daily_quota", methods=["POST"])
def api_save_daily_quota():
    payload = request.get_json(force=True)
    tanggal = payload.get("tanggal")
    do_per_plant = payload.get("do_per_plant", {})
    alloc_matrix = payload.get("alloc_matrix", {})

    db.create_backup()
    if do_per_plant:
        db.upsert_daily_do(tanggal, do_per_plant)
    db.save_daily_allocation(tanggal, alloc_matrix)

    return jsonify({
        "status": "ok",
        "message": f"Data DO & Alokasi Kuota Harian tanggal {tanggal} berhasil disimpan ke {EXCEL_PATH} (Sheet DO, PER DAY & ALOKASI_HARIAN)!",
    })



@app.route("/api/do_list", methods=["GET"])
def api_do_list():
    """Mengambil seluruh histori DO dari Sheet DO beserta kolom numerik BULAN."""
    sort_order = request.args.get("sort", "newest")  # 'newest' atau 'chronological'
    bulan_filter = request.args.get("bulan", "ALL").strip()
    tanggal_filter = request.args.get("tanggal", "").strip()

    do_records = db.get_do_data()
    if bulan_filter and bulan_filter != "ALL":
        try:
            b_int = int(bulan_filter)
            do_records = [r for r in do_records if int(r["BULAN"]) == b_int]
        except ValueError:
            pass

    if tanggal_filter:
        clean_t = parse_date_str(tanggal_filter)
        if clean_t:
            do_records = [r for r in do_records if r["TANGGAL_STR"] == clean_t]

    if sort_order == "chronological":
        do_records.sort(
            key=lambda r: (
                r["TANGGAL_STR"],
                PLANTS.index(r["PLANT"]) if r["PLANT"] in PLANTS else 99,
            )
        )
    else:
        # Terbaru di atas, namun urutan Plant per tanggal tetap FM1 s.d. FM9
        all_dates = sorted(list({r["TANGGAL_STR"] for r in do_records}), reverse=True)
        date_rank = {d: i for i, d in enumerate(all_dates)}
        do_records.sort(
            key=lambda r: (
                date_rank.get(r["TANGGAL_STR"], 9999),
                PLANTS.index(r["PLANT"]) if r["PLANT"] in PLANTS else 99,
            )
        )

    # Buat juga ringkasan per tanggal untuk kemudahan klik/pilih tanggal
    by_date = {}
    for r in db.get_do_data():
        d_str = r["TANGGAL_STR"]
        if d_str not in by_date:
            by_date[d_str] = {"TANGGAL": d_str, "BULAN": int(r["BULAN"]), "TOTAL_QTY": 0}
            for p in PLANTS:
                by_date[d_str][p] = 0
        by_date[d_str][r["PLANT"]] = int(r["QTY"])
        by_date[d_str]["TOTAL_QTY"] += int(r["QTY"])
        by_date[d_str]["BULAN"] = int(r["BULAN"])

    return jsonify({
        "status": "ok",
        "columns": ["TANGGAL", "PLANT", "QTY", "BULAN"],
        "rows": [
            {
                "TANGGAL": r["TANGGAL_STR"],
                "PLANT": r["PLANT"],
                "QTY": int(r["QTY"]),
                "BULAN": int(r["BULAN"]),
            }
            for r in do_records
        ],
        "by_date": by_date,
        "total_rows": len(do_records),
        "total_dates": len(by_date),
    })


@app.route("/api/add_do", methods=["POST"])
def api_add_do():
    payload = request.get_json(force=True)
    tanggal = parse_date_str(payload.get("tanggal"))
    if not tanggal:
        return jsonify({"status": "error", "message": "Tanggal DO tidak valid!"}), 400
    bulan = payload.get("bulan", None)
    do_per_plant = payload.get("do_per_plant", {})

    db.create_backup()
    db.upsert_daily_do(tanggal, do_per_plant, bulan_override=bulan)

    bulan_num = int(bulan) if (bulan is not None and str(bulan).strip() != "") else int(tanggal[5:7])
    tot_unit = sum(int(do_per_plant.get(p, 0)) for p in PLANTS)
    return jsonify({
        "status": "ok",
        "tanggal": tanggal,
        "bulan": bulan_num,
        "total_qty": tot_unit,
        "message": f"Data DO Harian tanggal {tanggal} (Bulan: {bulan_num}, Total: {tot_unit:,} unit) berhasil disimpan & terhubung ke histori Sheet DO!",
    })


@app.route("/api/perday", methods=["GET"])
def api_perday():
    """
    Mengambil data Kuota Harian yang sudah disimpan untuk ditampilkan pada Menu Per Day
    sesuai format 2 blok/kolom di Sheet 'PER DAY' (H-1 & H, lengkap dengan tabel MTD dan tabel Harian).
    """
    month_str = request.args.get("month", "10").strip()
    year_str = request.args.get("year", "2026").strip()
    tanggal_filter = request.args.get("tanggal", "").strip()
    sort_order = request.args.get("sort", "newest").strip()

    all_allocs = db.get_saved_daily_allocations()
    available_dates = sorted(list({r["TANGGAL"] for r in all_allocs}))

    filtered = all_allocs
    if month_str and month_str != "ALL":
        try:
            m_int = int(month_str)
            filtered = [r for r in filtered if int(r["BULAN"]) == m_int]
        except ValueError:
            pass
    if year_str and year_str != "ALL":
        try:
            y_int = int(year_str)
            filtered = [r for r in filtered if int(r["TAHUN"]) == y_int]
        except ValueError:
            pass

    # Hitung Rekap MTD dari seluruh tanggal di bulan terpilih (sebelum filter 1 tanggal)
    mtd_map = {x: {p: 0 for p in PLANTS} for x in EXPEDITIONS}
    for r in filtered:
        x = r["XPDC"]
        if x in mtd_map:
            for p in PLANTS:
                mtd_map[x][p] += int(r[p])

    mtd_grand_all = sum(sum(mtd_map[x][p] for p in PLANTS) for x in EXPEDITIONS)
    mtd_rows = []
    for x in EXPEDITIONS:
        row_d = {"EKSPEDISI": x}
        gt = 0
        for p in PLANTS:
            v = int(mtd_map[x][p])
            row_d[p] = v
            gt += v
        row_d["GRAND_TOTAL"] = gt
        row_d["PCT"] = round((gt / mtd_grand_all * 100.0), 2) if mtd_grand_all > 0 else 0.0
        mtd_rows.append(row_d)

    mtd_totals = {"EKSPEDISI": "TOTAL"}
    for p in PLANTS:
        mtd_totals[p] = sum(int(mtd_map[x][p]) for x in EXPEDITIONS)
    mtd_totals["GRAND_TOTAL"] = mtd_grand_all
    mtd_totals["PCT"] = 100.0 if mtd_grand_all > 0 else 0.0

    # Filter per tanggal spesifik jika dipilih untuk tabel rincian harian
    daily_filtered = filtered
    if tanggal_filter:
        clean_t = parse_date_str(tanggal_filter)
        if clean_t:
            daily_filtered = [r for r in daily_filtered if r["TANGGAL"] == clean_t]

    dates_in_view = sorted(
        list({r["TANGGAL"] for r in daily_filtered}),
        reverse=(sort_order == "newest"),
    )

    daily_groups = []
    flat_rows = []
    for d_str in dates_in_view:
        d_sub = {r["XPDC"]: r for r in daily_filtered if r["TANGGAL"] == d_str}
        d_tot_all = sum(int(d_sub.get(x, {}).get("GRAND_TOTAL", 0)) for x in EXPEDITIONS)
        group_rows = []
        for x in EXPEDITIONS:
            r_in = d_sub.get(x, {})
            row_d = {"TANGGAL": d_str, "EKSPEDISI": x}
            gt = 0
            for p in PLANTS:
                v = int(r_in.get(p, 0))
                row_d[p] = v
                gt += v
            row_d["GRAND_TOTAL"] = gt
            row_d["PCT"] = round((gt / d_tot_all * 100.0), 2) if d_tot_all > 0 else 0.0
            group_rows.append(row_d)
            flat_rows.append(row_d)

        d_totals = {"TANGGAL": d_str, "EKSPEDISI": "TOTAL"}
        for p in PLANTS:
            d_totals[p] = sum(int(r[p]) for r in group_rows)
        d_totals["GRAND_TOTAL"] = d_tot_all
        d_totals["PCT"] = 100.0 if d_tot_all > 0 else 0.0

        daily_groups.append({
            "tanggal": d_str,
            "rows": group_rows,
            "totals": d_totals,
        })

    from engine_ddc import MONTH_NAMES_ID
    m_name = MONTH_NAMES_ID.get(int(month_str), "ALL") if (month_str and month_str.isdigit()) else "ALL"

    return jsonify({
        "status": "ok",
        "month": month_str,
        "month_name": m_name,
        "year": year_str,
        "available_dates": available_dates,
        "total_dates": len(dates_in_view),
        "total_rows": len(flat_rows),
        "mtd_rows": mtd_rows,
        "mtd_totals": mtd_totals,
        "daily_groups": daily_groups,
        "flat_rows": flat_rows,
    })

@app.route("/api/summary", methods=["GET"])
def api_summary():
    cal_info = get_calendar_info()
    default_rep_date = cal_info.get("do_date_h_minus_1", "2026-11-08")
    report_date = parse_date_str(request.args.get("report_date", default_rep_date)) or default_rep_date
    cutoff_date = request.args.get("cutoff_date", "2026-09-29") or None
    unfilled_h_1 = int(request.args.get("unfilled_h_1", 379))

    rep_year = int(report_date[:4])
    rep_month = int(report_date[5:7])

    sl_records = db.get_sl_data()
    do_records = db.get_do_data()
    alloc_records = db.get_saved_daily_allocations()

    mq = calculate_monthly_quota(
        sl_records, target_month=rep_month, target_year=rep_year, cutoff_date=cutoff_date
    )
    target_pct_dict = {x: mq["pct_matrix"][x]["TARGET_PCT"] for x in EXPEDITIONS}

    # Cek apakah ada Target Bulanan Perencanaan yang tersimpan di sistem
    planned = db.get_target_bulanan(tahun=rep_year, bulan=rep_month)
    planned_targets = planned.get("TARGETS") if planned else None

    summ = calculate_summary_and_performance(
        sl_records,
        do_records,
        alloc_records,
        target_pct_dict,
        report_date_str=report_date,
        unfilled_h_minus_1=unfilled_h_1,
        planned_monthly_target=planned_targets,
    )
    if planned:
        summ["planned_target_info"] = {
            "tahun": planned.get("TAHUN"),
            "bulan": planned.get("BULAN"),
            "total_target": planned.get("TOTAL_TARGET"),
            "updated_at": planned.get("UPDATED_AT"),
            "catatan": planned.get("CATATAN", ""),
        }
    return jsonify({"status": "ok", "data": summ})


@app.route("/api/target_bulanan", methods=["GET"])
def api_get_target_bulanan():
    year = int(request.args.get("year", 2026))
    month = int(request.args.get("month", 10))
    total_target_param = request.args.get("total_target", None)

    planned = db.get_target_bulanan(tahun=year, bulan=month)
    all_targets = db.get_all_target_bulanan()

    sl_records = db.get_sl_data()
    mq = calculate_monthly_quota(sl_records, target_month=month, target_year=year)
    quota_pct = {x: mq["pct_matrix"][x]["TARGET_PCT"] for x in EXPEDITIONS}

    tot_target_val = 0
    if total_target_param and str(total_target_param).strip().isdigit():
        tot_target_val = int(total_target_param)
    elif planned and planned.get("TOTAL_TARGET"):
        tot_target_val = int(planned.get("TOTAL_TARGET"))
    else:
        tot_target_val = 85000

    from engine_ddc import _round_preserving_sum
    suggested = _round_preserving_sum(
        {x: quota_pct.get(x, 0.0) * tot_target_val for x in EXPEDITIONS},
        tot_target_val,
    )

    from engine_ddc import MONTH_NAMES_ID
    return jsonify({
        "status": "ok",
        "year": year,
        "month": month,
        "month_name": MONTH_NAMES_ID.get(month, str(month)),
        "current_target": planned,
        "all_targets": all_targets,
        "quota_3m_pct": {x: round(quota_pct.get(x, 0.0) * 100, 2) for x in EXPEDITIONS},
        "suggested_total": tot_target_val,
        "suggested_allocation": suggested,
    })


@app.route("/api/save_target_bulanan", methods=["POST"])
def api_save_target_bulanan():
    payload = request.get_json(force=True)
    year = int(payload.get("year", 2026))
    month = int(payload.get("month", 10))
    targets = payload.get("targets", {})
    total_target = payload.get("total_target", None)
    catatan = payload.get("catatan", "")

    saved = db.save_target_bulanan(
        tahun=year,
        bulan=month,
        target_per_xpdc_dict=targets,
        total_target=total_target,
        catatan=catatan,
    )
    return jsonify({
        "status": "ok",
        "message": f"Target Bulanan periode {month}/{year} berhasil disimpan!",
        "data": saved,
    })


@app.route("/api/sheet/<sheet_name>", methods=["GET"])
def api_get_sheet(sheet_name):
    max_rows = int(request.args.get("max_rows", 200))
    cols, rows = db.read_any_sheet(sheet_name, max_rows=max_rows)
    total_rows = len(db.get_sl_data()) if sheet_name.upper() == "SL" else len(rows)
    return jsonify({
        "status": "ok",
        "sheet_name": sheet_name,
        "columns": cols,
        "rows": rows,
        "total_rows": total_rows
    })


@app.route("/api/add_sheet", methods=["POST"])
def api_add_sheet():
    payload = request.get_json(force=True)
    sheet_name = payload.get("sheet_name", "").strip()
    rows = payload.get("rows", [])
    if not sheet_name:
        return jsonify({"status": "error", "message": "Nama sheet tidak boleh kosong"}), 400

    db.create_backup()
    db.add_or_replace_sheet(sheet_name, rows)
    return jsonify({
        "status": "ok",
        "message": f"Sheet '{sheet_name}' berhasil ditambahkan/diupdate di {EXCEL_PATH}!",
        "sheets": db.get_sheet_names(),
    })


@app.route("/api/add_sl", methods=["POST"])
def api_add_sl():
    try:
        p = request.get_json(force=True)
        db.create_backup()
        row_num = db.append_sl_row(
            tanggal=p["tanggal"],
            nopol=p["nopol"],
            sopir=p["sopir"],
            xpdc=p["xpdc"],
            plant=p["plant"],
            gudang=p["gudang"],
            qty=int(p["qty"]),
        )
        total_sl = len(db.get_sl_data())
        return jsonify({
            "status": "ok",
            "row_num": row_num,
            "total_sl": total_sl,
            "message": f"Data aktual pengiriman berhasil di-update ke Sheet SL pada file DDC.xlsx (Baris ke-{row_num}, Total: {total_sl:,} baris)!",
        })
    except PermissionError as e:
        return jsonify({"status": "error", "message": str(e)}), 409
    except PermissionError as e:
        return jsonify({"status": "error", "message": str(e)}), 409
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500



@app.route("/api/upload_sl", methods=["POST"])
def api_upload_sl():
    if "file" not in request.files:
        return jsonify({"status": "error", "message": "Tidak ada file yang diunggah"}), 400
    
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"status": "error", "message": "File kosong"}), 400
        
    try:
        import openpyxl
        wb = openpyxl.load_workbook(file, data_only=True)
        if "SL" in wb.sheetnames:
            ws = wb["SL"]
        elif "SL_UPLOAD" in wb.sheetnames:
            ws = wb["SL_UPLOAD"]
        else:
            ws = wb.active
        
        rows_list = []
        header = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                header = [str(c).strip().upper() if c is not None else "" for c in row]
                continue
            
            if not row or all(c is None or str(c).strip() == "" for c in row):
                continue
                
            raw_dict = {}
            for j, col_name in enumerate(header):
                if j < len(row) and col_name:
                    raw_dict[col_name] = row[j]
            
            # Normalisasi nama kolom jika ada variasi nama (misal Ekspedisi -> XPDC)
            norm_dict = {}
            for k, v in raw_dict.items():
                k_clean = str(k).strip().upper()
                if k_clean in ("TANGGAL", "DATE", "TGL", "TGL PENGIRIMAN"):
                    norm_dict["TANGGAL"] = v
                elif k_clean in ("NOPOL", "NO POLISI", "NO. POLISI", "PLAT", "NOPOL TRUK"):
                    norm_dict["NOPOL"] = v
                elif k_clean in ("SOPIR", "DRIVER", "SUPIR", "NAMA SOPIR"):
                    norm_dict["SOPIR"] = v
                elif k_clean in ("XPDC", "EKSPEDISI", "TRANSPORTER", "VENDOR"):
                    norm_dict["XPDC"] = v
                elif k_clean in ("PLANT", "PABRIK"):
                    norm_dict["PLANT"] = v
                elif k_clean in ("GUDANG", "GUDANG TUJUAN", "DESTINASI", "TUJUAN"):
                    norm_dict["GUDANG"] = v
                elif k_clean in ("QTY", "QTY (UNIT)", "JUMLAH", "UNIT", "QUANTITY"):
                    norm_dict["QTY"] = v
                else:
                    norm_dict[k_clean] = v
                    
            if "TANGGAL" in norm_dict and "XPDC" in norm_dict and "PLANT" in norm_dict:
                rows_list.append(norm_dict)
                
        wb.close()
        
        if not rows_list:
            return jsonify({
                "status": "error",
                "message": "Format file tidak sesuai atau data kosong. Pastikan kolom: TANGGAL (contoh: 31-Okt-26), NOPOL, SOPIR, XPDC, PLANT, GUDANG, QTY."
            }), 400
            
        db.create_backup()
        added = db.append_sl_bulk(rows_list)
        total_sl = len(db.get_sl_data())
        
        return jsonify({
            "status": "ok", 
            "added": added,
            "total_sl": total_sl,
            "message": f"Berhasil meng-update {added} data pengiriman ke Sheet SL pada file DDC.xlsx! Total data saat ini: {total_sl:,} baris."
        })
        
    except PermissionError as e:
        return jsonify({"status": "error", "message": str(e)}), 409
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/download_template_sl", methods=["GET"])
def api_download_template_sl():
    import openpyxl
    from io import BytesIO
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SL_UPLOAD"
    ws.append(["TANGGAL", "NOPOL", "SOPIR", "XPDC", "PLANT", "GUDANG", "QTY"])
    ws.append(["31-Okt-26", "B 1234 CD", "BUDI", "TUNAS MUDA", "FM6", "KARAWANG", 64])
    ws.append(["31-Okt-26", "B 9999 XX", "ANDI", "RJTM", "FM8", "BANDUNG", 54])
    
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 15
    ws.column_dimensions["G"].width = 10

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, as_attachment=True, download_name="Template_Upload_SL.xlsx")

@app.route("/api/download_excel", methods=["GET"])
def api_download_excel():
    target_path = getattr(db, "excel_path", EXCEL_PATH)
    return send_file(target_path, as_attachment=True, download_name="DDC.xlsx")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
