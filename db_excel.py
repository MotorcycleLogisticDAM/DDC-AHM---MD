"""
db_excel.py - Modul Database Lokal Excel (DDC.xlsx) Pure Python (Tanpa C++ / Pandas)
====================================================================================
Menggunakan murni `openpyxl` dan pustaka bawaan Python sehingga dapat berjalan
langsung di Python 3.11 s.d. Python 3.15+ di Windows tanpa memerlukan Visual Studio C++.
"""

import os
import re
import shutil
from datetime import datetime, date
import openpyxl

DEFAULT_EXCEL_PATH = os.environ.get("DDC_EXCEL_PATH", "DDC.xlsx")
PLANTS = ["FM1", "FM2", "FM6", "FM6A", "FM7", "FM8", "FM9"]
EXPEDITIONS = ["TUNAS MUDA", "RJTM", "LBK", "WSS", "PUNINAR"]


MONTH_LOOKUP = {
    "jan": 1, "januari": 1, "january": 1,
    "feb": 2, "februari": 2, "february": 2,
    "mar": 3, "maret": 3, "march": 3,
    "apr": 4, "april": 4,
    "mei": 5, "may": 5,
    "jun": 6, "juni": 6, "june": 6,
    "jul": 7, "juli": 7, "july": 7,
    "agu": 8, "agustus": 8, "aug": 8, "august": 8,
    "sep": 9, "september": 9,
    "okt": 10, "oktober": 10, "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "des": 12, "desember": 12, "dec": 12, "december": 12,
}


def parse_date_str(val):
    """
    Mengubah nilai tanggal Excel (datetime / string / format '31-Okt-26' / 'DD-Mmm-YY')
    menjadi string 'YYYY-MM-DD'.
    """
    if val is None or val == "":
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, date):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if not s or s.lower() in ("nan", "nat", "none"):
        return None

    # 1. Parsing format dengan nama/singkatan bulan (misal: '31-Okt-26', '31-Okt-2026', '04-Okt-26')
    m_regex = re.match(r"^(\d{1,2})[-/\s]([a-zA-Z]+)[-/\s](\d{2,4})$", s)
    if m_regex:
        d = int(m_regex.group(1))
        mon_str = m_regex.group(2).lower()
        y_str = m_regex.group(3)
        if mon_str in MONTH_LOOKUP:
            mon = MONTH_LOOKUP[mon_str]
            y = int(y_str)
            if y < 100:
                y += 2000
            return f"{y:04d}-{mon:02d}-{d:02d}"

    # 2. Parsing format tanggal standar numerik
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s[:19], fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(s[:10]).strftime("%Y-%m-%d")
    except Exception:
        return None


MONTH_ABBR_ID = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr",
    5: "Mei", 6: "Jun", 7: "Jul", 8: "Agu",
    9: "Sep", 10: "Okt", 11: "Nov", 12: "Des"
}


def format_date_d_mmm_yy(val):
    """
    Mengubah format tanggal 'YYYY-MM-DD' atau datetime/date menjadi 'DD-Mmm-YY',
    contoh: '2026-10-31' -> '31-Okt-26'.
    """
    if not val:
        return ""
    t_str = parse_date_str(val) if not isinstance(val, str) or "-" not in val else val
    if not t_str or len(t_str) < 10:
        return str(val)
    try:
        parts = t_str[:10].split("-")
        y, m, d = parts[0], int(parts[1]), parts[2]
        m_name = MONTH_ABBR_ID.get(m, f"{m:02d}")
        yy = y[-2:]
        return f"{d}-{m_name}-{yy}"
    except Exception:
        return str(val)


def to_int(val, default=0):
    """Mengubah nilai angka atau rumus aritmatika sederhana Excel (misal '=2167-50') menjadi integer."""
    if val is None or val == "":
        return default
    if isinstance(val, (int, float)):
        return int(round(float(val)))
    s = str(val).strip()
    if not s:
        return default
    if s.startswith("="):
        expr = s[1:].strip()
        if re.match(r"^[0-9+\-*/().\s]+$", expr):
            try:
                return int(round(float(eval(expr, {"__builtins__": None}, {}))))
            except Exception:
                return default
        return default
    try:
        return int(round(float(s)))
    except Exception:
        return default


def to_float(val, default=0.0):
    if val is None or val == "":
        return default
    try:
        return float(val)
    except Exception:
        return default


class ExcelDatabase:
    def __init__(self, excel_path=DEFAULT_EXCEL_PATH):
        self.excel_path = excel_path

    # ------------------------------------------------------------------
    # 1. FUNGSI MANAJEMEN SHEET (LIHAT, TAMBAH, UPDATE, BACKUP)
    # ------------------------------------------------------------------
    def get_sheet_names(self):
        """Mengembalikan daftar seluruh nama sheet yang ada di DDC.xlsx."""
        if not os.path.exists(self.excel_path):
            return []
        wb = openpyxl.load_workbook(self.excel_path, read_only=True, data_only=True)
        names = wb.sheetnames
        wb.close()
        return names

    def create_backup(self):
        """Membuat salinan backup otomatis sebelum melakukan perubahan besar."""
        try:
            base_dir = os.path.dirname(self.excel_path) or "."
            backup_dir = os.path.join(base_dir, "backups")
            os.makedirs(backup_dir, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = os.path.join(backup_dir, f"DDC_backup_{ts}.xlsx")
            shutil.copy(self.excel_path, backup_path)
            return backup_path
        except Exception:
            return None

    def read_any_sheet(self, sheet_name, max_rows=500):
        """
        Membaca sheet apa saja menjadi list of dicts untuk ditampilkan di Web UI.
        Khusus untuk sheet 'DO' dan 'SL', gunakan pembacaan terstruktur agar header,
        urutan tanggal terbaru, dan kolom numerik BULAN tampil akurat.
        """
        s_upper = str(sheet_name).strip().upper()
        if s_upper == "DO":
            do_recs = self.get_do_data()
            # Urutkan dari tanggal terbaru ke terlama (namun urutan Plant tetap FM1..FM9)
            do_recs_sorted = sorted(
                do_recs,
                key=lambda r: (
                    r["TANGGAL_STR"],
                    -(PLANTS.index(r["PLANT"]) if r["PLANT"] in PLANTS else 99),
                ),
                reverse=True,
            )
            if max_rows and max_rows > 0:
                do_recs_sorted = do_recs_sorted[:max_rows]
            cols = ["TANGGAL", "PLANT", "QTY", "BULAN"]
            rows = [
                {
                    "TANGGAL": r["TANGGAL_STR"],
                    "PLANT": r["PLANT"],
                    "QTY": str(r["QTY"]),
                    "BULAN": str(r["BULAN"]),
                }
                for r in do_recs_sorted
            ]
            return cols, rows

        if s_upper == "SL":
            sl_recs = self.get_sl_data()
            # Tampilkan baris terbaru terlebih dahulu agar input/upload baru langsung terlihat
            sl_recs_rev = list(reversed(sl_recs))
            if max_rows and max_rows > 0:
                sl_recs_rev = sl_recs_rev[:max_rows]
            cols = ["TANGGAL", "NOPOL", "SOPIR", "XPDC", "PLANT", "GUDANG", "QTY", "MONTH", "DAY", "YEAR"]
            rows = [
                {
                    "TANGGAL": format_date_d_mmm_yy(r["TANGGAL_STR"]),
                    "NOPOL": r.get("NOPOL", ""),
                    "SOPIR": r.get("SOPIR", ""),
                    "XPDC": r["XPDC"],
                    "PLANT": r["PLANT"],
                    "GUDANG": r["GUDANG"],
                    "QTY": str(r["QTY"]),
                    "MONTH": str(r["MONTH"]),
                    "DAY": str(r["DAY"]),
                    "YEAR": str(r["YEAR"]),
                }
                for r in sl_recs_rev
            ]
            return cols, rows

        wb = openpyxl.load_workbook(self.excel_path, read_only=True, data_only=True)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return [], []
        ws = wb[sheet_name]
        rows_iter = ws.iter_rows(values_only=True)
        try:
            first_row = next(rows_iter)
        except StopIteration:
            wb.close()
            return [], []

        cols = [
            str(c).strip() if (c is not None and str(c).strip() != "") else f"COL_{i+1}"
            for i, c in enumerate(first_row)
        ]
        seen = {}
        unique_cols = []
        for c in cols:
            if c in seen:
                seen[c] += 1
                unique_cols.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                unique_cols.append(c)

        data_rows = []
        for idx, r in enumerate(rows_iter):
            if max_rows and idx >= max_rows:
                break
            if not r or all(v is None or str(v).strip() == "" for v in r):
                continue
            row_dict = {}
            for c_idx, col_name in enumerate(unique_cols):
                val = r[c_idx] if c_idx < len(r) else ""
                if isinstance(val, (datetime, date)):
                    val = val.strftime("%Y-%m-%d")
                elif val is None:
                    val = ""
                row_dict[col_name] = str(val)
            data_rows.append(row_dict)

        wb.close()
        return unique_cols, data_rows

    def add_or_replace_sheet(self, sheet_name, rows_list):
        """
        Menambahkan sheet baru atau meng-update isi sheet tertentu di DDC.xlsx
        TANPA menghapus sheet lainnya.
        """
        wb = openpyxl.load_workbook(self.excel_path)
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        ws = wb.create_sheet(title=sheet_name)

        if rows_list and len(rows_list) > 0:
            headers = list(rows_list[0].keys())
            ws.append(headers)
            for r in rows_list:
                ws.append([r.get(h, "") for h in headers])
        else:
            ws.append(["KOLOM_1", "KOLOM_2"])

        wb.save(self.excel_path)
        wb.close()
        return True

    # ------------------------------------------------------------------
    # 2. FUNGSI BACA & TAMBAH DATA INPUT: SHEET 'SL' & SHEET 'DO'
    # ------------------------------------------------------------------
    def get_sl_data(self):
        """Membaca data distribusi aktual dari sheet 'SL' menjadi list of dicts."""
        wb = openpyxl.load_workbook(self.excel_path, read_only=True, data_only=True)
        ws = wb["SL"]
        records = []
        for idx, row in enumerate(ws.iter_rows(values_only=True)):
            if idx == 0:
                continue
            if not row or len(row) < 7:
                continue
            tgl_str = parse_date_str(row[0])
            xpdc = str(row[3] or "").strip().upper()
            plant = str(row[4] or "").strip().upper()
            if not tgl_str or not xpdc or not plant or xpdc == "NAN" or plant == "NAN":
                continue
            dt = datetime.strptime(tgl_str, "%Y-%m-%d")
            nopol = str(row[1] or "").strip().upper()
            sopir = str(row[2] or "").strip().upper()
            gudang = str(row[5] or "ALL").strip().upper()
            qty = to_int(row[6], 0)
            month = to_int(row[7], dt.month) if (len(row) > 7 and row[7] is not None) else dt.month
            day = to_int(row[8], dt.day) if (len(row) > 8 and row[8] is not None) else dt.day
            year = to_int(row[9], dt.year) if (len(row) > 9 and row[9] is not None) else dt.year
            records.append({
                "TANGGAL_STR": tgl_str,
                "NOPOL": nopol,
                "SOPIR": sopir,
                "XPDC": xpdc,
                "PLANT": plant,
                "GUDANG": gudang,
                "QTY": qty,
                "MONTH": month,
                "DAY": day,
                "YEAR": year,
            })
        wb.close()
        return records

    def _find_last_data_row(self, ws, start_row=2):
        """Mencari nomor baris terakhir yang benar-benar berisi data secara cepat (scan mundur)."""
        max_r = ws.max_row
        for r in range(max_r, start_row - 1, -1):
            v1 = ws.cell(r, 1).value
            v4 = ws.cell(r, 4).value
            if (v1 is not None and str(v1).strip() != "") or (v4 is not None and str(v4).strip() != ""):
                return r
        return start_row - 1

    def append_sl_row(self, tanggal, nopol, sopir, xpdc, plant, gudang, qty):
        """Menambahkan baris pengiriman baru ke bagian paling bawah sheet 'SL'."""
        dt = datetime.strptime(parse_date_str(tanggal), "%Y-%m-%d")
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb["SL"]
        next_row = self._find_last_data_row(ws, start_row=2) + 1
        c_tgl = ws.cell(next_row, 1, dt)
        c_tgl.number_format = "d-mmm-yy"
        ws.cell(next_row, 2, str(nopol).upper().strip())
        ws.cell(next_row, 3, str(sopir).upper().strip())
        ws.cell(next_row, 4, str(xpdc).upper().strip())
        ws.cell(next_row, 5, str(plant).upper().strip())
        ws.cell(next_row, 6, str(gudang).upper().strip())
        ws.cell(next_row, 7, int(qty))
        ws.cell(next_row, 8, int(dt.month))
        ws.cell(next_row, 9, int(dt.day))
        ws.cell(next_row, 10, int(dt.year))
        try:
            wb.save(self.excel_path)
            wb.close()
        except PermissionError:
            wb.close()
            raise PermissionError(
                f"File '{os.path.basename(self.excel_path)}' sedang dibuka atau dikunci oleh program lain (seperti Microsoft Excel). "
                f"Mohon tutup Microsoft Excel terlebih dahulu agar data dapat disimpan."
            )
        return next_row

    def append_sl_bulk(self, rows_list):
        """Menambahkan banyak baris pengiriman baru sekaligus ke bagian bawah sheet 'SL' pada file DDC.xlsx."""
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb["SL"]
        start_row = self._find_last_data_row(ws, start_row=2) + 1
        current_row = start_row

        for r in rows_list:
            try:
                dt_str = parse_date_str(r.get("TANGGAL", ""))
                if not dt_str:
                    continue
                dt = datetime.strptime(dt_str, "%Y-%m-%d")
                qty = to_int(r.get("QTY", r.get("QTY (UNIT)", 0)), 0)
                if qty == 0:
                    continue

                c_tgl = ws.cell(current_row, 1, dt)
                c_tgl.number_format = "d-mmm-yy"
                ws.cell(current_row, 2, str(r.get("NOPOL", "")).upper().strip())
                ws.cell(current_row, 3, str(r.get("SOPIR", "")).upper().strip())
                ws.cell(current_row, 4, str(r.get("XPDC", "")).upper().strip())
                ws.cell(current_row, 5, str(r.get("PLANT", "")).upper().strip())
                ws.cell(current_row, 6, str(r.get("GUDANG", "ALL")).upper().strip())
                ws.cell(current_row, 7, qty)
                ws.cell(current_row, 8, int(dt.month))
                ws.cell(current_row, 9, int(dt.day))
                ws.cell(current_row, 10, int(dt.year))
                current_row += 1
            except Exception:
                pass

        try:
            wb.save(self.excel_path)
            wb.close()
        except PermissionError:
            wb.close()
            raise PermissionError(
                f"File '{os.path.basename(self.excel_path)}' sedang dibuka atau dikunci oleh program lain (seperti Microsoft Excel). "
                f"Mohon tutup Microsoft Excel terlebih dahulu agar data dapat disimpan."
            )
        return current_row - start_row

    def get_do_data(self):
        """
        Membaca data DO Harian dari sheet 'DO' menjadi list of dicts.
        Menangani baik nilai cache maupun sel rumus Excel (seperti '=MONTH(A4)' atau '=2167-50')
        sehingga tidak ada data atau angka BULAN yang hilang.
        """
        wb_val = openpyxl.load_workbook(self.excel_path, read_only=True, data_only=True)
        ws_val = wb_val["DO"]
        val_rows = list(ws_val.iter_rows(values_only=True))
        wb_val.close()

        wb_raw = openpyxl.load_workbook(self.excel_path, read_only=True, data_only=False)
        ws_raw = wb_raw["DO"]
        raw_rows = list(ws_raw.iter_rows(values_only=True))
        wb_raw.close()

        records = []
        max_len = max(len(val_rows), len(raw_rows))
        for idx in range(max_len):
            r_val = val_rows[idx] if idx < len(val_rows) else ()
            r_raw = raw_rows[idx] if idx < len(raw_rows) else ()
            if not r_val and not r_raw:
                continue

            raw_tgl = r_val[0] if (len(r_val) > 0 and r_val[0] is not None) else (r_raw[0] if len(r_raw) > 0 else None)
            raw_plant = r_val[1] if (len(r_val) > 1 and r_val[1] is not None) else (r_raw[1] if len(r_raw) > 1 else None)
            tgl_str = parse_date_str(raw_tgl)
            plant = str(raw_plant or "").strip().upper()
            if not tgl_str or plant not in PLANTS:
                continue

            dt = datetime.strptime(tgl_str, "%Y-%m-%d")
            raw_qty = r_val[2] if (len(r_val) > 2 and r_val[2] is not None) else (r_raw[2] if len(r_raw) > 2 else 0)
            qty = to_int(raw_qty, 0)

            raw_bulan = r_val[3] if (len(r_val) > 3 and r_val[3] is not None) else (r_raw[3] if len(r_raw) > 3 else None)
            bulan = to_int(raw_bulan, dt.month)
            if bulan < 1 or bulan > 12:
                bulan = int(dt.month)

            # Kolom E: QTY_AWAL, Kolom F: ADJUSTMENT, Kolom G: KETERANGAN
            raw_qty_awal = r_val[4] if (len(r_val) > 4 and r_val[4] is not None) else (r_raw[4] if len(r_raw) > 4 else None)
            qty_awal = to_int(raw_qty_awal, qty) if raw_qty_awal is not None else qty

            raw_adj = r_val[5] if (len(r_val) > 5 and r_val[5] is not None) else (r_raw[5] if len(r_raw) > 5 else None)
            adjustment = to_int(raw_adj, 0) if raw_adj is not None else (qty - qty_awal)

            raw_ket = r_val[6] if (len(r_val) > 6 and r_val[6] is not None) else (r_raw[6] if len(r_raw) > 6 else "")
            keterangan = str(raw_ket or "").strip()
            if not keterangan:
                if adjustment > 0:
                    keterangan = f"Penambahan DO (+{adjustment})"
                elif adjustment < 0:
                    keterangan = "Pembatalan DO" if qty == 0 else f"Pengurangan DO ({adjustment})"
                else:
                    keterangan = "Normal"

            records.append({
                "TANGGAL_STR": tgl_str,
                "PLANT": plant,
                "QTY": qty,
                "QTY_AWAL": qty_awal,
                "ADJUSTMENT": adjustment,
                "QTY_FINAL": qty,
                "KETERANGAN": keterangan,
                "BULAN": int(bulan),
                "TAHUN": int(dt.year),
            })
        return records

    def upsert_daily_do(self, tanggal_str, do_dict_per_plant, bulan_override=None):
        """
        Menyimpan atau meng-update DO Harian untuk 7 Plant pada tanggal tertentu di sheet 'DO'.
        Setiap baris DO selalu menyimpan kolom BULAN dalam bentuk angka numerik (1-12)
        dan membersihkan baris kosong berlebih agar histori tersusun rapi sesuai Sheet DO.
        """
        clean_tgl = parse_date_str(tanggal_str)
        dt = datetime.strptime(clean_tgl, "%Y-%m-%d")
        bulan = to_int(bulan_override, int(dt.month))
        if bulan < 1 or bulan > 12:
            bulan = int(dt.month)

        # Baca seluruh data DO yang ada terlebih dahulu agar rumus lama ikut terkonversi menjadi angka pasti
        existing_records = self.get_do_data()

        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb["DO"]

        # Pastikan Header Baris 1 & Baris 3 sesuai format asli Sheet DO
        if ws.cell(1, 1).value is None:
            ws.cell(1, 1, "DO OPEN AHM")
        ws.cell(3, 1, "TANGGAL")
        ws.cell(3, 2, "PLANT")
        ws.cell(3, 3, "QTY")
        ws.cell(3, 4, "BULAN")
        ws.cell(3, 5, "QTY_AWAL")
        ws.cell(3, 6, "ADJUSTMENT")
        ws.cell(3, 7, "KETERANGAN")

        # Gabungkan data lama dengan data tanggal yang baru diinput/diupdate
        do_map = {}
        for rec in existing_records:
            do_map[(rec["TANGGAL_STR"], rec["PLANT"])] = (
                rec["QTY"],
                rec["BULAN"],
                rec.get("QTY_AWAL", rec["QTY"]),
                rec.get("ADJUSTMENT", 0),
                rec.get("KETERANGAN", "Normal"),
            )

        for p in PLANTS:
            qty_val = to_int(do_dict_per_plant.get(p, 0), 0)
            prev = do_map.get((clean_tgl, p))
            if prev and prev[0] != qty_val:
                # Ada perubahan manual dari form submit
                q_awal = prev[2]
                adj = qty_val - q_awal
                ket = f"Update Qty ({adj:+d})" if adj != 0 else prev[4]
                do_map[(clean_tgl, p)] = (qty_val, int(bulan), q_awal, adj, ket)
            elif prev:
                do_map[(clean_tgl, p)] = (qty_val, int(bulan), prev[2], prev[3], prev[4])
            else:
                do_map[(clean_tgl, p)] = (qty_val, int(bulan), qty_val, 0, "Normal")

        # Urutkan berdasarkan Tanggal (Kronologis sesuai Sheet DO) dan urutan Plant (FM1..FM9)
        all_dates = sorted(list({k[0] for k in do_map.keys()}))

        # Bersihkan baris 4 ke bawah lalu tulis ulang secara rapi tanpa baris kosong di tengah
        if ws.max_row >= 4:
            ws.delete_rows(4, ws.max_row - 3)

        curr_row = 4
        for d_str in all_dates:
            d_dt = datetime.strptime(d_str, "%Y-%m-%d")
            for p in PLANTS:
                if (d_str, p) in do_map:
                    q_val, b_val, q_awal, adj, ket = do_map[(d_str, p)]
                else:
                    q_val, b_val, q_awal, adj, ket = 0, int(d_dt.month), 0, 0, "Normal"
                cell_tgl = ws.cell(curr_row, 1, d_dt)
                cell_tgl.number_format = "yyyy-mm-dd"
                ws.cell(curr_row, 2, p)
                ws.cell(curr_row, 3, int(q_val))
                ws.cell(curr_row, 4, int(b_val))
                ws.cell(curr_row, 5, int(q_awal))
                ws.cell(curr_row, 6, int(adj))
                ws.cell(curr_row, 7, str(ket))
                curr_row += 1

        wb.save(self.excel_path)
        wb.close()
        return True

    def adjust_do(self, tanggal_str, plant, adjustment_val=None, keterangan="", qty_final=None, qty_awal=None):
        """
        Menyimpan adjustment khusus untuk satu baris (TANGGAL, PLANT) pada Sheet 'DO'.
        Memperbarui kolom QTY (netto aktif), QTY_AWAL, ADJUSTMENT (+/-), dan KETERANGAN.
        """
        clean_tgl = parse_date_str(tanggal_str)
        if not clean_tgl:
            raise ValueError(f"Tanggal tidak valid: {tanggal_str}")
        plant = str(plant or "").strip().upper()
        if plant not in PLANTS:
            raise ValueError(f"Plant tidak valid: {plant}")

        self.create_backup()
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb["DO"]

        # Pastikan Header Baris 3
        ws.cell(3, 1, "TANGGAL")
        ws.cell(3, 2, "PLANT")
        ws.cell(3, 3, "QTY")
        ws.cell(3, 4, "BULAN")
        ws.cell(3, 5, "QTY_AWAL")
        ws.cell(3, 6, "ADJUSTMENT")
        ws.cell(3, 7, "KETERANGAN")

        target_row = None
        current_qty = 0
        current_qty_awal = None

        for row_idx in range(4, ws.max_row + 1):
            cell_tgl = ws.cell(row_idx, 1).value
            cell_plant = str(ws.cell(row_idx, 2).value or "").strip().upper()
            tgl_row = parse_date_str(cell_tgl)
            if tgl_row == clean_tgl and cell_plant == plant:
                target_row = row_idx
                current_qty = to_int(ws.cell(row_idx, 3).value, 0)
                current_qty_awal = ws.cell(row_idx, 5).value
                break

        if not target_row:
            wb.close()
            raise ValueError(f"Baris DO untuk {clean_tgl} plant {plant} tidak ditemukan di Sheet DO!")

        # Tentukan Qty Awal
        if current_qty_awal is not None:
            base_awal = to_int(current_qty_awal, current_qty)
        elif qty_awal is not None:
            base_awal = to_int(qty_awal, current_qty)
        else:
            base_awal = current_qty

        # Hitung Qty Final dan Adjustment
        if qty_final is not None and adjustment_val is None:
            final_q = max(0, to_int(qty_final, 0))
            adj = final_q - base_awal
        elif adjustment_val is not None:
            adj = to_int(adjustment_val, 0)
            if qty_final is not None:
                final_q = max(0, to_int(qty_final, 0))
            else:
                final_q = max(0, base_awal + adj)
        else:
            final_q = current_qty
            adj = 0

        # Keterangan default jika kosong
        ket = str(keterangan or "").strip()
        if not ket:
            if final_q == 0 and base_awal > 0:
                ket = "Pembatalan DO"
            elif adj < 0:
                ket = f"Pengurangan DO ({adj})"
            elif adj > 0:
                ket = f"Penambahan DO (+{adj})"
            else:
                ket = "Normal"

        # Simpan ke baris target di Excel
        ws.cell(target_row, 3, int(final_q))
        ws.cell(target_row, 5, int(base_awal))
        ws.cell(target_row, 6, int(adj))
        ws.cell(target_row, 7, str(ket))

        wb.save(self.excel_path)
        wb.close()
        return {
            "tanggal": clean_tgl,
            "plant": plant,
            "qty_awal": int(base_awal),
            "adjustment": int(adj),
            "qty_final": int(final_q),
            "keterangan": ket,
        }

    # ------------------------------------------------------------------
    # 3. FUNGSI BACA & SIMPAN ALOKASI KUOTA HARIAN ('ALOKASI_HARIAN' & 'PER DAY')
    # ------------------------------------------------------------------
    def get_saved_daily_allocations(self):
        """
        Membaca data alokasi kuota harian yang sudah disimpan.
        Mengutamakan pembacaan dari sheet 'ALOKASI_HARIAN' dan menggabungkannya dengan
        sheet 'PER DAY' (menangani baik nilai angka maupun rumus seperti '=590-14').
        """
        wb_val = openpyxl.load_workbook(self.excel_path, read_only=True, data_only=True)
        wb_raw = openpyxl.load_workbook(self.excel_path, read_only=True, data_only=False)
        records_map = {}

        # 1. Baca dari sheet 'PER DAY' terlebih dahulu sebagai baseline
        if "PER DAY" in wb_val.sheetnames:
            ws_v = wb_val["PER DAY"]
            ws_r = wb_raw["PER DAY"]
            v_rows = list(ws_v.iter_rows(values_only=True))
            r_rows = list(ws_r.iter_rows(values_only=True))
            max_len = max(len(v_rows), len(r_rows))
            for idx in range(12, max_len):
                rv = v_rows[idx] if idx < len(v_rows) else ()
                rr = r_rows[idx] if idx < len(r_rows) else ()
                if not rv and not rr:
                    continue
                raw_tgl = rv[0] if (len(rv) > 0 and rv[0] is not None) else (rr[0] if len(rr) > 0 else None)
                raw_xpdc = rv[1] if (len(rv) > 1 and rv[1] is not None) else (rr[1] if len(rr) > 1 else None)
                tgl_str = parse_date_str(raw_tgl)
                xpdc = str(raw_xpdc or "").strip().upper()
                if not tgl_str or xpdc not in EXPEDITIONS:
                    continue
                dt = datetime.strptime(tgl_str, "%Y-%m-%d")
                rec = {"TANGGAL": tgl_str, "XPDC": xpdc}
                for p_i, p in enumerate(PLANTS):
                    cell_v = rv[2 + p_i] if (len(rv) > 2 + p_i and rv[2 + p_i] is not None) else (rr[2 + p_i] if len(rr) > 2 + p_i else 0)
                    rec[p] = to_int(cell_v, 0)
                rec["GRAND_TOTAL"] = sum(rec[p] for p in PLANTS)
                cell_pct = rv[10] if (len(rv) > 10 and rv[10] is not None) else (rr[10] if len(rr) > 10 else 0.0)
                rec["PCT"] = to_float(cell_pct, 0.0)
                rec["BULAN"] = int(dt.month)
                rec["TAHUN"] = int(dt.year)
                records_map[(tgl_str, xpdc)] = rec

        # 2. Timpa/lengkapi dengan data terbaru dari sheet 'ALOKASI_HARIAN'
        if "ALOKASI_HARIAN" in wb_val.sheetnames:
            ws_a = wb_val["ALOKASI_HARIAN"]
            for idx, row in enumerate(ws_a.iter_rows(values_only=True)):
                if idx == 0 or not row or len(row) < 11:
                    continue
                tgl_str = parse_date_str(row[0])
                xpdc = str(row[1] or "").strip().upper()
                if not tgl_str or xpdc not in EXPEDITIONS:
                    continue
                dt = datetime.strptime(tgl_str, "%Y-%m-%d")
                rec = {"TANGGAL": tgl_str, "XPDC": xpdc}
                for p_i, p in enumerate(PLANTS):
                    rec[p] = to_int(row[2 + p_i], 0)
                rec["GRAND_TOTAL"] = to_int(row[9], sum(rec[p] for p in PLANTS))
                rec["PCT"] = to_float(row[10], 0.0)
                rec["BULAN"] = to_int(row[11], dt.month) if len(row) > 11 else dt.month
                rec["TAHUN"] = to_int(row[12], dt.year) if len(row) > 12 else dt.year
                records_map[(tgl_str, xpdc)] = rec

        wb_val.close()
        wb_raw.close()

        all_recs = list(records_map.values())
        all_recs.sort(
            key=lambda r: (
                r["TANGGAL"],
                EXPEDITIONS.index(r["XPDC"]) if r["XPDC"] in EXPEDITIONS else 99,
            )
        )
        return all_recs

    def save_daily_allocation(self, tanggal_str, alloc_matrix_dict):
        """
        Menyimpan hasil kalkulasi/balancing (data yang sudah di-adjust) Kuota Harian ke:
        1) Sheet 'ALOKASI_HARIAN'
        2) Sheet 'PER DAY' pada 2 blok/kolom yang sudah disiapkan di Excel:
           - Blok Kiri (H-1: Kolom A s.d. K) & Blok Kanan (H: Kolom M s.d. W)
           - Tabel Atas (Rekap MTD Baris 5 s.d. 10) & Tabel Bawah (Rincian Tanggal Baris 14+)
        """
        clean_tgl = parse_date_str(tanggal_str)
        dt = datetime.strptime(clean_tgl, "%Y-%m-%d")
        existing = [
            r for r in self.get_saved_daily_allocations() if r["TANGGAL"] != clean_tgl
        ]

        total_all = sum(
            sum(to_int(alloc_matrix_dict.get(x, {}).get(p, 0), 0) for p in PLANTS)
            for x in EXPEDITIONS
        )

        new_rows = []
        for x in EXPEDITIONS:
            r_data = {"TANGGAL": clean_tgl, "XPDC": x}
            gt = 0
            for p in PLANTS:
                val = to_int(alloc_matrix_dict.get(x, {}).get(p, 0), 0)
                r_data[p] = val
                gt += val
            r_data["GRAND_TOTAL"] = gt
            r_data["PCT"] = round(gt / total_all, 6) if total_all > 0 else 0.0
            r_data["BULAN"] = int(dt.month)
            r_data["TAHUN"] = int(dt.year)
            new_rows.append(r_data)

        all_rows = existing + new_rows
        all_rows.sort(
            key=lambda r: (
                r["TANGGAL"],
                EXPEDITIONS.index(r["XPDC"]) if r["XPDC"] in EXPEDITIONS else 99,
            )
        )

        self.add_or_replace_sheet("ALOKASI_HARIAN", all_rows)

        # Update penuh Sheet 'PER DAY' (Kedua blok H-1 [Kolom A-K] dan H [Kolom M-W])
        wb = openpyxl.load_workbook(self.excel_path)
        if "PER DAY" in wb.sheetnames:
            ws = wb["PER DAY"]

            # 1. Hitung Rekap MTD untuk bulan dari tanggal yang disimpan
            target_m = int(dt.month)
            target_y = int(dt.year)
            month_rows = [
                r for r in all_rows
                if int(r["BULAN"]) == target_m and int(r["TAHUN"]) == target_y
            ]
            mtd_by_xpdc = {x: {p: 0 for p in PLANTS} for x in EXPEDITIONS}
            for r in month_rows:
                x = r["XPDC"]
                if x in mtd_by_xpdc:
                    for p in PLANTS:
                        mtd_by_xpdc[x][p] += int(r[p])

            mtd_grand_all = sum(
                sum(mtd_by_xpdc[x][p] for p in PLANTS) for x in EXPEDITIONS
            )

            # Tulis Header & Rekap MTD di Baris 4-10 pada Blok Kiri (offset=0) & Blok Kanan (offset=12)
            for offset, label_h in [(0, "H-1 "), (12, "H")]:
                ws.cell(2, 1 + offset, "MONITORING QUOTA HARIAN (unit)")
                ws.cell(2, 7 + offset, label_h)
                ws.cell(2, 10 + offset, "MONTH :")
                ws.cell(4, 1 + offset, "MTD")
                ws.cell(4, 2 + offset, "EKSPEDISI")
                for p_i, p in enumerate(PLANTS):
                    ws.cell(4, 3 + p_i + offset, p)
                ws.cell(4, 10 + offset, "Grand Total")
                ws.cell(4, 11 + offset, "%")

                for x_i, x in enumerate(EXPEDITIONS):
                    r_idx = 5 + x_i
                    ws.cell(r_idx, 2 + offset, x)
                    x_gt = 0
                    for p_i, p in enumerate(PLANTS):
                        val_p = int(mtd_by_xpdc[x][p])
                        ws.cell(r_idx, 3 + p_i + offset, val_p)
                        x_gt += val_p
                    ws.cell(r_idx, 10 + offset, x_gt)
                    ws.cell(r_idx, 11 + offset, round(x_gt / mtd_grand_all, 6) if mtd_grand_all > 0 else 0.0)

                # Baris 10: TOTAL MTD
                ws.cell(10, 2 + offset, "TOTAL")
                for p_i, p in enumerate(PLANTS):
                    col_tot = sum(int(mtd_by_xpdc[x][p]) for x in EXPEDITIONS)
                    ws.cell(10, 3 + p_i + offset, col_tot)
                ws.cell(10, 10 + offset, mtd_grand_all)
                ws.cell(10, 11 + offset, 1.0 if mtd_grand_all > 0 else 0.0)

                # Baris 12: Header Tabel Harian
                ws.cell(12, 1 + offset, "TANGGAL")
                ws.cell(12, 2 + offset, "EKSPEDISI")
                for p_i, p in enumerate(PLANTS):
                    ws.cell(12, 3 + p_i + offset, p)
                ws.cell(12, 10 + offset, "Grand Total")
                ws.cell(12, 11 + offset, "%")

            # 2. Tulis Rincian Harian per Tanggal mulai Baris 14 pada Kedua Blok (H-1 dan H)
            if ws.max_row >= 14:
                ws.delete_rows(14, ws.max_row - 13)

            unique_dates = sorted(list({r["TANGGAL"] for r in all_rows}))
            curr_r = 14
            for d_str in unique_dates:
                d_dt = datetime.strptime(d_str, "%Y-%m-%d")
                d_sub = {r["XPDC"]: r for r in all_rows if r["TANGGAL"] == d_str}
                for x in EXPEDITIONS:
                    r_data = d_sub.get(x, {"TANGGAL": d_str, "XPDC": x, **{p: 0 for p in PLANTS}, "GRAND_TOTAL": 0, "PCT": 0.0})
                    for offset in [0, 12]:
                        c_tgl = ws.cell(curr_r, 1 + offset, d_dt)
                        c_tgl.number_format = "yyyy-mm-dd"
                        ws.cell(curr_r, 2 + offset, x)
                        for p_i, p in enumerate(PLANTS):
                            ws.cell(curr_r, 3 + p_i + offset, int(r_data.get(p, 0)))
                        ws.cell(curr_r, 10 + offset, int(r_data.get("GRAND_TOTAL", 0)))
                        ws.cell(curr_r, 11 + offset, float(r_data.get("PCT", 0.0)))
                    curr_r += 1
                curr_r += 1  # 1 baris pemisah antar tanggal sesuai format asli PER DAY

            wb.save(self.excel_path)
        wb.close()
        return True

    def get_all_target_bulanan(self):
        """Membaca seluruh target bulanan perencanaan dari Sheet TARGET_BULANAN atau target_bulanan.json."""
        targets = []
        json_path = os.path.join(os.path.dirname(os.path.abspath(self.excel_path)), "target_bulanan.json")
        
        # 1. Coba baca dari Excel
        try:
            wb = openpyxl.load_workbook(self.excel_path, read_only=True, data_only=True)
            if "TARGET_BULANAN" in wb.sheetnames:
                ws = wb["TARGET_BULANAN"]
                rows = list(ws.iter_rows(values_only=True))
                if len(rows) > 1:
                    by_period = {}
                    for r in rows[1:]:
                        if not r or len(r) < 4 or r[0] is None or r[1] is None:
                            continue
                        th = to_int(r[0])
                        bl = to_int(r[1])
                        xpdc = str(r[2] or "").strip().upper()
                        tgt = to_int(r[3], 0)
                        tot = to_int(r[4], 0) if len(r) > 4 else 0
                        upd = str(r[5] or "") if len(r) > 5 else ""
                        cat = str(r[6] or "") if len(r) > 6 else ""
                        if not th or not bl or xpdc not in EXPEDITIONS:
                            continue
                        k = (th, bl)
                        if k not in by_period:
                            by_period[k] = {
                                "TAHUN": th,
                                "BULAN": bl,
                                "TARGETS": {x: 0 for x in EXPEDITIONS},
                                "TOTAL_TARGET": tot,
                                "UPDATED_AT": upd,
                                "CATATAN": cat,
                            }
                        by_period[k]["TARGETS"][xpdc] = tgt
                    for (th, bl), item in by_period.items():
                        calc_tot = sum(item["TARGETS"].values())
                        if item["TOTAL_TARGET"] == 0:
                            item["TOTAL_TARGET"] = calc_tot
                        targets.append(item)
            wb.close()
        except Exception:
            pass

        # 2. Coba sinkronisasi / fallback dengan file JSON jika ada
        if os.path.exists(json_path):
            try:
                import json
                with open(json_path, "r", encoding="utf-8") as jf:
                    j_data = json.load(jf)
                    if isinstance(j_data, list):
                        ex_keys = {(t["TAHUN"], t["BULAN"]) for t in targets}
                        for jd in j_data:
                            k = (to_int(jd.get("TAHUN")), to_int(jd.get("BULAN")))
                            if k not in ex_keys:
                                targets.append(jd)
            except Exception:
                pass

        targets.sort(key=lambda t: (t["TAHUN"], t["BULAN"]), reverse=True)
        return targets

    def get_target_bulanan(self, tahun=2026, bulan=10):
        """Mengambil target bulanan spesifik untuk tahun dan bulan tertentu."""
        th = to_int(tahun)
        bl = to_int(bulan)
        all_tgt = self.get_all_target_bulanan()
        for t in all_tgt:
            if t["TAHUN"] == th and t["BULAN"] == bl:
                return t
        return None

    def save_target_bulanan(self, tahun, bulan, target_per_xpdc_dict, total_target=None, catatan=""):
        """Menyimpan Target Bulanan Perencanaan ke Excel Sheet TARGET_BULANAN dan cache target_bulanan.json."""
        th = to_int(tahun)
        bl = to_int(bulan)
        targets_clean = {x: to_int(target_per_xpdc_dict.get(x, 0)) for x in EXPEDITIONS}
        calc_tot = sum(targets_clean.values())
        tot = to_int(total_target, calc_tot)
        if tot == 0:
            tot = calc_tot
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 1. Update ke file JSON
        json_path = os.path.join(os.path.dirname(os.path.abspath(self.excel_path)), "target_bulanan.json")
        existing_json = []
        if os.path.exists(json_path):
            try:
                import json
                with open(json_path, "r", encoding="utf-8") as jf:
                    existing_json = json.load(jf)
            except Exception:
                existing_json = []
        existing_json = [r for r in existing_json if not (r.get("TAHUN") == th and r.get("BULAN") == bl)]
        new_entry = {
            "TAHUN": th,
            "BULAN": bl,
            "TARGETS": targets_clean,
            "TOTAL_TARGET": tot,
            "UPDATED_AT": now_str,
            "CATATAN": catatan,
        }
        existing_json.append(new_entry)
        try:
            import json
            with open(json_path, "w", encoding="utf-8") as jf:
                json.dump(existing_json, jf, indent=2)
        except Exception as e:
            print("Warning saving target_bulanan.json:", e)

        # 2. Update ke Sheet TARGET_BULANAN di DDC.xlsx
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            if "TARGET_BULANAN" not in wb.sheetnames:
                ws = wb.create_sheet("TARGET_BULANAN")
                ws.append(["TAHUN", "BULAN", "XPDC", "TARGET_BULAN", "TOTAL_TARGET", "UPDATED_AT", "CATATAN"])
            else:
                ws = wb["TARGET_BULANAN"]

            rows_keep = []
            for r in ws.iter_rows(values_only=True):
                if not r or r[0] == "TAHUN":
                    continue
                if to_int(r[0]) == th and to_int(r[1]) == bl:
                    continue
                rows_keep.append(list(r))

            for x in EXPEDITIONS:
                rows_keep.append([th, bl, x, targets_clean[x], tot, now_str, catatan])

            ws.delete_rows(1, max(1, ws.max_row))
            ws.append(["TAHUN", "BULAN", "XPDC", "TARGET_BULAN", "TOTAL_TARGET", "UPDATED_AT", "CATATAN"])
            for row in rows_keep:
                ws.append(row)

            wb.save(self.excel_path)
            wb.close()
        except PermissionError:
            pass
        except Exception as e:
            print("Warning saving to Excel TARGET_BULANAN:", e)

        return new_entry
