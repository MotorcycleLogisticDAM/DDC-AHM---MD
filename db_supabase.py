import os
import sys
import json
import time
import urllib.request
import urllib.error
from datetime import datetime

PLANTS = ["FM1", "FM2", "FM6", "FM6A", "FM7", "FM8", "FM9"]
EXPEDITIONS = ["TUNAS MUDA", "RJTM", "LBK", "WSS", "PUNINAR"]

DEFAULT_SUPABASE_URL = "https://jpquejyuuhwzudwodbth.supabase.co"
DEFAULT_SUPABASE_KEY = "sb_publishable_XXS_kSxECzJIlD_dUjJflA_v4pAnc4n"

class SupabaseDatabase:
    """
    Database client untuk Supabase PostgreSQL REST API.
    Mendukung operasi CRUD penuh untuk DDC logistik dengan performa tinggi,
    in-memory smart caching, dan kompatibilitas 100% pada lingkungan Vercel Serverless & Multi-device.
    """
    def __init__(self, excel_path=None, supabase_url=None, supabase_key=None):
        self.excel_path = excel_path or os.environ.get("DDC_EXCEL_PATH", r"J:\My Drive\DDC PROJECT\DDC.xlsx")
        self.url = (supabase_url or os.environ.get("SUPABASE_URL") or DEFAULT_SUPABASE_URL).rstrip("/")
        self.key = supabase_key or os.environ.get("SUPABASE_KEY") or DEFAULT_SUPABASE_KEY
        self._sl_cache = None
        self._sl_cache_time = 0
        self._do_cache = None
        self._do_cache_time = 0
        self._alloc_cache = None
        self._alloc_cache_time = 0

    def _headers(self, prefer=None):
        h = {
            "apikey": self.key,
            "Authorization": f"Bearer {self.key}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        if prefer:
            h["Prefer"] = prefer
        return h

    def _get(self, endpoint):
        req = urllib.request.Request(f"{self.url}/rest/v1/{endpoint}", headers=self._headers())
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))

    def _post(self, endpoint, data, on_conflict=None):
        url = f"{self.url}/rest/v1/{endpoint}"
        if on_conflict:
            url += f"?on_conflict={on_conflict}"
        body = json.dumps(data).encode("utf-8")
        prefer = "resolution=merge-duplicates,return=representation" if on_conflict else "return=representation"
        req = urllib.request.Request(url, data=body, headers=self._headers(prefer), method="POST")
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))

    def _delete(self, endpoint):
        req = urllib.request.Request(f"{self.url}/rest/v1/{endpoint}", headers=self._headers(), method="DELETE")
        with urllib.request.urlopen(req, timeout=30) as resp:
            return True

    def create_backup(self):
        """Metode kompatibilitas backup."""
        return True

    def get_sheet_names(self):
        """Mengembalikan daftar sheet virtual untuk UI."""
        return ["DO", "SL", "PER DAY", "ALOKASI_HARIAN", "TARGET_BULANAN"]

    # =========================================================
    # 1. DELIVERY ORDER (DO)
    # =========================================================
    def get_do_data(self):
        """Mengambil seluruh data DO dari tabel delivery_orders (dengan cache pintar)."""
        if self._do_cache is not None and (time.time() - self._do_cache_time < 30):
            return self._do_cache

        try:
            records = self._get("delivery_orders?select=*&order=tanggal.asc,plant.asc")
            formatted = []
            for r in records:
                formatted.append({
                    "TANGGAL_STR": str(r["tanggal"]),
                    "PLANT": str(r["plant"]),
                    "QTY": int(r["qty"]),
                    "QTY_AWAL": int(r.get("qty_awal", r["qty"])),
                    "ADJUSTMENT": int(r.get("adjustment", 0)),
                    "QTY_FINAL": int(r["qty"]),
                    "KETERANGAN": str(r.get("keterangan", "Normal")),
                    "BULAN": int(r.get("bulan", 10)),
                    "TAHUN": int(r.get("tahun", 2026)),
                })
            self._do_cache = formatted
            self._do_cache_time = time.time()
            return formatted
        except Exception as e:
            print(f"[Supabase] Error get_do_data: {e}")
            return self._do_cache or []

    def upsert_daily_do(self, tanggal_str, do_dict_per_plant, bulan_override=None):
        """Menyimpan atau meng-update data DO per plant untuk tanggal tertentu."""
        dt = datetime.strptime(str(tanggal_str)[:10], "%Y-%m-%d")
        bulan = int(bulan_override) if bulan_override else dt.month
        tahun = dt.year

        existing = {r["plant"]: r for r in self._get(f"delivery_orders?tanggal=eq.{tanggal_str}")}

        rows = []
        for p in PLANTS:
            qty_val = int(do_dict_per_plant.get(p, 0))
            if p in existing:
                prev = existing[p]
                q_awal = int(prev.get("qty_awal", prev["qty"]))
                adj = qty_val - q_awal
                ket = f"Update Qty ({adj:+d})" if adj != 0 else prev.get("keterangan", "Normal")
            else:
                q_awal = qty_val
                adj = 0
                ket = "Normal"

            rows.append({
                "tanggal": tanggal_str,
                "plant": p,
                "qty": qty_val,
                "bulan": bulan,
                "tahun": tahun,
                "qty_awal": q_awal,
                "adjustment": adj,
                "keterangan": ket
            })

        self._post("delivery_orders", rows, on_conflict="tanggal,plant")
        self._do_cache = None  # Invalidate cache
        return True

    def adjust_do(self, tanggal_str, plant, qty_final, keterangan):
        """Menyesuaikan DO plant tertentu."""
        existing = self._get(f"delivery_orders?tanggal=eq.{tanggal_str}&plant=eq.{plant}")
        if not existing:
            return self.upsert_daily_do(tanggal_str, {plant: qty_final})

        row = existing[0]
        q_awal = int(row.get("qty_awal", row["qty"]))
        adj = int(qty_final) - q_awal
        self._post("delivery_orders", [{
            "tanggal": tanggal_str,
            "plant": plant,
            "qty": int(qty_final),
            "bulan": int(row["bulan"]),
            "tahun": int(row["tahun"]),
            "qty_awal": q_awal,
            "adjustment": adj,
            "keterangan": str(keterangan)
        }], on_conflict="tanggal,plant")
        self._do_cache = None
        return True

    # =========================================================
    # 2. SHIPPING LIST (SL)
    # =========================================================
    def get_sl_data(self):
        """Mengambil data SL dari tabel shipping_lists dengan pagination dan cache efisien."""
        if self._sl_cache is not None and (time.time() - self._sl_cache_time < 120):
            return self._sl_cache

        try:
            records = []
            page_size = 1000
            offset = 0
            while True:
                batch = self._get(f"shipping_lists?select=*&order=id.asc&limit={page_size}&offset={offset}")
                if not batch:
                    break
                records.extend(batch)
                if len(batch) < page_size:
                    break
                offset += page_size

            formatted = []
            for r in records:
                m_val = int(r.get("bulan", 10))
                d_val = int(r.get("hari", 1))
                y_val = int(r.get("tahun", 2026))
                tgl_str = str(r["tanggal"])
                formatted.append({
                    "TANGGAL": tgl_str,
                    "TANGGAL_STR": tgl_str,
                    "NOPOL": str(r.get("nopol", "")),
                    "SOPIR": str(r.get("sopir", "")),
                    "XPDC": str(r.get("xpdc", "")),
                    "PLANT": str(r.get("plant", "")),
                    "GUDANG": str(r.get("gudang", "ALL")),
                    "QTY": int(r.get("qty", 0)),
                    "QTY (UNIT)": int(r.get("qty", 0)),
                    "MONTH": m_val,
                    "BULAN": m_val,
                    "DAY": d_val,
                    "HARI": d_val,
                    "YEAR": y_val,
                    "TAHUN": y_val,
                })
            self._sl_cache = formatted
            self._sl_cache_time = time.time()
            return formatted
        except Exception as e:
            print(f"[Supabase] Error get_sl_data: {e}")
            return self._sl_cache or []

    def append_sl_bulk(self, rows_list):
        """Menambahkan batch baris ke shipping_lists."""
        to_insert = []
        for r in rows_list:
            tgl = str(r.get("TANGGAL", "")).strip()[:10]
            try:
                dt = datetime.strptime(tgl, "%Y-%m-%d")
            except Exception:
                continue
            to_insert.append({
                "tanggal": tgl,
                "nopol": str(r.get("NOPOL", "")),
                "sopir": str(r.get("SOPIR", "")),
                "xpdc": str(r.get("XPDC", "")),
                "plant": str(r.get("PLANT", "")),
                "gudang": str(r.get("GUDANG", "ALL")),
                "qty": int(r.get("QTY", 0)),
                "bulan": int(dt.month),
                "hari": int(dt.day),
                "tahun": int(dt.year)
            })

        if not to_insert:
            return 0

        total = len(to_insert)
        for i in range(0, total, 200):
            batch = to_insert[i:i + 200]
            self._post("shipping_lists", batch)
        self._sl_cache = None  # Invalidate cache
        return total

    def append_sl_row(self, row_dict):
        return self.append_sl_bulk([row_dict])

    # =========================================================
    # 3. DAILY ALLOCATIONS (KUOTA HARIAN / PER DAY)
    # =========================================================
    def get_saved_daily_allocations(self):
        """Mengambil data alokasi harian dari tabel daily_allocations."""
        if self._alloc_cache is not None and (time.time() - self._alloc_cache_time < 30):
            return self._alloc_cache

        try:
            records = self._get("daily_allocations?select=*&order=tanggal.asc,xpdc.asc")
            formatted = []
            for r in records:
                formatted.append({
                    "TANGGAL": str(r["tanggal"]),
                    "XPDC": str(r["xpdc"]),
                    "FM1": int(r.get("fm1", 0)),
                    "FM2": int(r.get("fm2", 0)),
                    "FM6": int(r.get("fm6", 0)),
                    "FM6A": int(r.get("fm6a", 0)),
                    "FM7": int(r.get("fm7", 0)),
                    "FM8": int(r.get("fm8", 0)),
                    "FM9": int(r.get("fm9", 0)),
                    "GRAND_TOTAL": int(r.get("grand_total", 0)),
                    "PCT": float(r.get("pct", 0.0)),
                    "BULAN": int(r.get("bulan", 10)),
                    "TAHUN": int(r.get("tahun", 2026)),
                    "TOTAL_AWAL": int(r.get("total_awal", r.get("grand_total", 0))),
                    "ADJUSTMENT": int(r.get("adjustment", 0)),
                    "KETERANGAN": str(r.get("keterangan", "Normal")),
                    "DETAILS_JSON": str(r.get("details_json", "{}"))
                })
            self._alloc_cache = formatted
            self._alloc_cache_time = time.time()
            return formatted
        except Exception as e:
            print(f"[Supabase] Error get_saved_daily_allocations: {e}")
            return self._alloc_cache or []

    def save_daily_allocation(self, tanggal_str, alloc_matrix_dict):
        """Menyimpan alokasi kuota harian untuk 5 ekspedisi."""
        dt = datetime.strptime(str(tanggal_str)[:10], "%Y-%m-%d")

        total_all = sum(
            sum(int(alloc_matrix_dict.get(x, {}).get(p, 0)) for p in PLANTS)
            for x in EXPEDITIONS
        )

        rows = []
        for x in EXPEDITIONS:
            x_data = alloc_matrix_dict.get(x, {})
            gt = sum(int(x_data.get(p, 0)) for p in PLANTS)
            pct = round(gt / total_all, 6) if total_all > 0 else 0.0
            rows.append({
                "tanggal": tanggal_str,
                "xpdc": x,
                "fm1": int(x_data.get("FM1", 0)),
                "fm2": int(x_data.get("FM2", 0)),
                "fm6": int(x_data.get("FM6", 0)),
                "fm6a": int(x_data.get("FM6A", 0)),
                "fm7": int(x_data.get("FM7", 0)),
                "fm8": int(x_data.get("FM8", 0)),
                "fm9": int(x_data.get("FM9", 0)),
                "grand_total": gt,
                "pct": pct,
                "bulan": int(dt.month),
                "tahun": int(dt.year),
                "total_awal": int(x_data.get("TOTAL_AWAL", gt)),
                "adjustment": int(x_data.get("ADJUSTMENT", 0)),
                "keterangan": str(x_data.get("KETERANGAN", "Normal")),
                "details_json": str(x_data.get("DETAILS_JSON", "{}"))
            })

        self._post("daily_allocations", rows, on_conflict="tanggal,xpdc")
        self._alloc_cache = None
        return True

    def adjust_daily_allocation(self, tanggal_str, xpdc, plant, delta_val, keterangan):
        """Adjust alokasi manual."""
        clean_tgl = str(tanggal_str)[:10]
        existing = self._get(f"daily_allocations?tanggal=eq.{clean_tgl}&xpdc=eq.{xpdc}")
        if not existing:
            return False

        r = existing[0]
        p_col = plant.lower()
        curr_p = int(r.get(p_col, 0))
        new_p = max(0, curr_p + int(delta_val))
        r[p_col] = new_p

        gt = sum(int(r.get(p.lower(), 0)) for p in PLANTS)
        tot_awal = int(r.get("total_awal", gt))
        adj_total = gt - tot_awal

        r["grand_total"] = gt
        r["adjustment"] = adj_total
        r["keterangan"] = str(keterangan)

        self._post("daily_allocations", [r], on_conflict="tanggal,xpdc")
        self._alloc_cache = None
        return True

    # =========================================================
    # 4. TARGET BULANAN
    # =========================================================
    def get_all_target_bulanan(self):
        """Mengambil seluruh target bulanan."""
        try:
            records = self._get("monthly_targets?select=*&order=tahun.desc,bulan.desc")
            grouped = {}
            for r in records:
                k = (int(r["bulan"]), int(r["tahun"]))
                if k not in grouped:
                    grouped[k] = {
                        "BULAN": int(r["bulan"]),
                        "TAHUN": int(r["tahun"]),
                        "TARGETS": {},
                        "TOTAL_TARGET": 0,
                        "UPDATED_AT": str(r.get("created_at", "")),
                        "CATATAN": f"Perencanaan Target {r['bulan']}/{r['tahun']}"
                    }
                grouped[k]["TARGETS"][r["xpdc"]] = int(r.get("target_unit", 0))
                grouped[k]["TOTAL_TARGET"] += int(r.get("target_unit", 0))
            return list(grouped.values())
        except Exception as e:
            print(f"[Supabase] Error get_all_target_bulanan: {e}")
            return []

    def get_target_bulanan(self, bulan, tahun):
        all_t = self.get_all_target_bulanan()
        for t in all_t:
            if t["BULAN"] == int(bulan) and t["TAHUN"] == int(tahun):
                return t
        return None

    def save_target_bulanan(self, bulan, tahun, targets_dict, total_target, catatan=""):
        rows = []
        tot = int(total_target) if total_target > 0 else sum(int(v) for v in targets_dict.values())
        for xpdc in EXPEDITIONS:
            unit_val = int(targets_dict.get(xpdc, 0))
            pct = round(unit_val / tot * 100.0, 4) if tot > 0 else 0.0
            rows.append({
                "bulan": int(bulan),
                "tahun": int(tahun),
                "xpdc": xpdc,
                "target_pct": pct,
                "target_unit": unit_val
            })
        self._post("monthly_targets", rows, on_conflict="bulan,tahun,xpdc")
        return True

    # =========================================================
    # 5. METODE DUMMY / HELPER UNTUK SHEET PREVIEW
    # =========================================================
    def read_any_sheet(self, sheet_name, max_rows=200):
        """Membuat representasi tabular untuk tab Database Viewer."""
        if sheet_name == "DO":
            rows = self.get_do_data()
            headers = ["TANGGAL", "PLANT", "QTY", "BULAN", "QTY_AWAL", "ADJUSTMENT", "KETERANGAN"]
            data = [[r["TANGGAL_STR"], r["PLANT"], r["QTY"], r["BULAN"], r["QTY_AWAL"], r["ADJUSTMENT"], r["KETERANGAN"]] for r in rows[:max_rows]]
            return headers, data
        elif sheet_name == "SL":
            rows = self.get_sl_data()
            headers = ["TANGGAL", "NOPOL", "SOPIR", "XPDC", "PLANT", "GUDANG", "QTY", "BULAN", "HARI", "TAHUN"]
            data = [[r["TANGGAL"], r["NOPOL"], r["SOPIR"], r["XPDC"], r["PLANT"], r["GUDANG"], r["QTY"], r["BULAN"], r["HARI"], r["TAHUN"]] for r in rows[:max_rows]]
            return headers, data
        elif sheet_name in ["PER DAY", "ALOKASI_HARIAN"]:
            rows = self.get_saved_daily_allocations()
            headers = ["TANGGAL", "XPDC", "FM1", "FM2", "FM6", "FM6A", "FM7", "FM8", "FM9", "GRAND_TOTAL", "PCT", "BULAN", "TAHUN"]
            data = [[r["TANGGAL"], r["XPDC"], r["FM1"], r["FM2"], r["FM6"], r["FM6A"], r["FM7"], r["FM8"], r["FM9"], r["GRAND_TOTAL"], r["PCT"], r["BULAN"], r["TAHUN"]] for r in rows[:max_rows]]
            return headers, data
        return [], []

    def add_or_replace_sheet(self, sheet_name, headers, rows):
        return True
