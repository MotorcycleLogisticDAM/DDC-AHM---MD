import os
import sys
import shutil

# Tambahkan direktori root proyek ke path Python
ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

# Di Vercel Serverless, hanya folder /tmp yang writable
if os.environ.get("VERCEL"):
    tmp_excel = "/tmp/DDC.xlsx"
    src_excel = os.path.join(ROOT_DIR, "DDC.xlsx")
    if not os.path.exists(tmp_excel) and os.path.exists(src_excel):
        shutil.copy(src_excel, tmp_excel)
    os.environ["DDC_EXCEL_PATH"] = tmp_excel

from app import app
