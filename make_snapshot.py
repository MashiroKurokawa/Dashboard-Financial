"""Ekspor Excel -> js/snapshot.js (cadangan saat index.html dibuka tanpa server).
Jalankan: python tools/make_snapshot.py   (butuh: pip install openpyxl)"""
import json, pathlib
from openpyxl import load_workbook
root = pathlib.Path(__file__).resolve().parent.parent
wb = load_workbook(root/"data/kelompok_3_warung_makan_mie_sederhana.xlsx", data_only=True)
out = {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb}
(root/"js/snapshot.js").write_text("window.SNAPSHOT="+json.dumps(out, ensure_ascii=False)+";", encoding="utf-8")
print("OK:", list(out))