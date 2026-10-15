"""Xem các bảng sau Bước 1 (làm sạch) và Bước 2 (tóm tắt) – để tự kiểm tra bằng mắt.

Chạy: python tools/xem_bang.py [thư_mục_excel] [--key KEY_CUA_BAN]
"""
import argparse
import os
import sys

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, GOC)

from smartexcel.ai import AI  # noqa: E402
from smartexcel.doc_bang import doc_tat_ca  # noqa: E402
from smartexcel.tom_tat import tom_tat_tat_ca  # noqa: E402

p = argparse.ArgumentParser()
p.add_argument("thu_muc", nargs="?", default=os.path.join(GOC, "test_data", "xlsx_files"))
p.add_argument("--key", default=os.environ.get("OPENROUTER_API_KEY", ""))
p.add_argument("--model", default="qwen/qwen3.7-flash")
a = p.parse_args()

ai = AI(a.key, a.model, "https://openrouter.ai/api/v1") if a.key else None
cac_bang = doc_tat_ca(a.thu_muc, ai)
tom_tat_tat_ca(cac_bang)

for bang in cac_bang:
    print("=" * 70)
    print(f"{bang.ten}  |  hướng: {bang.huong}  |  {len(bang.df)} dòng × {bang.df.shape[1]} cột")
    print("-" * 30, "BẢNG ĐÃ LÀM SẠCH (Bước 1)")
    print(bang.df.to_string(max_rows=10))
    print("-" * 30, f"TÓM TẮT GỬI AI (Bước 2) – {len(bang.tom_tat)} ký tự")
    print(bang.tom_tat)
    print("-" * 30, f"TỪ KHÓA ({len(bang.tu_khoa)})")
    print(" ".join(sorted(bang.tu_khoa)))
