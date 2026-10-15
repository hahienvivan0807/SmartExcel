"""Xem Bước 3 làm gì với từng câu hỏi: chọn bảng nào, gợi ý gì, prompt dài bao nhiêu.

Chạy:
    python tools/xem_prompt.py                 # tóm tắt mọi câu
    python tools/xem_prompt.py --cau 3         # in nguyên prompt của câu 3
    python tools/xem_prompt.py --key KEY       # gọi AI thật, in code AI viết
"""
import argparse
import json
import os
import sys

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, GOC)

from smartexcel.ai import AI  # noqa: E402
from smartexcel.cau_hoi import doc_cau_hoi  # noqa: E402
from smartexcel.doc_bang import doc_tat_ca  # noqa: E402
from smartexcel.hoi_dap import chon_bang, goi_y_gia_tri, lay_code, tao_messages  # noqa: E402
from smartexcel.tom_tat import tom_tat_tat_ca  # noqa: E402

TEST = os.path.join(GOC, "test_data")
p = argparse.ArgumentParser()
p.add_argument("--xlsx-dir", default=os.path.join(TEST, "xlsx_files"))
p.add_argument("--queries", default=os.path.join(TEST, "queries.json"))
p.add_argument("--cau", help="id câu muốn in nguyên prompt")
p.add_argument("--key", default=os.environ.get("OPENROUTER_API_KEY", ""))
p.add_argument("--model", default="qwen/qwen3.7-flash")
a = p.parse_args()

ai = AI(a.key, a.model, "https://openrouter.ai/api/v1") if a.key else None
cac_bang = doc_tat_ca(a.xlsx_dir, ai)
tom_tat_tat_ca(cac_bang)

for qid, cau_hoi in doc_cau_hoi(a.queries):
    if a.cau and qid != a.cau:
        continue
    bang_chon = chon_bang(cau_hoi, cac_bang)
    goi_y = goi_y_gia_tri(cau_hoi, bang_chon)
    messages = tao_messages(cau_hoi, bang_chon, goi_y)
    do_dai = sum(len(m["content"]) for m in messages)

    print("=" * 70)
    print(f"[{qid}] {cau_hoi}")
    print(f"  Bảng chọn: {', '.join(b.ten for b in bang_chon)}")
    for g in goi_y:
        print(f"  Gợi ý: {g}")
    print(f"  Prompt: {do_dai} ký tự")
    if a.cau:
        for m in messages:
            print(f"\n----- {m['role'].upper()} -----\n{m['content']}")
    if ai is not None:
        print("  Code AI viết:")
        print("    " + lay_code(ai.hoi(messages, max_tokens=500)).replace("\n", "\n    "))
