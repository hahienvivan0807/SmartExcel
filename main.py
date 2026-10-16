"""SmartExcel – điểm vào chương trình (ENTRYPOINT của Docker).

Luồng chạy:
    đọc câu hỏi → đọc & làm sạch bảng (Bước 1) → tóm tắt (Bước 2)
    → với mỗi câu: AI viết code (3) → chạy code (4) → chuẩn hóa (5)
    → ghi CSV
"""
import argparse
import sys

from smartexcel import cau_hinh
from smartexcel.ai import AI
from smartexcel.cau_hoi import doc_cau_hoi, ghi_ket_qua
from smartexcel.doc_bang import doc_tat_ca
from smartexcel.chuan_hoa import chuan_hoa
from smartexcel.du_phong import tra_loi_cau_hoi
from smartexcel.log import log
from smartexcel.tom_tat import tom_tat_tat_ca


def nhan_tham_so():
    p = argparse.ArgumentParser(description="SmartExcel - trả lời câu hỏi về bảng tính")
    p.add_argument("--xlsx-dir", required=True)
    p.add_argument("--queries", required=True)
    p.add_argument("--output", required=True)
    # Máy chấm truyền key qua --openrouter-key. Chạy trên máy mình thì bỏ trống → lấy từ
    # biến môi trường OPENROUTER_API_KEY hoặc file .env (xem smartexcel/cau_hinh.py).
    p.add_argument("--openrouter-key", default="")
    p.add_argument("--openrouter-model", default=cau_hinh.MODEL)
    p.add_argument("--openrouter-base-url", default=cau_hinh.BASE_URL)
    return p.parse_args()


def tra_loi(qid, cau_hoi, cac_bang, dfs, ai):
    # Bước 3 + 4: AI viết code → chạy → lỗi thì gửi lại cho AI sửa → vẫn sai thì dự phòng
    answer = tra_loi_cau_hoi(ai, qid, cau_hoi, cac_bang, dfs)
    if answer is None:
        return ""
    # Bước 5: chuẩn hóa thành 1 chuỗi ngắn
    dap_an = chuan_hoa(answer)
    log(f"[câu {qid}] đáp án: {dap_an!r}")
    return dap_an


def main():
    a = nhan_tham_so()
    a.openrouter_key = a.openrouter_key or cau_hinh.lay_key()
    cac_cau_hoi = doc_cau_hoi(a.queries)
    log(f"Đã đọc {len(cac_cau_hoi)} câu hỏi")

    # Đáp án mặc định là rỗng → dù có lỗi gì, CSV vẫn đủ mọi id
    dap_an = {qid: "" for qid, _ in cac_cau_hoi}
    try:
        ai = AI(a.openrouter_key, a.openrouter_model, a.openrouter_base_url) if a.openrouter_key else None

        # Bước 1: đọc & làm sạch
        cac_bang = doc_tat_ca(a.xlsx_dir, ai)
        log(f"Đã đọc {len(cac_bang)} bảng")

        # Bước 2: tóm tắt
        tom_tat_tat_ca(cac_bang)
        dfs = {b.ten: b.df for b in cac_bang}

        for qid, cau_hoi in cac_cau_hoi:
            try:
                dap_an[qid] = tra_loi(qid, cau_hoi, cac_bang, dfs, ai)
            except Exception as e:
                log(f"[câu {qid}] lỗi: {e}")
    except Exception as e:
        log(f"Lỗi chung: {e}")
    finally:
        ghi_ket_qua(a.output, cac_cau_hoi, dap_an)
        log(f"Đã ghi kết quả vào {a.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
