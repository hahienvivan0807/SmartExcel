"""Đọc file câu hỏi và ghi file kết quả CSV."""
import csv
import json
import os

# Đề chưa nói rõ tên trường trong queries.json → chấp nhận nhiều kiểu tên
KHOA_ID = ("id", "qid", "query_id", "question_id")
KHOA_CAU_HOI = ("question", "query", "q", "text", "content", "cau_hoi")


def _lay(d, cac_khoa):
    for k in cac_khoa:
        if k in d:
            return d[k]
    return None


def doc_cau_hoi(duong_dan):
    """Trả về danh sách [(id, câu hỏi), ...] theo đúng thứ tự trong file."""
    with open(duong_dan, encoding="utf-8-sig") as f:
        data = json.load(f)

    # Dạng {"queries": [...]} → lấy danh sách bên trong
    if isinstance(data, dict):
        ds = next((v for v in data.values() if isinstance(v, list)), None)
        if ds is not None:
            data = ds
        else:
            # Dạng {"1": "câu hỏi", "2": "..."}
            return [(str(k), str(v)) for k, v in data.items()]

    ket_qua = []
    for i, muc in enumerate(data, start=1):
        if isinstance(muc, dict):
            qid = _lay(muc, KHOA_ID)
            cau = _lay(muc, KHOA_CAU_HOI)
            ket_qua.append((str(qid if qid is not None else i), str(cau or "")))
        else:
            ket_qua.append((str(i), str(muc)))
    return ket_qua


def ghi_ket_qua(duong_dan, cac_cau_hoi, dap_an):
    """Ghi CSV 2 cột id,answer – mỗi id đúng 1 lần, theo thứ tự câu hỏi."""
    thu_muc = os.path.dirname(duong_dan)
    if thu_muc:
        os.makedirs(thu_muc, exist_ok=True)
    da_ghi = set()
    with open(duong_dan, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["id", "answer"])
        for qid, _ in cac_cau_hoi:
            if qid in da_ghi:
                continue
            da_ghi.add(qid)
            w.writerow([qid, dap_an.get(qid, "")])
