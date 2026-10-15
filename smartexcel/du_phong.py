"""Phương án dự phòng khi cách chính (AI viết code, sửa tối đa 3 lần) vẫn thất bại.

    Cách chính ── thất bại ──► Dự phòng 1: VIẾT LẠI TỪ ĐẦU, NHÌN RỘNG HƠN
                                 cuộc trò chuyện mới (không bị "kẹt" theo code sai cũ),
                                 gửi thêm bảng (lỡ lúc đầu chọn thiếu bảng), báo lỗi lần trước
                                 → AI viết code → chạy (sửa tối đa 1 lần)
                 ── thất bại ──► Dự phòng 2: ĐỌC THẲNG DỮ LIỆU
                                 lọc ≤ 40 dòng liên quan nhất, gửi nguyên dòng cho AI
                                 → AI trả lời thẳng (không viết code)
                 ── thất bại ──► để trống

Để trống chắc chắn sai; trả lời thẳng tuy kém chắc chắn hơn chạy code nhưng còn có cơ hội đúng.
Dự phòng chỉ chạy với câu thất bại nên hầu như không tốn thêm token.
"""
import pandas as pd

from .chay_code import giai_cau_hoi, thu_va_sua
from .hoi_dap import chon_bang, goi_y_gia_tri, lay_code, tao_messages, tim_gia_tri, xep_hang_bang
from .log import log

KY_TU_TOI_DA_DP1 = 8000    # tổng độ dài tóm tắt gửi ở dự phòng 1
DONG_TOI_DA_DP2 = 40       # số dòng dữ liệu tối đa mỗi bảng ở dự phòng 2
DO_DAI_O_DP2 = 60


def tra_loi_cau_hoi(ai, qid, cau_hoi, cac_bang, dfs):
    """Trả về answer thô (chưa chuẩn hóa), hoặc None nếu mọi cách đều thất bại."""
    answer, loi, bang_chon = giai_cau_hoi(ai, qid, cau_hoi, cac_bang, dfs)
    if answer is not None or ai is None:
        return answer

    log(f"[câu {qid}] → dự phòng 1: viết lại từ đầu, nhìn rộng hơn")
    answer = viet_lai_tu_dau(ai, qid, cau_hoi, cac_bang, dfs, loi)
    if answer is not None:
        return answer

    log(f"[câu {qid}] → dự phòng 2: đọc thẳng dữ liệu")
    answer = doc_thang(ai, qid, cau_hoi, cac_bang)
    if answer is None:
        log(f"[câu {qid}] ❌ mọi cách đều thất bại → để trống")
    return answer


# ============================================================
# Dự phòng 1: viết lại từ đầu, nhìn rộng hơn
# ============================================================

def chon_bang_rong(cau_hoi, cac_bang, ky_tu_toi_da=KY_TU_TOI_DA_DP1):
    """Lấy MỌI bảng có liên quan (điểm > 0, không áp ngưỡng 40% như lần đầu) theo thứ hạng,
    cho tới khi đủ ky_tu_toi_da ký tự tóm tắt (ít nhất 1 bảng). Không bảng nào khớp → xét tất cả."""
    xep_hang = xep_hang_bang(cau_hoi, cac_bang)
    co_diem = [(d, b) for d, b in xep_hang if d > 0] or xep_hang
    chon, tong = [], 0
    for _, b in co_diem:
        if chon and tong + len(b.tom_tat) > ky_tu_toi_da:
            break
        chon.append(b)
        tong += len(b.tom_tat)
    return chon


def viet_lai_tu_dau(ai, qid, cau_hoi, cac_bang, dfs, loi_cu):
    bang = chon_bang_rong(cau_hoi, cac_bang)
    messages = tao_messages(cau_hoi, bang, goi_y_gia_tri(cau_hoi, bang))
    dong_loi = (loi_cu or "").splitlines()[0] if loi_cu else "không rõ"
    messages[-1]["content"] += (
        f"\n\nLưu ý: lần thử trước đã thất bại (lỗi: {dong_loi}). "
        "Kiểm tra kỹ bảng nào thật sự chứa dữ liệu cần; có thể phải ghép nhiều bảng "
        "hoặc dùng cột khác. Viết lại code từ đầu.")
    code = lay_code(ai.hoi(messages, max_tokens=600))
    answer, _ = thu_va_sua(ai, qid, code, messages, dfs, [b.ten for b in bang],
                           so_lan_sua=1, nhan="dự phòng 1, ")
    return answer


# ============================================================
# Dự phòng 2: đọc thẳng dữ liệu
# ============================================================

def chon_dong_lien_quan(bang, gia_tri_tim_thay, toi_da=DONG_TOI_DA_DP2):
    """Chọn ≤ toi_da dòng của bảng để gửi thẳng cho AI. Trả về (DataFrame, ghi chú).

    - Bảng nhỏ → gửi hết.
    - Bảng lớn → ưu tiên dòng chứa nhiều giá trị câu hỏi nhắc tới nhất
      (vd "Kho Bình Dương … bị hoàn": dòng có cả "Bình Dương" và "Hoàn" lên đầu).
    - Không có giá trị nào để lọc → lấy các dòng đầu (đành chấp nhận thiếu).
    """
    df = bang.df
    if len(df) <= toi_da:
        return df, f"toàn bộ {len(df)} dòng"
    can_tim = [(col, v) for ten, col, v in gia_tri_tim_thay if ten == bang.ten]
    if can_tim:
        diem = sum((df[col].astype(str) == v).astype(int) for col, v in can_tim)
        co = diem[diem > 0]
        if len(co):
            # dòng khớp nhiều giá trị nhất lên trước, lấy ≤ toi_da dòng, rồi xếp lại theo thứ tự gốc
            chon = sorted(co.sort_values(ascending=False, kind="stable").index[:toi_da])
            ten_gt = ", ".join(f'"{v}"' for _, v in can_tim)
            return df.loc[chon], f"{len(chon)}/{len(df)} dòng có chứa {ten_gt}"
    return df.head(toi_da), f"{toi_da}/{len(df)} dòng đầu (bảng quá dài, chỉ là một phần)"


PROMPT_DOC_THANG = """Trả lời câu hỏi dựa trên dữ liệu bảng bên dưới.
Chỉ ghi ĐÁP ÁN CUỐI CÙNG: 1 số hoặc 1 cụm từ ngắn; nhiều kết quả thì nối bằng ", ".
Không giải thích, không viết code."""


def _bang_csv(df):
    def o(x):
        if x is None or (not isinstance(x, str) and pd.isna(x)):
            return ""
        if isinstance(x, pd.Timestamp):
            return str(x.date()) if not (x.hour or x.minute) else str(x)
        s = str(x)
        return s if len(s) <= DO_DAI_O_DP2 else s[:DO_DAI_O_DP2] + "…"
    return df.map(o).to_csv(index=False).strip()


def tao_messages_doc_thang(cau_hoi, cac_bang):
    tim_thay = tim_gia_tri(cau_hoi, cac_bang)
    phan = []
    for b in cac_bang:
        dong, ghi_chu = chon_dong_lien_quan(b, tim_thay)
        phan.append(f'Bảng "{b.ten}" – {ghi_chu}:\n{_bang_csv(dong)}')
    return [
        {"role": "system", "content": PROMPT_DOC_THANG},
        {"role": "user", "content": "\n\n".join(phan) + f"\n\nCâu hỏi: {cau_hoi}"},
    ]


def doc_thang(ai, qid, cau_hoi, cac_bang):
    bang = chon_bang(cau_hoi, cac_bang)
    tra_loi = ai.hoi(tao_messages_doc_thang(cau_hoi, bang), max_tokens=150)
    dong = [d for d in tra_loi.strip().splitlines() if d.strip()]
    if not dong:
        return None
    log(f"[câu {qid}] ✅ dự phòng 2 (đọc thẳng): {dong[0]!r}")
    return dong[0]
