"""Kiểm tra Bước 3 → 4 → 5 và phương án dự phòng bằng "AI GIẢ LẬP" – không cần key, miễn phí.

AI giả lập trả về những câu trả lời soạn sẵn (code sai, code đúng, câu trả lời thẳng…) để kiểm tra:
  - vòng chạy code → gửi lỗi → AI sửa
  - các lớp bảo vệ khi chạy code AI (chặn đọc/ghi file, giới hạn thời gian, không hỏng bảng gốc)
  - dự phòng 1 (viết lại từ đầu) và dự phòng 2 (đọc thẳng dữ liệu)
  - chuẩn hóa đáp án trên kết quả thật của cả chuỗi

Chạy: python tools/test_ai_gia_lap.py      (cần chạy tools/tao_bo_test.py trước nếu chưa có test_data/)
"""
import os
import sys
import time

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, GOC)

from smartexcel import chay_code as B4  # noqa: E402
from smartexcel.chuan_hoa import chuan_hoa  # noqa: E402
from smartexcel.doc_bang import doc_tat_ca  # noqa: E402
from smartexcel.du_phong import tra_loi_cau_hoi  # noqa: E402
from smartexcel.tom_tat import tom_tat_tat_ca  # noqa: E402


class AIGia:
    """Trả lần lượt các câu trả lời soạn sẵn; ghi lại các tin nhắn đã nhận."""

    def __init__(self, cac_tra_loi):
        self.ds, self.nhan = list(cac_tra_loi), []

    def hoi(self, messages, max_tokens=0):
        self.nhan.append(messages)
        return self.ds.pop(0) if self.ds else ""


def code(s):
    return f"Đây là code:\n```python\n{s}\n```"


KQ = []


def kiem_tra(dat, mo_ta):
    KQ.append(dat)
    print(f"  {'✅' if dat else '❌'} {mo_ta}")


bang = doc_tat_ca(os.path.join(GOC, "test_data", "xlsx_files"))
tom_tat_tat_ca(bang)
dfs = {b.ten: b.df for b in bang}
SAI = code("answer = dfs['khong_co']")

# ------------------------------------------------------------------
print("=" * 70, "\nBƯỚC 4: CHẠY CODE + SỬA LỖI")
ai = AIGia([code('df = dfs["chi_nhanh.xlsx::DoanhSo"]\nanswer = df.loc[df["Chi nhánh"]=="Đà Nẵng", "Quý 2"].iloc[0]'),
            code('df = dfs["chi_nhanh.xlsx::DoanhSo"]\nanswer = df.loc[df["Chi nhánh"]=="Đà Nẵng", "Doanh số - Quý 2"].iloc[0]')])
kq = B4.giai_cau_hoi(ai, "A", "Doanh số quý 2 của chi nhánh Đà Nẵng là bao nhiêu?", bang, dfs)[0]
kiem_tra(kq == 950000 and len(ai.nhan) == 2, f"sai tên cột → AI sửa ở lần 2 → {kq}")
kiem_tra("Tên bảng/cột đúng" in ai.nhan[1][-1]["content"], "tin nhắn sửa lỗi có kèm danh sách cột đúng")

ai = AIGia([code('df = dfs["diem_thi.xlsx::HK1"]\nanswer = df.loc[df["Họ tên"]=="ngô minh châu", "Văn"]'),
            code('df = dfs["diem_thi.xlsx::HK1"]\nanswer = df.loc[df["Họ tên"]=="Ngô Minh Châu", "Văn"].iloc[0]')])
kq = B4.giai_cau_hoi(ai, "B", "Điểm Văn của Ngô Minh Châu là bao nhiêu?", bang, dfs)[0]
kiem_tra(kq == 9.0, f"lọc sai chữ thường → kết quả rỗng bị coi là lỗi → AI sửa → {kq}")

ai = AIGia([code("answer = 1/0")] * 3)
kq = B4.giai_cau_hoi(ai, "C", "Lớp 10A có bao nhiêu học sinh?", bang, dfs)[0]
kiem_tra(kq is None and len(ai.nhan) == 3, f"sai cả 3 lần → bỏ cuộc (None), gọi AI {len(ai.nhan)} lần")

for c, ten in [('import os\nos.remove("x")\nanswer=1', "import os"),
               ('answer = open("/input/queries.json").read()', "mở file"), ("exit()\nanswer=1", "thoát chương trình")]:
    try:
        B4.chay_code(c, dfs)
        kiem_tra(False, f"chặn {ten}")
    except B4.LoiChayCode:
        kiem_tra(True, f"chặn {ten}")

cu = B4.THOI_GIAN_TOI_DA
B4.THOI_GIAN_TOI_DA = 2
t = time.time()
try:
    B4.chay_code("while True: pass", dfs)
    kiem_tra(False, "dừng vòng lặp vô hạn")
except B4.LoiChayCode:
    kiem_tra(time.time() - t < 5, f"dừng vòng lặp vô hạn sau {time.time() - t:.1f}s")
B4.THOI_GIAN_TOI_DA = cu

B4.chay_code('df = dfs["diem_thi.xlsx::HK1"]\ndf.drop(index=[0,1,2], inplace=True)\ndf["Toán"] = 0\nanswer = 1', dfs)
kiem_tra(len(dfs["diem_thi.xlsx::HK1"]) == 7 and dfs["diem_thi.xlsx::HK1"]["Toán"].iloc[0] == 8.5,
         "code AI xóa/sửa bảng → bảng gốc không bị ảnh hưởng")

# ------------------------------------------------------------------
print("=" * 70, "\nCẢ CHUỖI BƯỚC 3 → 4 → 5 (code đúng, dữ liệu khó)")
CODE_DUNG = {
    "Tổng trọng lượng các đơn gửi đi Hà Nội là bao nhiêu kg?":
        ('df = dfs["van_don.xlsx::VanDon"]\nanswer = df.loc[df["Tỉnh nhận"]=="Hà Nội", "Trọng lượng (kg)"].sum()', "30.5"),
    "Có bao nhiêu đơn được gửi trong tháng 10/2026?":
        ('d = dfs["van_don.xlsx::VanDon"]["Ngày gửi"]\nanswer = ((d.dt.month==10)&(d.dt.year==2026)).sum()', "9"),
    "Liệt kê các mã vận đơn có mã bưu chính 070000.":
        ('df = dfs["van_don.xlsx::VanDon"]\nanswer = df.loc[df["Mã bưu chính"]=="070000", "Mã vận đơn"]', "VD0007, VD0012"),
    "Tổng chi phí khám của bệnh nhân Trần Văn Bảo là bao nhiêu?":
        ('bn = dfs["benh_nhan.xlsx::BenhNhan"]; kb = dfs["benh_nhan.xlsx::KhamBenh"]\n'
         'ma = bn.loc[bn["Họ tên"]=="Trần Văn Bảo","Mã BN"].iloc[0]\nanswer = kb.loc[kb["Mã BN"]==ma,"Chi phí"].sum()', "800000"),
    "Có bao nhiêu lượt khám có tỷ lệ BHYT 80%?":
        ('df = dfs["benh_nhan.xlsx::KhamBenh"]\nanswer = (df["Tỷ lệ BHYT (%)"]==80).sum()', "4"),
    "Mã số thuế của công ty là gì?":
        ('answer = dfs["cong_ty.xlsx::ThongTin"]["Mã số thuế"]', "0312345678"),
}
for cau, (c, dung) in CODE_DUNG.items():
    kq = chuan_hoa(tra_loi_cau_hoi(AIGia([code(c)]), "x", cau, bang, dfs))
    kiem_tra(kq == dung, f"{cau} → {kq!r} (đúng: {dung!r})")

# ------------------------------------------------------------------
print("=" * 70, "\nPHƯƠNG ÁN DỰ PHÒNG")
CAU = "Tổng chi phí khám của bệnh nhân Trần Văn Bảo là bao nhiêu?"
C_DUNG = CODE_DUNG[CAU][0]

ai = AIGia([SAI, SAI, SAI, code(C_DUNG)])
kq = chuan_hoa(tra_loi_cau_hoi(ai, "D1", CAU, bang, dfs))
kiem_tra(kq == "800000" and len(ai.nhan) == 4, f"cách chính sai 3 lần → dự phòng 1 đúng → {kq!r}")
kiem_tra(len(ai.nhan[3]) == 2 and "lần thử trước đã thất bại" in ai.nhan[3][-1]["content"],
         "dự phòng 1 là cuộc trò chuyện mới, có báo lỗi lần trước")

ai = AIGia([SAI] * 5 + ["Đáp án: 800.000 đồng."])
kq = chuan_hoa(tra_loi_cau_hoi(ai, "D2", CAU, bang, dfs))
kiem_tra(kq == "800000" and len(ai.nhan) == 6, f"dự phòng 1 cũng sai → dự phòng 2 trả lời thẳng → {kq!r}")

ai = AIGia([SAI] * 5 + ["400"])
tra_loi_cau_hoi(ai, "D3", "Danh mục Hóa đơn có bao nhiêu giao dịch Chi?", bang, dfs)
noi_dung = ai.nhan[5][-1]["content"]
kiem_tra("40/3000 dòng" in noi_dung and len(noi_dung) < 5000,
         f"bảng 3000 dòng: dự phòng 2 chỉ gửi dòng liên quan ({len(noi_dung)} ký tự)")

ai = AIGia([SAI] * 5 + [""])
kiem_tra(tra_loi_cau_hoi(ai, "D4", CAU, bang, dfs) is None, "mọi cách thất bại → để trống, chương trình không lỗi")
kiem_tra(tra_loi_cau_hoi(None, "D5", CAU, bang, dfs) is None, "không có AI → không chạy dự phòng")

# ------------------------------------------------------------------
print("=" * 70)
print(f"AI GIẢ LẬP: đạt {sum(KQ)}/{len(KQ)} kiểm tra")
sys.exit(0 if all(KQ) else 1)
