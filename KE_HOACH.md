# SmartExcel – Kế hoạch giải Bảng C (AI 2026)

## Bài toán
- Đầu vào: thư mục nhiều file Excel + file `queries.json` chứa câu hỏi.
- Đầu ra: file CSV 2 cột `id,answer`, mỗi id đúng 1 lần.
- AI dùng: model do máy chấm đưa (`qwen/qwen3.7-flash` qua OpenRouter). Không ghi cứng key/model.

## Ý tưởng cốt lõi
> AI hiểu câu hỏi và viết code. Pandas chạy code để tính đáp án chính xác.

## Quy trình 5 bước

| Bước | Việc | Dùng AI? | Làm mấy lần |
|---|---|---|---|
| 1 | Đọc & làm sạch bảng (bỏ dòng trống, tìm dòng tiêu đề, đoán ngang/dọc/chéo, xoay bảng dọc, sửa số, bỏ dòng "Tổng") | Chỉ hỏi kiểu bảng (góc 6×6), có thuật toán dự phòng | 1 lần / bảng |
| 2 | Tóm tắt bảng (tên cột chính xác, kiểu dữ liệu, giá trị của cột ít lựa chọn, 1–2 dòng mẫu) | Không | 1 lần / bảng |
| 3 | Lọc bảng liên quan + gợi ý giá trị, rồi nhờ AI viết code pandas | Có | 1 lần / câu |
| 4 | Chạy code; lỗi thì gửi dòng lỗi cho AI sửa (tối đa 2–3 lần) | Có (khi lỗi) | 1 lần / câu |
| 5 | Chuẩn hóa đáp án (11.0 → 11, bỏ chữ thừa) và ghi CSV | Không | 1 lần / câu |

## Quy định nộp bài
- File `teamXXX.zip` ≤ 100MB, `Dockerfile` nằm ngay gốc.
- Docker: Linux amd64, không GPU, cài sẵn mọi thư viện.
- Nhận 6 tham số: `--xlsx-dir`, `--queries`, `--output`, `--openrouter-key`, `--openrouter-model`, `--openrouter-base-url`.
- Log in ra màn hình, không ghi vào CSV. Không sửa `/input`. Thoát mã 0 sau khi ghi xong.
- Không để API key trong bài nộp.

## Nguyên tắc
1. Đúng trước, nhanh sau, tiết kiệm sau cùng.
2. Thuật toán làm được thì không nhờ AI.
3. Việc làm 1 lần / bảng thì không lặp lại cho từng câu hỏi.
4. Luôn đo độ chính xác bằng bộ test tự tạo (`test_data/`).

## Bộ test (2 tầng)
Bộ test chia theo 3 trục: **cấu trúc bảng × kiểu dữ liệu × loại câu hỏi**. Lĩnh vực (giáo dục, nhân sự, bán hàng,
logistics, y tế, tài chính, kho, doanh nghiệp) chỉ là "lớp áo" xoay vòng. Mỗi câu/bảng có gắn nhãn đặc điểm.

| Lệnh | Việc |
|---|---|
| `python tools/tao_bo_test.py` | Tạo lại bộ test: 11 file, 14 bảng, 31 câu |
| `python tools/test_tang1.py` | **Tầng 1** – không AI, miễn phí: kiểm tra Bước 1–2 + chọn bảng. Chạy sau MỖI lần sửa code |
| `python tools/cham_diem.py --key KEY` | **Tầng 2** – có AI: đo % câu đúng, báo cáo theo nhãn |
| `... --nhan ngay_chu,ma_so_0` | Chỉ chạy câu có nhãn đó |
| `... --chi-cau-sai` | Chỉ chạy lại câu sai lần trước |
| `... --khong-cache` | Bỏ qua câu trả lời AI đã lưu (mặc định: prompt không đổi → không gọi lại AI) |

Khi thêm tính năng mới: thêm bảng/câu test có nhãn tương ứng vào `tao_bo_test.py` TRƯỚC, chạy thấy ❌, rồi mới sửa code.

### Bộ thi thử `excel/` (tạo bằng `python tools/tao_excel_thu.py`)
5 dạng bảng MỚI, không dùng lúc viết code → đo khả năng xử lý bảng lạ. Thêm `--bo-test excel` vào lệnh test.
Lần chạy đầu (tầng 1): **106/120**
- ✅ Đạt: nhiều sheet cùng cấu trúc (3 lớp), thời khóa biểu (bảng chéo ô chữ), bảng tiếng Anh, ngày ISO, True/False,
  tiêu đề công ty + tiêu đề 2 tầng + ô gộp dọc
- ✅ ĐÃ SỬA – Hai bảng trong 1 sheet: tách theo dòng trống → "file::Sheet#1", "#2"; dòng tiêu đề gần nhất
  ("BẢNG NHẬP KHO") được giữ lại, đưa vào tóm tắt + từ khóa để chọn đúng bảng
- ✅ ĐÃ SỬA – Dòng "Cộng Miền Bắc"… giữa bảng: xóa khi nhãn "Cộng…/Tổng cộng/Total" là ô chữ duy nhất của dòng và dòng có số
- ❌ CHƯA SỬA – Tiền đô "$1,200" không được nhận là số
- ❌ CHƯA SỬA – Câu hỏi tiếng Việt trên bảng tiếng Anh ("Ai vào công ty sớm nhất?") không có từ trùng → chọn sai bảng
- Sau sửa: **129/133** (4 kiểm tra còn lại thuộc 2 lỗi chưa sửa)
- ⚠️ Bộ `excel/` đã được dùng để sửa code → không còn "khách quan" cho 2 dạng này. Muốn đo thật: tạo bộ thi thử mới.

Bẫy chống sửa quá tay (trong test_data, `bao_cao_tai_chinh.xlsx`): dòng trống lỡ tay giữa bảng KHÔNG tách,
chữ ký "Người lập biểu" cuối sheet bị bỏ, "Tổng doanh thu" trong bảng chỉ tiêu và "Cộng tác viên" KHÔNG bị xóa.

## Tiến độ
- [x] Bước 0: Khung chương trình (nhận tham số, đọc câu hỏi, ghi CSV, Dockerfile)
- [x] Bước 1: Đọc & làm sạch bảng (`smartexcel/doc_bang.py`, xem kết quả: `python tools/xem_bang.py`)
  - Đã xử lý: ô gộp (ngang + dọc), dòng tiêu đề báo cáo, tiêu đề 2 tầng, bảng dọc (kể cả 2 cột "thuộc tính | giá trị"),
    dòng "Tổng cộng", số dạng chữ "1.200.000"
  - Đã xử lý (sau bộ test 3 trục): mã có số 0 đầu giữ nguyên chữ (SĐT, mã bưu chính, mã số thuế, mã hàng);
    số có đơn vị "12,5 kg" / "35.000 đ" / "80%" → số, đơn vị ghi vào tên cột "Trọng lượng (kg)";
    ngày dạng chữ "15/10/2026" → ngày; giờ "08:30" giữ chữ
  - Đoán hướng: chỉ xét TỪ ĐẦU của ô (tránh nhầm tên người "Nguyễn Văn An"); không thêm từ dễ trùng tên (văn, lê, công ty)
  - Tầng 1 (test_data): đạt 231/231 kiểm tra
  - Nhiều bảng trong 1 sheet (xếp chồng theo chiều dọc), dòng cộng nhóm giữa bảng: đã xử lý
  - Chưa xử lý: 2 bảng đặt CẠNH NHAU (cách nhau bởi cột trống), file .xls/.csv, tiền "$1,200"
  - Chưa thử phần AI phân loại bảng (cần key OpenRouter)
- [x] Bước 2: Tóm tắt bảng (`smartexcel/tom_tat.py`)
  - Mỗi bảng: tên cột chính xác + kiểu (số/ngày/chữ) + khoảng giá trị hoặc danh sách giá trị + 2 dòng mẫu CSV
  - Kèm tập từ khóa không dấu để Bước 3 lọc bảng liên quan
  - Cột mỗi dòng 1 giá trị riêng (mã, SĐT) chỉ đưa 3 ví dụ thay vì liệt kê hết
  - Chưa tối ưu: bảng rất rộng (30+ cột cùng kiểu) vẫn liệt kê từng cột
- [x] Bước 3: AI viết code (`smartexcel/hoi_dap.py`, xem: `python tools/xem_prompt.py`)
  - Chọn bảng: điểm = Σ 1/(số bảng chứa từ) cho các từ trùng; giữ tối đa 3 bảng có điểm ≥ 40% bảng đầu; không khớp gì → gửi hết
  - Gợi ý giá trị: tìm sẵn giá trị chữ mà câu hỏi nhắc tới (không dấu, theo nguyên từ)
  - Prompt: quy tắc + 1 ví dụ mẫu, AI chỉ trả code gán vào biến `answer`
  - Test: chọn đúng bảng 31/31 câu (kể cả câu ghép 2 sheet), prompt khoảng 1.100–1.300 ký tự/câu
  - Chưa thử với AI thật (cần key OpenRouter)
  - Rủi ro: câu hỏi cần ghép 2 bảng có thể bị lọc mất bảng thứ 2
- [x] Bước 4: Chạy code + sửa lỗi (`smartexcel/chay_code.py`)
  - Chạy code AI trong môi trường riêng: chặn `open`, `exec`, `exit`, chỉ cho import pandas/numpy/math/re/datetime…
  - Giới hạn 20 giây/lần chạy; code AI sửa bảng không ảnh hưởng bảng gốc
  - Lỗi → gửi dòng lỗi + dòng code gây lỗi (+ danh sách cột thật nếu sai tên cột) cho AI sửa, tối đa 2 lần sửa
  - Kết quả rỗng/NaN cũng coi là lỗi (thường do lọc sai giá trị)
  - Đã thử bằng AI giả lập 7 tình huống; chưa thử với AI thật
- [x] Phương án dự phòng (`smartexcel/du_phong.py`) – chỉ chạy với câu mà cách chính thất bại
  - Dự phòng 1 "viết lại từ đầu, nhìn rộng hơn": cuộc trò chuyện mới, gửi mọi bảng có liên quan (≤ 8000 ký tự),
    báo lỗi lần trước → AI viết code mới (sửa tối đa 1 lần)
  - Dự phòng 2 "đọc thẳng dữ liệu": gửi ≤ 40 dòng liên quan nhất mỗi bảng (ưu tiên dòng chứa giá trị câu hỏi nhắc tới),
    AI trả lời thẳng không viết code
  - Mọi cách đều sai → để trống. Tối đa 6 lần gọi AI/câu (3 + 2 + 1)
  - Bước 5 hiểu thêm số có đơn vị/dấu phân cách trong câu trả lời thẳng: "800.000 đồng" → "800000"
  - Đã thử 5 kịch bản bằng AI giả lập; tầng 1: 218/218
- [x] Bước 5: Chuẩn hóa đáp án (`smartexcel/chuan_hoa.py`)
  - Số: 9.0 → "9", làm tròn 2 chữ số thập phân, dấu chấm thập phân, không phân cách hàng nghìn
  - Series/DataFrame/list: 1 phần tử → lấy phần tử; nhiều → nối ", "
  - Ngày → "dd/mm/yyyy"; True/False → "Có"/"Không"; NaN/rỗng → ""
  - Chữ: bỏ "Đáp án:", ngoặc kép, dấu chấm cuối, xuống dòng; mã có số 0 đầu giữ nguyên
  - Tầng 1: 32 trường hợp chuẩn hóa đạt hết (tổng 176/176)
  - ⚠️ CHƯA CHẮC định dạng đề muốn (số chữ số thập phân, kiểu ngày, Có/Không) → chỉnh hằng số đầu file khi có đáp án mẫu
- [ ] Đóng gói ZIP & test Docker
