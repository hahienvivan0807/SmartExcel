# SmartExcel

Bài dự thi **Cuộc thi Sáng tạo trẻ Quốc gia trong lĩnh vực AI 2026 – Bảng C**.

Chương trình đọc nhiều file Excel, trả lời danh sách câu hỏi về dữ liệu trong đó và ghi đáp án ra file CSV `id,answer`.
Ý tưởng chính: **AI hiểu câu hỏi và viết code pandas, máy tính chạy code để ra con số chính xác.**

Kế hoạch chi tiết, tiến độ và các quyết định thiết kế: xem [KE_HOACH.md](KE_HOACH.md).

## Quy trình 5 bước

```
Excel ──► [1] Đọc & làm sạch bảng        smartexcel/doc_bang.py   (1 lần / bảng)
              ▼
          [2] Tóm tắt bảng               smartexcel/tom_tat.py    (1 lần / bảng, không dùng AI)
              ▼
Câu hỏi ─►[3] Chọn bảng + AI viết code   smartexcel/hoi_dap.py
              ▼
          [4] Chạy code, lỗi → AI sửa    smartexcel/chay_code.py
              │  vẫn sai → dự phòng      smartexcel/du_phong.py
              ▼
          [5] Chuẩn hóa đáp án → CSV     smartexcel/chuan_hoa.py
```

## Cài đặt

Cần Python 3.12.

```bash
pip install -r requirements.txt
```

## Chạy giống máy chấm

```bash
python main.py --xlsx-dir test_data/xlsx_files --queries test_data/queries.json --output test_data/output/submission.csv --openrouter-key KEY_CUA_BAN --openrouter-model qwen/qwen3.7-flash --openrouter-base-url https://openrouter.ai/api/v1
```

⚠️ **Không bao giờ ghi key vào code hay commit lên GitHub.** Khi chạy trên máy, copy `.env.example`
thành `.env` và dán key OpenRouter vào (hoặc đặt biến môi trường `OPENROUTER_API_KEY`) – `main.py` và
các công cụ test tự đọc, khỏi phải gõ `--key` mỗi lần. Chi tiết: [CACH_HOAT_DONG.md](CACH_HOAT_DONG.md).

## Chạy bằng Docker

```bash
docker build --platform linux/amd64 -t ai26c0072 .
```

Truyền key theo **một trong hai cách**:

```bash
# Cách 1: truyền thẳng qua tham số
docker run --rm -v "$(pwd)/test_data:/data:ro" -v "$(pwd)/out:/out" ai26c0072 \
  --xlsx-dir /data/xlsx_files --queries /data/queries.json --output /out/submission.csv \
  --openrouter-key KEY_CUA_BAN --openrouter-model qwen/qwen3.7-flash --openrouter-base-url https://openrouter.ai/api/v1

# Cách 2: copy .env.example thành .env, dán key vào, rồi thêm --env-file .env
docker run --rm --env-file .env -v "$(pwd)/test_data:/data:ro" -v "$(pwd)/out:/out" ai26c0072 \
  --xlsx-dir /data/xlsx_files --queries /data/queries.json --output /out/submission.csv
```

File `.env` **không** được đóng vào image, nên chạy Docker phải có `--env-file .env` thì mới đọc được key.
Không có key chương trình vẫn chạy và ghi đủ `submission.csv`, nhưng đáp án để trống.

## Hướng dẫn kiểm thử

### Bước 0 – Tạo dữ liệu test (chạy 1 lần, hoặc sau khi sửa script tạo test)

```bash
python tools/tao_bo_test.py
python tools/tao_excel_thu.py
```

| Folder | Vai trò |
|---|---|
| `test_data/` | Bộ test phát triển: 11 file, 14 bảng, 31 câu. Chia theo 3 trục **cấu trúc bảng × kiểu dữ liệu × loại câu hỏi**, lĩnh vực xoay vòng (giáo dục, logistics, y tế, tài chính…) |
| `excel/` | Bộ thi thử: 5 dạng bảng lạ (báo cáo có dòng cộng nhóm, nhiều sheet, thời khóa biểu, bảng tiếng Anh, 2 bảng trong 1 sheet) |

### Ba công cụ kiểm thử

| Lệnh | Cần key? | Thời gian | Kiểm tra gì |
|---|---|---|---|
| `python tools/test_tang1.py` | Không | vài giây | **Tầng 1**: Bước 1 (làm sạch), Bước 2 (tóm tắt), chọn bảng, dự phòng, Bước 5 (chuẩn hóa) |
| `python tools/test_ai_gia_lap.py` | Không | vài giây | Bước 3→4→5 + dự phòng, dùng AI giả lập trả lời soạn sẵn |
| `python tools/cham_diem.py --key KEY` | **Có** | vài phút | **Tầng 2**: chạy cả chương trình với AI thật, chấm % câu đúng theo từng nhãn |

Thêm `--bo-test excel` vào `test_tang1.py` hoặc `cham_diem.py` để chạy bộ thi thử.

Kết quả hiện tại:
- Tầng 1 `test_data/`: **231/231**
- Tầng 1 `excel/`: **129/133** (còn 2 lỗi đã biết: tiền đô `$1,200`, câu hỏi tiếng Việt trên bảng tiếng Anh)
- AI giả lập: **21/21**
- Tầng 2: chưa chạy (cần key)

### Tùy chọn của tầng 2 (tiết kiệm token)

```bash
python tools/cham_diem.py --key KEY --nhan ngay_chu,ma_so_0   # chỉ câu có nhãn này
python tools/cham_diem.py --key KEY --chi-cau-sai             # chỉ chạy lại câu sai lần trước
python tools/cham_diem.py --key KEY --khong-cache             # bắt buộc gọi AI thật
```

Câu trả lời của AI được lưu ở `<bộ test>/output/cache_ai.json`: prompt không đổi thì lần sau không gọi AI lại.
Cache này chỉ hoạt động khi chạy bằng `cham_diem.py`, lúc máy chấm chạy luôn gọi AI thật.

### Công cụ xem bằng mắt

```bash
python tools/xem_bang.py                     # bảng sau khi làm sạch + bản tóm tắt gửi AI + từ khóa
python tools/xem_bang.py excel/xlsx_files    # xem bảng của folder khác (vd file Excel thật của đề)
python tools/xem_prompt.py                   # mỗi câu hỏi chọn bảng nào, gợi ý gì, prompt dài bao nhiêu
python tools/xem_prompt.py --cau 3           # in nguyên prompt của câu 3
python tools/xem_prompt.py --key KEY         # gọi AI thật, in code AI viết cho từng câu
```

### Quy trình khi sửa code

1. Thêm bảng/câu test cho tình huống mới vào `tools/tao_bo_test.py` **trước**, chạy để thấy ❌.
2. Sửa code.
3. Chạy `test_tang1.py` + `test_ai_gia_lap.py` (miễn phí) cho tới khi hết ❌, kể cả các ca cũ.
4. Khi cần đo độ chính xác thật mới chạy `cham_diem.py` (tốn token), ưu tiên `--nhan` / `--chi-cau-sai`.

## Nộp bài

- File `AI26C0072.zip` ≤ 100MB, `Dockerfile` nằm ngay gốc file ZIP.
- Docker chạy Linux amd64, không GPU, đã cài sẵn thư viện (`requirements.txt`).
- Không đưa key, `test_data/output/`, `excel/output/` vào file ZIP.
