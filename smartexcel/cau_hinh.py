"""Cấu hình AI: Qwen 3.7 Flash qua OpenRouter (không dùng Gemini, không dùng Ollama).

API key KHÔNG ghi trong code. Thứ tự lấy key:
    1. tham số --openrouter-key (máy chấm truyền vào)
    2. biến môi trường OPENROUTER_API_KEY
    3. file .env ở thư mục gốc dự án, dòng: OPENROUTER_API_KEY=sk-or-v1-...
"""
import os
import time

MODEL = "qwen/qwen3.7-flash"
BASE_URL = "https://openrouter.ai/api/v1"
TEN_BIEN_KEY = "OPENROUTER_API_KEY"

# Máy chấm cho tối đa 40 phút → tự dừng ở 35 phút, chừa 5 phút cho khởi động Docker và ghi CSV.
# Hết giờ thì không gọi AI nữa, câu còn lại để trống, nhưng vẫn ghi đủ file kết quả.
THOI_GIAN_TONG = 35 * 60
_BAT_DAU = time.monotonic()


def con_lai():
    """Số giây còn lại trước khi hết THOI_GIAN_TONG (âm nếu đã quá giờ)."""
    return THOI_GIAN_TONG - (time.monotonic() - _BAT_DAU)

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
FILE_ENV = os.path.join(GOC, ".env")


def _doc_file_env():
    """Đọc file .env dạng TEN=gia_tri, bỏ qua dòng trống và dòng bắt đầu bằng #."""
    ket_qua = {}
    if not os.path.exists(FILE_ENV):
        return ket_qua
    with open(FILE_ENV, encoding="utf-8-sig") as f:
        for dong in f:
            dong = dong.strip()
            if not dong or dong.startswith("#") or "=" not in dong:
                continue
            ten, gia_tri = dong.split("=", 1)
            ket_qua[ten.strip()] = gia_tri.strip().strip('"').strip("'")
    return ket_qua


def lay_key():
    """Key từ biến môi trường, không có thì lấy từ file .env ("" nếu không có)."""
    return os.environ.get(TEN_BIEN_KEY) or _doc_file_env().get(TEN_BIEN_KEY, "")
