"""Kết nối tới Qwen 3.7 Flash qua OpenRouter (dùng thư viện openai, API chuẩn OpenAI)."""
import hashlib
import json
import os
import time

from openai import OpenAI

from . import cau_hinh
from .log import log

THOI_GIAN_MOI_LAN_GOI = 60   # giây tối đa cho 1 lần gọi API
CON_LAI_TOI_THIEU = 10       # còn ít hơn chừng này giây thì thôi không gọi AI nữa


class AI:
    def __init__(self, key, model, base_url):
        self.model = model
        # max_retries=0: tắt retry ngầm của thư viện openai (mặc định 2 lần) → số lần gọi
        # chỉ do vòng lặp trong hoi() quyết định, dễ kiểm soát thời gian và chi phí.
        self.client = OpenAI(api_key=key, base_url=base_url,
                             timeout=THOI_GIAN_MOI_LAN_GOI, max_retries=0) if key else None
        # Chỉ dùng khi test trên máy (tools/cham_diem.py đặt biến này): lưu câu trả lời
        # của AI để lần chạy sau, prompt nào không đổi thì khỏi gọi lại → khỏi tốn tiền.
        # Lúc máy chấm chạy không có biến này → luôn gọi AI thật.
        self.file_cache = os.environ.get("SMARTEXCEL_CACHE")
        self.cache = {}
        if self.file_cache and os.path.exists(self.file_cache):
            with open(self.file_cache, encoding="utf-8") as f:
                self.cache = json.load(f)

    def hoi(self, messages, max_tokens=600, so_lan_thu=3):
        """Gửi danh sách messages, trả về chữ AI trả lời ("" nếu thất bại)."""
        if self.client is None:
            log("Không có API key → bỏ qua gọi AI")
            return ""
        khoa = None
        if self.file_cache:
            khoa = hashlib.sha256(json.dumps([self.model, messages, max_tokens],
                                             ensure_ascii=False).encode()).hexdigest()
            if khoa in self.cache:
                return self.cache[khoa]

        for lan in range(so_lan_thu):
            con = cau_hinh.con_lai()
            if con < CON_LAI_TOI_THIEU:
                log("Sắp hết tổng thời gian → bỏ qua gọi AI")
                return ""
            try:
                res = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    temperature=0,
                    max_tokens=max_tokens,
                    timeout=min(THOI_GIAN_MOI_LAN_GOI, con - CON_LAI_TOI_THIEU / 2),
                )
                tra_loi = (res.choices[0].message.content or "").strip()
                if khoa:
                    self.cache[khoa] = tra_loi
                    with open(self.file_cache, "w", encoding="utf-8") as f:
                        json.dump(self.cache, f, ensure_ascii=False)
                return tra_loi
            except Exception as e:
                log(f"Gọi AI lỗi (lần {lan + 1}): {e}")
                if lan + 1 < so_lan_thu:
                    time.sleep(2 * (lan + 1))
        return ""
