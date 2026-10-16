"""Kết nối tới Qwen 3.7 Flash qua OpenRouter (dùng thư viện openai, API chuẩn OpenAI)."""
import hashlib
import json
import os
import time

from openai import OpenAI

from .log import log


class AI:
    def __init__(self, key, model, base_url):
        self.model = model
        self.client = OpenAI(api_key=key, base_url=base_url, timeout=60) if key else None
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
            try:
                res = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    temperature=0,
                    max_tokens=max_tokens,
                )
                tra_loi = (res.choices[0].message.content or "").strip()
                if khoa:
                    self.cache[khoa] = tra_loi
                    with open(self.file_cache, "w", encoding="utf-8") as f:
                        json.dump(self.cache, f, ensure_ascii=False)
                return tra_loi
            except Exception as e:
                log(f"Gọi AI lỗi (lần {lan + 1}): {e}")
                time.sleep(2 * (lan + 1))
        return ""
