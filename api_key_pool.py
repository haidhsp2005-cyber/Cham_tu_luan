"""
API Key Pool Manager for Google Gemini API
Features:
- Quản lý danh sách nhiều API Key miễn phí.
- Cơ chế Round-Robin (xoay vòng lần lượt các key).
- Tự động phát hiện lỗi 429 (ResourceExhausted / Rate Limit) và chuyển ngay sang key kế tiếp (Auto-Fallback).
- Tự động đặt Cooldown cho key bị chạm trần hạn mức.
- Thống kê lượt gọi, tỉ lệ thành công của từng key.
"""

import time
import json
import io
import os
from typing import List, Dict, Tuple, Optional
from PIL import Image

class GeminiKeyPool:
    def __init__(self, keys_input: str = ""):
        self.keys: List[Dict] = []
        self.current_idx = 0
        self.update_keys(keys_input)

    def update_keys(self, keys_input: str):
        """Cập nhật danh sách keys từ text (mỗi dòng 1 key hoặc cách nhau dấu phẩy)"""
        raw_keys = [k.strip() for k in keys_input.replace(",", "\n").split("\n") if k.strip()]
        
        # Giữ lại thống kê cũ nếu key đã tồn tại
        existing_map = {k["key"]: k for k in self.keys}
        new_keys = []
        for rk in raw_keys:
            if rk in existing_map:
                new_keys.append(existing_map[rk])
            else:
                new_keys.append({
                    "key": rk,
                    "masked": rk[:6] + "..." + rk[-4:] if len(rk) > 10 else rk,
                    "total_requests": 0,
                    "successful_requests": 0,
                    "errors": 0,
                    "cooldown_until": 0,
                    "status": "Ready"
                })
        self.keys = new_keys

    def get_available_key(self) -> Tuple[Optional[str], Optional[int]]:
        """Lấy key tiếp theo theo thuật toán Round-Robin và bỏ qua key đang trong cooldown"""
        if not self.keys:
            return None, None

        now = time.time()
        n = len(self.keys)
        
        # Thử duyệt qua các key theo vòng tròn
        for i in range(n):
            idx = (self.current_idx + i) % n
            key_info = self.keys[idx]
            
            # Kiểm tra xem key có hết thời gian cooldown chưa
            if now >= key_info["cooldown_until"]:
                key_info["status"] = "Active"
                self.current_idx = (idx + 1) % n
                return key_info["key"], idx
            else:
                key_info["status"] = f"Cooldown ({int(key_info['cooldown_until'] - now)}s)"

        # Nếu tất cả các key đều đang bị cooldown, lấy key có thời gian chờ ngắn nhất
        min_wait_idx = min(range(n), key=lambda idx: self.keys[idx]["cooldown_until"])
        wait_seconds = max(1, int(self.keys[min_wait_idx]["cooldown_until"] - now))
        return None, wait_seconds

    def record_success(self, key_idx: int):
        if 0 <= key_idx < len(self.keys):
            self.keys[key_idx]["total_requests"] += 1
            self.keys[key_idx]["successful_requests"] += 1
            self.keys[key_idx]["status"] = "Ready"

    def record_rate_limit(self, key_idx: int, cooldown_seconds: int = 40):
        """Khi key bị 429, cho key nghỉ một khoảng thời gian"""
        if 0 <= key_idx < len(self.keys):
            self.keys[key_idx]["total_requests"] += 1
            self.keys[key_idx]["errors"] += 1
            self.keys[key_idx]["cooldown_until"] = time.time() + cooldown_seconds
            self.keys[key_idx]["status"] = f"Cooldown ({cooldown_seconds}s)"

    def get_stats(self) -> List[Dict]:
        return [
            {
                "Key": k["masked"],
                "Trạng thái": k["status"],
                "Tổng gọi": k["total_requests"],
                "Thành công": k["successful_requests"],
                "Lỗi": k["errors"]
            }
            for k in self.keys
        ]

    def call_gemini_with_rotation(self, image_pil: Image.Image, master_db: dict, model_name: str = "gemini-2.5-pro", max_retries: int = 3):
        """Gọi Gemini API với cơ chế tự động xoay vòng key khi gặp lỗi Rate Limit (429)"""
        from google import genai
        from google.genai import types
        from gemini_grader import build_system_prompt, ExamExtractionResult

        if not self.keys:
            # Fallback lấy key từ biến môi trường nếu có
            env_key = os.environ.get("GEMINI_API_KEY", "").strip()
            if env_key:
                self.update_keys(env_key)
            else:
                return None, "Chưa thiết lập API Key nào trong Key Pool."

        # Chuẩn bị ảnh
        img_byte_arr = io.BytesIO()
        image_pil.save(img_byte_arr, format='JPEG', quality=95)
        image_bytes = img_byte_arr.getvalue()

        sample_key = list(master_db.values())[0] if master_db else {}
        system_prompt = build_system_prompt(sample_key)

        prompt_content = f"""Bạn là giám khảo chấm thi tốt nghiệp THPT.
Dưới đây là KHO BAREM ĐÁP ÁN của các mã đề:
{json.dumps(master_db, ensure_ascii=False, indent=2)}

Nhiệm vụ của bạn:
1. Quan sát ảnh bài thi viết tay của học sinh.
2. Tìm và bóc tách:
   - SBD (Số báo danh trong ô lục giác hoặc khung phách).
   - Mã đề thi mà học sinh ghi trên bài thi (ví dụ 301, 302...).
3. Bóc tách câu trả lời của học sinh:
   - Phần I: 16 câu trắc nghiệm (A, B, C, D). Nếu học sinh gạch xóa viết lại, lấy chữ cái cuối cùng.
   - Phần II: 4 câu Đúng/Sai (các ý a, b, c, d với giá trị Đ hoặc S).
   - Phần III: Đọc toàn bộ chữ viết tay phần Tự luận (chú ý thẻ HTML và bài văn luận), so sánh với barem của mã đề học sinh làm và chấm điểm theo tiêu chí.
"""

        attempts = 0
        last_error = ""

        while attempts < len(self.keys) * 2 and attempts < max_retries * len(self.keys):
            api_key, key_idx = self.get_available_key()
            if not api_key:
                time.sleep(key_idx if key_idx else 2)
                continue

            attempts += 1
            try:
                client = genai.Client(api_key=api_key)
                response = client.models.generate_content(
                    model=model_name,
                    contents=[
                        types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
                        prompt_content
                    ],
                    config=types.GenerateContentConfig(
                        system_instruction=system_prompt,
                        response_mime_type="application/json",
                        response_schema=ExamExtractionResult,
                        temperature=0.1
                    )
                )
                data = json.loads(response.text)
                self.record_success(key_idx)
                return data, None

            except Exception as e:
                err_str = str(e)
                last_error = err_str
                # Kiểm tra lỗi Rate Limit 429 hoặc Quota Exceeded
                if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str or "Quota" in err_str:
                    # Cho key này nghỉ 45 giây và thử ngay key kế tiếp!
                    self.record_rate_limit(key_idx, cooldown_seconds=45)
                    continue
                else:
                    # Các lỗi khác (như ảnh lỗi, prompt...)
                    return None, f"Lỗi gọi AI: {err_str}"

        return None, f"Tất cả các API Key đều tạm thời chạm trần hạn mức. Lỗi cuối: {last_error}"
