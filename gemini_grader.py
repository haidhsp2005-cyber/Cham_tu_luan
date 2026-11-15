"""
Gemini Vision Grading Pipeline for Handwritten Vietnamese Exams (GDPT 2018)
Requires: google-genai, pydantic, pillow
"""

import os
import json
from typing import List, Dict, Optional
from pydantic import BaseModel, Field

# Định nghĩa Pydantic Schema cho kết quả bóc tách từ Gemini Vision
class StudentInfo(BaseModel):
    sbd: str = Field(description="Số báo danh của học sinh ghi trong ô lục giác hoặc khung phách")
    full_name: Optional[str] = Field(default="", description="Họ và tên nếu có")
    exam_code_detected: Optional[str] = Field(default="301", description="Mã đề thi phát hiện được")

class Part1Item(BaseModel):
    question: int
    student_choice: str = Field(description="Đáp án A, B, C, hoặc D (bỏ qua dấu gạch xóa, lấy kết quả sau cùng)")
    confidence: float = Field(ge=0.0, le=1.0)

class Part2SubItem(BaseModel):
    a: str = Field(description="Đ hoặc S")
    b: str = Field(description="Đ hoặc S")
    c: str = Field(description="Đ hoặc S")
    d: str = Field(description="Đ hoặc S")

class Part2Item(BaseModel):
    question: int
    sub_items: Part2SubItem
    confidence: float

class EssayCriterionEval(BaseModel):
    criterion_id: str
    criterion_name: str
    matched: bool
    evidence: str = Field(description="Trích dẫn bằng chứng từ bài làm của học sinh chứng minh cho việc đạt hoặc chưa đạt")
    awarded_score: float
    max_score: float

class EssayQuestionEval(BaseModel):
    question_number: int
    transcribed_text: str = Field(description="Nội dung chữ viết tay/mã HTML học sinh viết trên giấy")
    evaluations: List[EssayCriterionEval]
    subtotal_score: float
    teacher_feedback: str = Field(description="Nhận xét ngắn gọn, chỉ rõ lỗi hoặc ưu điểm")

class ExamExtractionResult(BaseModel):
    student_info: StudentInfo
    part_1: List[Part1Item]
    part_2: List[Part2Item]
    part_3: List[EssayQuestionEval]

def build_system_prompt(master_key: dict) -> str:
    return f"""
Bạn là giám khảo khảo thí kỳ thi tốt nghiệp THPT tại Việt Nam.
Nhiệm vụ của bạn là đọc bài làm viết tay của học sinh từ ảnh chụp tờ giấy thi và bóc tách câu trả lời theo Barem môn Tin học 12 (Mã đề: {master_key['exam_info']['exam_code']}).

LƯU Ý ĐẶC BIỆT KHI ĐỌC CHỮ VIẾT TAY:
1. Số báo danh (SBD): Nằm ở góc trên bên phải trong ô hình lục giác.
2. Phần I (16 câu trắc nghiệm): Học sinh có thể viết 1.B, 2.A hoặc kẻ bảng hoặc viết liền 1B, 2A. Nếu có gạch xóa, hãy lấy chữ cái viết lại sau cùng.
3. Phần II (Đúng/Sai): Học sinh có thể viết Đ/S, D/S, True/False. Hãy chuẩn hóa thành "Đ" hoặc "S".
4. Phần III (Tự luận):
   - Câu 1: Đánh giá dựa trên 3 tiêu chí: Cảnh báo nguy cơ, Nguyên nhân đạo đức/thiên vị, Trách nhiệm công dân.
   - Câu 2: Đánh giá kỹ năng viết thẻ HTML (<h1 style='color:red'>, <p>, <sub>, color:blue/green). Lưu ý chữ viết tay thẻ HTML có thể hơi nguệch ngoạc nhưng nếu đúng bản chất thẻ và thuộc tính thì cho điểm.
"""

def grade_paper_image(image_path: str, master_key_path: str = "answer_key_301.json"):
    from google import genai
    from google.genai import types

    with open(master_key_path, "r", encoding="utf-8") as f:
        master_key = json.load(f)

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("[!] GEMINI_API_KEY chưa được thiết lập. Hãy gán biến môi trường GEMINI_API_KEY để chạy thực tế.")
        return None

    client = genai.Client(api_key=api_key)
    system_prompt = build_system_prompt(master_key)

    # Đọc ảnh
    with open(image_path, "rb") as img_file:
        image_bytes = img_file.read()

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[
            types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
            f"Hãy bóc tách toàn bộ bài làm của học sinh trên tờ giấy thi này và chấm điểm theo barem sau:\n{json.dumps(master_key, ensure_ascii=False)}"
        ],
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            response_mime_type="application/json",
            response_schema=ExamExtractionResult,
            temperature=0.1
        )
    )

    result_json = json.loads(response.text)
    return result_json

if __name__ == "__main__":
    print("[*] Gemini Vision Grader module loaded successfully.")
