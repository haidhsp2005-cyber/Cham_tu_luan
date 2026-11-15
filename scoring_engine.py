"""
Scoring Engine for Vietnamese High School Exams (GDPT 2018 Standard & Flexible)
Hỗ trợ LINH HOẠT SỐ CÂU:
1. Phần I: Trắc nghiệm đơn (A, B, C, D) - Linh hoạt từ 1 đến 60+ câu.
   - Tự động chia điểm mỗi câu = Tổng điểm Phần I / Số câu.
2. Phần II: Trắc nghiệm Đúng / Sai (a, b, c, d) - Linh hoạt số lượng câu hỏi.
   - Áp dụng chuẩn bậc thang Bộ GD&ĐT: 1 ý = 0.1đ, 2 ý = 0.25đ, 3 ý = 0.5đ, 4 ý = 1.0đ.
3. Phần III: Tự luận - Linh hoạt số lượng câu hỏi và thang điểm.
"""

import json
import re
from typing import Dict, Any, List

def normalize_choice(val: str) -> str:
    """Chuẩn hóa đáp án trắc nghiệm A, B, C, D từ chữ viết tay"""
    if not val:
        return ""
    v = str(val).strip().upper()
    clean = re.sub(r"[^A-D]", "", v)
    return clean[0] if clean else ""

def normalize_true_false(val: str) -> str:
    """Chuẩn hóa ký hiệu Đúng / Sai từ nhiều cách viết tay khác nhau"""
    if not val:
        return ""
    v = str(val).strip().lower()
    if v in ["đ", "d", "t", "true", "v", "✓", "1", "dung", "đúng"]:
        return "Đ"
    if v in ["s", "f", "false", "x", "✗", "0", "sai"]:
        return "S"
    return ""

def score_part_1(student_answers: Dict[str, str], answer_key: Dict[str, Any]) -> Dict[str, Any]:
    """Chấm Phần I: Trắc nghiệm đơn - Tự động thích ứng với số lượng câu bất kỳ (10, 16, 20, 40, 50...)"""
    details = []
    total_score = 0.0
    correct_count = 0
    key = answer_key.get("answers", {})
    total_q = len(key) if len(key) > 0 else 16
    max_score = answer_key.get("max_score", 4.0)
    
    # Tính điểm mỗi câu linh hoạt nếu chưa có
    unit_score = answer_key.get("score_per_question")
    if not unit_score or unit_score <= 0:
        unit_score = round(max_score / total_q, 4) if total_q > 0 else 0.25
    
    # Duyệt theo thứ tự số câu thực tế trong đề
    sorted_q_keys = sorted(key.keys(), key=lambda x: int(x) if str(x).isdigit() else 999)
    for q_num in sorted_q_keys:
        correct_ans = key[q_num]
        raw_ans = student_answers.get(str(q_num), "")
        cleaned_ans = normalize_choice(raw_ans)
        is_correct = (cleaned_ans == correct_ans)
        
        if is_correct:
            correct_count += 1
            total_score += unit_score
        
        details.append({
            "question": int(q_num) if str(q_num).isdigit() else q_num,
            "raw_answer": raw_ans,
            "normalized_answer": cleaned_ans,
            "correct_answer": correct_ans,
            "is_correct": is_correct,
            "score": round(unit_score, 3) if is_correct else 0.0
        })
        
    return {
        "part": "Phần I",
        "total_score": round(min(total_score, max_score), 2),
        "max_score": max_score,
        "correct_count": correct_count,
        "total_questions": len(sorted_q_keys),
        "unit_score": unit_score,
        "details": details
    }

def score_part_2(student_answers: Dict[str, Dict[str, str]], answer_key: Dict[str, Any]) -> Dict[str, Any]:
    """Chấm Phần II: Đúng / Sai theo thang điểm bậc thang GDPT 2018"""
    details = []
    total_score = 0.0
    key_questions = answer_key.get("questions", {})
    staircase_scale = {0: 0.0, 1: 0.10, 2: 0.25, 3: 0.50, 4: 1.00}
    
    sorted_q_keys = sorted(key_questions.keys(), key=lambda x: int(x) if str(x).isdigit() else 999)
    for q_num in sorted_q_keys:
        correct_sub = key_questions[q_num]
        student_sub = student_answers.get(str(q_num), {})
        sub_details = {}
        correct_in_q = 0
        
        for item in ["a", "b", "c", "d"]:
            raw_val = student_sub.get(item, "")
            norm_val = normalize_true_false(raw_val)
            expected_val = correct_sub.get(item, "")
            is_ok = (norm_val == expected_val)
            if is_ok:
                correct_in_q += 1
            sub_details[item] = {
                "raw": raw_val,
                "normalized": norm_val,
                "expected": expected_val,
                "is_correct": is_ok
            }
            
        q_score = staircase_scale.get(correct_in_q, 0.0)
        total_score += q_score
        
        details.append({
            "question": int(q_num) if str(q_num).isdigit() else q_num,
            "correct_sub_count": correct_in_q,
            "score": q_score,
            "sub_items": sub_details
        })
        
    max_score = answer_key.get("max_score", len(sorted_q_keys) * 1.0)
    return {
        "part": "Phần II",
        "total_score": round(min(total_score, max_score), 2),
        "max_score": max_score,
        "details": details
    }
