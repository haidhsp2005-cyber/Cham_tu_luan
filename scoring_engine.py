"""
Scoring Engine for Vietnamese High School Exams (GDPT 2018 Standard)
Supports:
1. Part I: Multiple Choice Single Answer (A, B, C, D)
2. Part II: True / False (4 sub-items: a, b, c, d) with GDPT 2018 staircase grading:
   - 1 correct: 0.10 pt
   - 2 correct: 0.25 pt
   - 3 correct: 0.50 pt
   - 4 correct: 1.00 pt
3. Part III: Essay / Code grading against criteria
"""

import json
import re
from typing import Dict, Any, List

def normalize_choice(val: str) -> str:
    """Chuẩn hóa đáp án trắc nghiệm A, B, C, D từ chữ viết tay"""
    if not val:
        return ""
    v = str(val).strip().upper()
    # Loại bỏ dấu ngoặc đơn, chấm ví dụ "(A)" -> "A", "A." -> "A"
    clean = re.sub(r"[^A-D]", "", v)
    return clean[0] if clean else ""

def normalize_true_false(val: str) -> str:
    """Chuẩn hóa ký hiệu Đúng / Sai từ nhiều cách viết tay khác nhau"""
    if not val:
        return ""
    v = str(val).strip().lower()
    
    # Các cách viết 'Đúng': Đ, đ, d, t, true, v, checkmark
    if v in ["đ", "d", "t", "true", "v", "✓", "1", "dung", "đúng"]:
        return "Đ"
    # Các cách viết 'Sai': S, s, f, false, x, cross
    if v in ["s", "f", "false", "x", "✗", "0", "sai"]:
        return "S"
    return ""

def score_part_1(student_answers: Dict[str, str], answer_key: Dict[str, Any]) -> Dict[str, Any]:
    """Chấm Phần I: Trắc nghiệm 4 lựa chọn"""
    details = []
    total_score = 0.0
    correct_count = 0
    key = answer_key["answers"]
    unit_score = answer_key.get("score_per_question", 0.25)
    
    for q_num, correct_ans in key.items():
        raw_ans = student_answers.get(str(q_num), "")
        cleaned_ans = normalize_choice(raw_ans)
        is_correct = (cleaned_ans == correct_ans)
        if is_correct:
            correct_count += 1
            total_score += unit_score
        
        details.append({
            "question": int(q_num),
            "raw_answer": raw_ans,
            "normalized_answer": cleaned_ans,
            "correct_answer": correct_ans,
            "is_correct": is_correct,
            "score": unit_score if is_correct else 0.0
        })
        
    return {
        "part": "Phần I",
        "total_score": round(total_score, 2),
        "max_score": answer_key.get("max_score", 4.0),
        "correct_count": correct_count,
        "total_questions": len(key),
        "details": details
    }

def score_part_2(student_answers: Dict[str, Dict[str, str]], answer_key: Dict[str, Any]) -> Dict[str, Any]:
    """Chấm Phần II: Đúng / Sai theo thang điểm bậc thang GDPT 2018"""
    details = []
    total_score = 0.0
    key_questions = answer_key["questions"]
    staircase_scale = {
        0: 0.0,
        1: 0.10,
        2: 0.25,
        3: 0.50,
        4: 1.00
    }
    
    for q_num, correct_sub in key_questions.items():
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
            "question": int(q_num),
            "correct_sub_count": correct_in_q,
            "score": q_score,
            "sub_items": sub_details
        })
        
    return {
        "part": "Phần II",
        "total_score": round(total_score, 2),
        "max_score": answer_key.get("max_score", 4.0),
        "details": details
    }

if __name__ == "__main__":
    # Test mô phỏng
    with open("answer_key_301.json", "r", encoding="utf-8") as f:
        master_key = json.load(f)
        
    # Giả lập bài học sinh:
    mock_student_part1 = {
        "1": "B", "2": "A", "3": "C", "4": "A",
        "5": "D", "6": "D", "7": "D", "8": "B",
        "9": "D", "10": "C", "11": "B", "12": "b.", # Sai câu 12 (Đáp án là A)
        "13": "A", "14": "b", "15": "B", "16": "C"
    }
    
    # Học sinh viết lúc d, lúc Đ, lúc S, lúc F
    mock_student_part2 = {
        "1": { "a": "s", "b": "đ", "c": "d", "d": "S" }, # Đúng cả 4 ý -> 1.0đ
        "2": { "a": "S", "b": "Đ", "c": "Đ", "d": "s" }, # Đúng cả 4 ý -> 1.0đ
        "3": { "a": "Đ", "b": "đ", "c": "s", "d": "đ" }, # Đúng 3 ý (a, b, c) -> 0.5đ
        "4": { "a": "d", "b": "s", "c": "s", "d": "đ" }  # Đúng cả 4 ý -> 1.0đ
    }
    
    r1 = score_part_1(mock_student_part1, master_key["part_1_mcq"])
    r2 = score_part_2(mock_student_part2, master_key["part_2_true_false"])
    
    print("=== KẾT QUẢ CHẤM THỬ NGHIỆM ===")
    print(f"Phần I:  {r1['total_score']} / {r1['max_score']} (Đúng {r1['correct_count']}/16 câu)")
    print(f"Phần II: {r2['total_score']} / {r2['max_score']} (Thang điểm bậc thang)")
    print(f"Tổng Trắc nghiệm: {r1['total_score'] + r2['total_score']} / 8.0 điểm")
