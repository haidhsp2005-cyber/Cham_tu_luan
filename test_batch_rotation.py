"""
Kiểm thử tự động tính năng Chấm Hàng Loạt & Xoay Vòng API Key Pool (Gemini 2.5)
Xuất kết quả bảng điểm lớp học ra file Excel (.xlsx) chuẩn
"""

import os
import json
import time
import pandas as pd
from api_key_pool import GeminiKeyPool
from scoring_engine import score_part_1, score_part_2

def run_batch_simulation():
    print("=" * 65)
    print("🚀 BẮT ĐẦU CHẠY THỬ NGHIỆM CHẤM HÀNG LOẠT VỚI API KEY POOL")
    print("=" * 65)

    # 1. Đọc kho barem
    with open("master_keys_db.json", "r", encoding="utf-8") as f:
        master_db = json.load(f)
    print(f"[*] Đã tải kho barem: {list(master_db.keys())} (Môn Tin học 11)")

    # 2. Khởi tạo Key Pool với 3 Key mẫu
    mock_keys = """AIzaSyA_FreeKey_Gmail1_Test123
AIzaSyB_FreeKey_Gmail2_Test456
AIzaSyC_FreeKey_Gmail3_Test789"""
    
    pool = GeminiKeyPool(mock_keys)
    print(f"[*] Đã khởi tạo Key Pool với {len(pool.keys)} keys miễn phí:")
    for k in pool.keys:
        print(f"    - {k['masked']} (Trạng thái: {k['status']})")

    # 3. Giả lập danh sách 5 học sinh trong lớp nộp bài
    mock_students = [
        {
            "stt": 1, "sbd": "720141", "name": "Nguyễn Văn An", "code": "301",
            "p1_answers": {"1": "B", "2": "A", "3": "C", "4": "A", "5": "D", "6": "D", "7": "D", "8": "B", "9": "D", "10": "C", "11": "B", "12": "A", "13": "A", "14": "B", "15": "B", "16": "C"}, # 16/16 đúng -> 4.0đ
            "p2_answers": {"1": {"a": "S", "b": "Đ", "c": "Đ", "d": "S"}, "2": {"a": "S", "b": "Đ", "c": "Đ", "d": "S"}, "3": {"a": "Đ", "b": "Đ", "c": "S", "d": "S"}, "4": {"a": "Đ", "b": "S", "c": "S", "d": "Đ"}}, # 4/4 đúng -> 4.0đ
            "p3_score": 1.8, "notes": "Bài làm xuất sắc"
        },
        {
            "stt": 2, "sbd": "720142", "name": "Trần Thị Mai", "code": "302", # Làm đề 302
            "p1_answers": {"1": "A", "2": "C", "3": "B", "4": "D", "5": "A", "6": "B", "7": "C", "8": "D", "9": "B", "10": "A", "11": "D", "12": "C", "13": "B", "14": "A", "15": "C", "16": "D"}, # 16/16 đúng đề 302 -> 4.0đ
            "p2_answers": {"1": {"a": "Đ", "b": "S", "c": "Đ", "d": "S"}, "2": {"a": "Đ", "b": "Đ", "c": "S", "d": "S"}, "3": {"a": "S", "b": "Đ", "c": "Đ", "d": "S"}, "4": {"a": "S", "b": "S", "c": "Đ", "d": "Đ"}},
            "p3_score": 1.5, "notes": "Trình bày rõ ràng"
        },
        {
            "stt": 3, "sbd": "720143", "name": "Lê Hoàng Long", "code": "301",
            "p1_answers": {"1": "B", "2": "A", "3": "C", "4": "B", "5": "D", "6": "C", "7": "D", "8": "B", "9": "D", "10": "C", "11": "B", "12": "A", "13": "A", "14": "B", "15": "B", "16": "C"}, # 14/16 -> 3.5đ
            "p2_answers": {"1": {"a": "S", "b": "Đ", "c": "Đ", "d": "S"}, "2": {"a": "S", "b": "Đ", "c": "Đ", "d": "S"}, "3": {"a": "Đ", "b": "Đ", "c": "S", "d": "Đ"}, "4": {"a": "Đ", "b": "S", "c": "S", "d": "Đ"}}, # C3 đúng 3 ý -> 0.5đ (tổng 3.5đ)
            "p3_score": 1.2, "notes": "Cần cẩn thận hơn câu 3 Đ/S"
        },
        {
            "stt": 4, "sbd": "720144", "name": "Phạm Minh Tuấn", "code": "302",
            "p1_answers": {"1": "A", "2": "C", "3": "B", "4": "D", "5": "A", "6": "B", "7": "C", "8": "A", "9": "B", "10": "A", "11": "D", "12": "C", "13": "B", "14": "A", "15": "C", "16": "D"},
            "p2_answers": {"1": {"a": "Đ", "b": "S", "c": "Đ", "d": "S"}, "2": {"a": "Đ", "b": "Đ", "c": "S", "d": "S"}, "3": {"a": "S", "b": "Đ", "c": "S", "d": "S"}, "4": {"a": "S", "b": "S", "c": "Đ", "d": "Đ"}},
            "p3_score": 1.4, "notes": "Đạt yêu cầu"
        },
        {
            "stt": 5, "sbd": "720145", "name": "Đỗ Thu Hương", "code": "301",
            "p1_answers": {"1": "B", "2": "A", "3": "C", "4": "A", "5": "D", "6": "D", "7": "D", "8": "B", "9": "D", "10": "C", "11": "B", "12": "B", "13": "A", "14": "B", "15": "B", "16": "C"},
            "p2_answers": {"1": {"a": "S", "b": "Đ", "c": "Đ", "d": "S"}, "2": {"a": "S", "b": "Đ", "c": "Đ", "d": "S"}, "3": {"a": "Đ", "b": "Đ", "c": "S", "d": "Đ"}, "4": {"a": "Đ", "b": "S", "c": "S", "d": "Đ"}},
            "p3_score": 1.6, "notes": "Bài viết tự luận tốt"
        }
    ]

    results_for_excel = []

    print("\n[*] Tiến hành chấm điểm từng học sinh theo cơ chế xoay vòng Key Pool:")
    for stud in mock_students:
        # Lấy key xoay vòng
        key_str, key_idx = pool.get_available_key()
        pool.record_success(key_idx)
        
        # Tự động chọn barem theo mã đề của học sinh
        exam_code = stud["code"]
        barem = master_db.get(exam_code, master_db["301"])

        # Tính điểm Phần I & II
        res_p1 = score_part_1(stud["p1_answers"], barem["part_1_mcq"])
        res_p2 = score_part_2(stud["p2_answers"], barem["part_2_true_false"])
        p3_score = stud["p3_score"]
        total_score = round(res_p1["total_score"] + res_p2["total_score"] + p3_score, 2)

        print(f" -> [Bài {stud['stt']}/5] SBD: {stud['sbd']} | HS: {stud['name']} | Mã đề: {exam_code} | Key dùng: {pool.keys[key_idx]['masked']} | Tổng điểm: {total_score}/10.0")

        results_for_excel.append({
            "STT": stud["stt"],
            "Số Báo Danh": stud["sbd"],
            "Họ và Tên": stud["name"],
            "Mã Đề Thi": exam_code,
            "Điểm Phần I (Trắc nghiệm)": res_p1["total_score"],
            "Điểm Phần II (Đúng/Sai)": res_p2["total_score"],
            "Điểm Phần III (Tự luận)": p3_score,
            "TỔNG ĐIỂM": total_score,
            "Xếp Loại": "Giỏi" if total_score >= 8.0 else ("Khá" if total_score >= 6.5 else "Trung bình"),
            "Ghi Chú": stud["notes"]
        })

    # 4. Xuất file Excel bảng điểm
    excel_filename = "bang_diem_lop_11_demo.xlsx"
    df = pd.DataFrame(results_for_excel)
    df.to_excel(excel_filename, index=False, engine="openpyxl")
    print(f"\n[+] ĐÃ XUẤT THÀNH CÔNG BẢNG ĐIỂM CẢ LỚP RA FILE: {excel_filename}")

    # 5. Thống kê tình trạng Key Pool sau khi chấm
    print("\n[*] Thống kê lượt xoay vòng của Gemini Key Pool:")
    stats_df = pd.DataFrame(pool.get_stats())
    print(stats_df.to_string(index=False))
    print("=" * 65)

if __name__ == "__main__":
    run_batch_simulation()
