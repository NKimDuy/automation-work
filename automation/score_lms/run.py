# =============================================================================
# MODULE: automation/score_lms/run.py
# MỤC ĐÍCH: Tự động chấm điểm chất lượng soạn thảo LMS theo các tiêu chí
#           (cấu trúc môn học, bài giảng, bài tập, diễn đàn, tỷ lệ truy cập LSA...).
#
# LUỒNG CHẠY CHÍNH (score_lms):
#   1. Load model phân loại văn bản (model_predict) và mở file Excel danh sách môn
#   2. Cào dữ liệu LSA (số SV đăng ký/truy cập) 1 lần trước khi vào vòng lặp
#   3. Đăng nhập LMS qua SSO
#   4. Với mỗi môn học trong file Excel:
#      a. Tìm môn trên LMS theo mã môn - phòng - nhóm
#      b. Lấy text trang môn học, tiền xử lý rồi đưa vào model để dự đoán tiêu chí
#      c. Chấm điểm theo từng tiêu chí (cấu trúc, bài tập, diễn đàn, % truy cập LSA)
#      d. Ghi kết quả + ghi chú vào file Excel, lưu lại sau mỗi môn
#
# CHẠY TRỰC TIẾP:
#   python -m automation.score_lms.run
# =============================================================================

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from modules.init_selenium import InitSelenium
# from db.connect_mysql import get_connection
from config import settings
from config import paths
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
import os
import time
import re
import copy
import joblib

class ScoreLMS:
    SCORE_CRITERIA = {
        "lms_done": {"score": "", "name": "Thực hiện LMS"},
        "structure": {"score": 0, "name": "Cấu trúc đầy đủ"},
        "course_info": {"score": 0, "name": "Thông tin về khóa học"},
        "objectives": {"score": 0, "name": "Mục tiêu, hoạt động học tập"},
        "lecture_content": {"score": 0, "name": "Thông tin bài giảng"},
        "midterm_grades": {"score": 0, "name": "Bảng điểm giữa kỳ"},
        "lecture_video": {"score": "", "name": "Video ghi hình buổi giảng"},
        "references": {"score": "", "name": "Tài liệu tham khảo/bài đọc thêm"},
        "assignments": {"score": 0, "name": "Bài tập tự luận/trắc nghiệm"},
        "forum": {"score": "", "name": "Diễn đàn"},
        "qa_video": {"score": "", "name": "video giải đáp thắc mắc cho sinh viên"},
        "students_enrolled": {"score": 0, "name": "Sinh viên đăng ký"},
        "students_accessed": {"score": 0, "name": "Sinh viên truy cập"},
        "student_access_rate": {"score": 0, "name": "Tỷ lệ sinh viên truy cập"},
        "teacher_access_rate": {"score": 0, "name": "Tỷ lệ giảng viên truy cập"},
        "total_score": {"score": 0, "name": "Tổng điểm"},
        "evaluation_result": {"score": "", "name": "kết quả đánh giá"},
    }

    # nhãn tiêu chí (1-10) mà model dự đoán trả về, dùng để dò xem 1 dòng text nói về tiêu chí nào
    CRITERIA_LABELS = {
        0: "Nhiễu",
        1: "Giới thiệu môn học",
        2: "Giới thiệu giảng viên",
        3: "Đề cương môn học",
        4: "Phương thức kiểm tra, đánh giá môn học",
        5: "Tỷ lệ điểm đánh giá",
        6: "Giới thiệu ngắn về mục tiêu, hoạt động học tập",
        7: "Nội dung bài giảng",
        8: "Tài liệu tham khảo/bài đọc thêm",
        9: "Video giảng online",
        10: "Bảng điểm giữa kỳ",
    }

    NOT_RATED = [
        "Thí nghiệm vật liệu xây dựng",
        "Thí nghiệm cơ chất lỏng",
        "Thí nghiệm sức bền vật liệu",
        "Thực tập trắc địa",
        "Địa chất công trình + thực tập",
        "Thí nghiệm cơ học đất",
        "Nhận thức ngành",
        'Thực tập địa chất công trình',
        "Đồ án nền móng",
        "Đồ án kết cấu bê tông cốt thép 1",
        "Đồ án kết cấu bê tông cốt thép 2",
        "Đồ án bê tông 1",
        "Đồ án bê tông 2",
        "Đồ án thi công",
        "Đồ án kết cấu thép",
        "Đồ án cấp thoát nước trong nhà",
        "Đồ án cấp thoát nước",
        "Đồ án mạng lưới cấp thoát nước",
        "Đồ án tổ chức và quản lý thi công",
        "Đồ án quản lý dự án xây dựng",
        "Đồ án lập và thảm định dự án ĐTXD",
        "Đồ án phân tích định lượng trong QLXD",
        "Thực tập tốt nghiệp",
        "Đồ án tốt nghiệp",
        "Thực tập tốt nghiệp - CNKTCTXD",
        "Thực tập tốt nghiệp - QLXD"
    ] 
    # các từ khóa nhiễu (menu, nút bấm...) cần loại khỏi text trước khi đưa vào model dự đoán
    NOISE_KEYWORDS = [
        'rút gọn',
        'các thông báo',
        'Bải tập',
        'Bài tập',
        'Câu hỏi',
        'câu hỏi thảo luận',
        'diễn đàn',
        'Chuyển tới nội dung chính',
        'LMS ĐTTX',
        'TX',
        'Opened',
        'Due',
        'Topic outline',
        'Thu gọn toàn bộ',
        'Buổi học trực tuyến',
        'New subsection',
        'Copyright © 2024. Powered by Moodle',
        'bạn đang đăng nhập',
        'được cung cấp',
        ' contact site support',
        'data retention summary',
        'không hiện hữu trừ khi',
        'có sẵn nhưng không nhìn thấy  trong trang khóa học',
        'trợ giúp và tài liệu tham khảo',
        'các nội dung chính trong',
        'available from',
        'trắc nghiệm',
        'phân biệt',
        'ngày học',
        'thời gian mở',
        'ẩn đối với người',
        'nhận xét học',
        'show more',
        'múi giờ',
        'text book',
        'trang chủ',
        'bảng điều khiển',
        'discussion board',
        'giám đốc',
        'nhân viên',
        'đào tạo',
        'công việc',
        'bố',
        'mẹ',
        'cậu',
        'dì',
        'bác',
        'ông',
        'bà',
        'chú',
        'cháu'
    ]

    # các cột bắt buộc phải có trong file Excel danh sách môn học, nếu thiếu sẽ không tiến hành đánh giá
    REQUIRED_COLUMNS = [
        "Mã MH",
        "Tên môn học",
        "Nhóm",
        "CBGD",
        "Tên CBGD"
    ]
  
    def __init__(self, model_file=None, excel_file=None):
        self.excel_file = excel_file or paths.score_lms_input_file_dir(settings.semester_score_lms, settings.score_lms_input_filename)
        self.model_file = model_file or paths.model_file_dir(settings.model_filename)
        self.link_data_ai = paths.harvested_file_dir(settings.semester_score_lms, settings.harvested_filename)
        self.training_set_file = paths.training_set_file_dir(settings.training_set_filename)
        self.existing_training_texts = self._load_existing_training_texts()
        self.url_lms = f"https://lms.oude.edu.vn/{settings.semester_score_lms}/course/search.php"
        # tên key đặt bằng tiếng Anh cho dễ nhớ/dễ gõ, "name" là nhãn tiếng Việt hiển thị ra cột Excel

    def preprocess(self, text: str) -> str:
        if not isinstance(text, str):
            return ""
        text = text.lower().strip()
        # Xoá ký tự đặc biệt, giữ chữ cái và khoảng trắng
        text = re.sub(r'^\d+[\.\)\-]\s*', '', text)   # xóa số thứ tự đầu câu: "4." "4)" "4-"
        text = re.sub(r'[^\w\s]', ' ', text)           # thay ký tự đặc biệt bằng khoảng trắng
        text = re.sub(r'\s+', ' ', text).strip()       # gộp khoảng trắng
        return text
    
    def _prepare_lines_for_prediction(self, driver):
        """Lấy toàn bộ text của trang môn học, bỏ dòng nhiễu (NOISE_KEYWORDS) rồi tách theo tiêu chí
        (CRITERIA_LABELS) để model dự đoán chính xác hơn. Đồng thời lưu lại các dòng này để sau này train lại model."""
        lines = []

        # 1. Tiêu đề lớn (section) — class bắt đầu bằng "sectionid"
        for el in driver.find_elements(By.CSS_SELECTOR, "[class*='sectionid']"):
            t = " ".join(el.text.split())
            if t:
                lines.append(t)

        # 2. Tên chương / tiêu đề tài nguyên
        for el in driver.find_elements(By.CSS_SELECTOR, ".activityname"):
            t = " ".join(el.text.split())
            if t:
                lines.append(t)

        # 3. Đoạn văn mô tả
        for el in driver.find_elements(By.CSS_SELECTOR, ".no-overflow"):
            t = " ".join(el.text.split())
            if t:
                lines.append(t)

        # Bỏ trùng, giữ thứ tự
        lines = list(dict.fromkeys(lines))

        # loại bỏ các dòng có 1 từ hoặc không có từ nào, vì những dòng này thường không mang nhiều thông tin
        lines = [item for item in lines if len(item.split()) > 1]

        # loại bỏ các dòng nhiễu (menu, nút bấm...) và tách các dòng có chứa tên tiêu chí ra thành nhiều dòng riêng
        filtered_lines = []
        for line in lines:
            has_noise = any(noise.lower() in line.lower() for noise in self.NOISE_KEYWORDS)
            if has_noise:
                continue

            # nếu dòng có chứa tên 1 trong các tiêu chí (giáo viên hay để chung nhiều tiêu chí trên 1 dòng,
            # cách nhau bằng dấu phẩy) thì tách riêng từng tiêu chí ra cho model dễ nhận diện hơn
            matches_criteria_label = any(label.lower() in line.lower() for label in self.CRITERIA_LABELS.values())
            if matches_criteria_label:
                for part in line.split(','):
                    filtered_lines.append(part.strip())
            else:
                filtered_lines.append(line)
        lines = filtered_lines

        lines = [self.preprocess(t) for t in lines]

        self._save_lines_for_training(lines)
        return lines

    def predict(self, model, texts: list) -> list:
        return list(model.predict(texts))

    def check_forum(self, driver, forum_link, teacher_name):
        subject_tab = driver.current_window_handle  # tab môn học, để quay lại sau khi đóng tab diễn đàn
        driver.execute_script("window.open(arguments[0]);", forum_link)
        driver.switch_to.window(driver.window_handles[-1])

        rows = driver.find_elements(By.CSS_SELECTOR, "tr.discussion")
        flag_forum = False
        if rows:
            for row in rows:
                try:
                    number_responses = row.find_element(By.CSS_SELECTOR, "td.text-center span").text.strip()

                    if int(number_responses) == 1:
                        try:
                            reply_author_name = row.find_element(By.CSS_SELECTOR, "td.text-start .mb-1").text.strip()
                        except Exception as e:
                            print(f"Không tìm thấy tên người trả lời trong diễn đàn: {e}")
                            reply_author_name = "Không có"

                        if teacher_name.lower() in reply_author_name.lower():
                            flag_forum = True
                            break
                    elif int(number_responses) > 1:
                        try:
                            forum_tab = driver.current_window_handle  # tab diễn đàn, để quay lại sau khi đóng tab topic

                            link_topic = row.find_element(By.CSS_SELECTOR, "th.topic a")
                            link_topic = link_topic.get_attribute("href")

                            driver.execute_script("window.open(arguments[0]);", link_topic)
                            driver.switch_to.window(driver.window_handles[-1])

                            try:
                                commenter_names = []
                                try:
                                    commenter_name_elements = driver.find_elements(By.CSS_SELECTOR, "header.mb-2 div.mb-3")
                                except Exception as e:
                                    print(f"Không tìm thấy phần tử chứa tên người bình luận: {e}")
                                    commenter_name_elements = []
                                for element in commenter_name_elements:
                                    try:
                                        commenter_name = element.find_element(By.CSS_SELECTOR, "a").text.strip()
                                        commenter_names.append(commenter_name.lower())
                                    except Exception as e:
                                        print(f"Không tìm thấy tên người bình luận trong diễn đàn: {e}")

                                driver.close()
                                driver.switch_to.window(forum_tab)
                                if commenter_names.count(teacher_name.lower()) > 1:
                                    flag_forum = True
                                    break
                            except Exception as e:
                                print(f"Không tìm thấy thông tin người bình luận: {e}")

                        except Exception as e:
                            print(f"Không tìm thấy link topic trong diễn đàn: {e}")

                except Exception as e:
                    print(f"Không tìm thấy số lượng phản hồi trong diễn đàn: {e}")

        driver.close()
        driver.switch_to.window(subject_tab)
        return flag_forum

    def _load_existing_training_texts(self):
        wb = load_workbook(self.training_set_file, read_only=True)
        ws = wb.active
        header_row = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        col_idx = header_row.index("tieu_chi")

        existing = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[col_idx]:
                existing.add(row[col_idx])
        return existing

    def _save_lines_for_training(self, lines):
        """Ghi thêm các dòng text đã lọc vào file excel dùng để train lại model sau này."""
        if os.path.exists(self.link_data_ai):
            wb = load_workbook(self.link_data_ai)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.append(["Dữ liệu"])

        for line in lines:
            ws.append([line])
            self.existing_training_texts.add(line)   # đánh dấu đã ghi, tránh trùng lặp giữa các môn trong cùng lượt chạy
        wb.save(self.link_data_ai)

    def _score_structure_criteria(self, score_result, results):
        """Chấm điểm các tiêu chí cấu trúc môn học dựa trên nhãn (1-10) mà model dự đoán được.
        Trả về (notes, sum_score) để cộng dồn ở hàm gọi."""
        notes = ''
        sum_score = 0

        # id 1-5 trong CRITERIA_LABELS là các tiêu chí bắt buộc phải có đủ để đạt "course_info"
        required_course_info_ids = [1, 2, 3, 4, 5]
        missing_course_info_ids = [x for x in required_course_info_ids if x not in set(results)]
        if not missing_course_info_ids:
            score_result["course_info"]["score"] = 10
            sum_score += 10
        else:
            missing_course_info_names = ["thiếu " + self.CRITERIA_LABELS[i] for i in missing_course_info_ids]
            notes += f"\nChưa đủ thông tin khóa học, cụ thể ({','.join(missing_course_info_names)})"

        if 6 in results:
            score_result["objectives"]["score"] = 20
            sum_score += 20
        else:
            notes += "\nThiếu giới thiệu về mục tiêu học tập"

        if 7 in results:
            score_result["lecture_content"]["score"] = 20
            sum_score += 20
        else:
            notes += "\nThiếu nội dung bài giảng"

        if 8 in results:
            score_result["references"]["score"] = "x"
        else:
            notes += "\nThiếu tài liệu tham khảo/bài đọc thêm"

        if 9 in results:
            score_result["lecture_video"]["score"] = "x"

        if 10 in results:
            score_result["midterm_grades"]["score"] = 15
            sum_score += 15
        else:
            notes += "\nThiếu bảng điểm giữa kỳ"

        check_score_structure = [
            score_result["course_info"]["score"] == 10,
            score_result["lecture_content"]["score"] == 20,
            score_result["midterm_grades"]["score"] == 15,
        ]

        score_result["structure"]["score"] = sum(check_score_structure) * 5
        sum_score += sum(check_score_structure) * 5

        return notes, sum_score

    def _craw_lsa_data(self, start_selenium):
        # tạo dictionary với key là mã môn học + mã nhóm + mã giảng viên, value là list chứa dữ liệu lsa
        lsa_data = {}
        get_lsa_data = start_selenium.process_get_detail_lsa()
        for lsa_row in get_lsa_data:
            if "subject_id" not in lsa_row:
                continue
            key = f"{lsa_row['subject_id']}_{lsa_row['group_id']}_{lsa_row['teacher_id']}"
            lsa_values = [lsa_row["count_total_student"], lsa_row["count_student_access"], lsa_row["percent_student_access"], lsa_row["percent_teacher_access"]]
            lsa_data[key] = lsa_values
        return lsa_data

    def _apply_lsa_percent(self, score_result, lsa_values):
        """Gán dữ liệu sinh viên đăng ký/truy cập lấy từ LSA vào điểm số.
        lsa_values theo đúng thứ tự [count_total_student, count_student_access, percent_student_access, percent_teacher_access]
        (xem nơi tạo ra list này ở score_lms(), ngay trước khi gọi get_results_score()). Nếu không có dữ liệu LSA
        (lsa_values rỗng/None) thì đánh dấu "lỗi" để hàm đánh giá kết quả cuối biết mà xử lý riêng."""
        score_result["students_enrolled"]["score"] = lsa_values[0] if lsa_values else "lỗi"
        score_result["students_accessed"]["score"] = lsa_values[1] if lsa_values else "lỗi"
        score_result["student_access_rate"]["score"] = lsa_values[2] if lsa_values else "lỗi"
        score_result["teacher_access_rate"]["score"] = lsa_values[3] if lsa_values else "lỗi"

    def _extract_forum_quiz_assign_links(self, driver):
        """Duyệt toàn bộ link trong trang môn học, phân loại theo forum/quiz/assign dựa trên href."""
        forums, quiz, assign = {}, {}, {}
        try:
            all_links = driver.find_elements(By.TAG_NAME, "a")
            for a in all_links:
                try:
                    href = a.get_attribute("href")
                    text = a.text.strip()
                except Exception as e:
                    print(f"Không đọc được thuộc tính của link: {e}")
                    continue
                if not href or not text:
                    continue
                if "forum" in href:
                    forums[text] = href
                if "quiz" in href:
                    quiz[text] = href
                if "assign" in href:
                    assign[text] = href
        except Exception as e:
            print(f"lỗi khi lấy link diễn đàn, bài tập: {e}")
        return forums, quiz, assign

    def _score_assignments(self, score_result, quiz, assign):
        """Chấm điểm tiêu chí bài tập: có quiz hoặc bài tập tự luận (assign) là đạt.
        Trả về (notes, sum_score) để cộng dồn ở hàm gọi."""
        notes = ''
        sum_score = 0
        try:
            assignment_types_found = [name for name, link_dict in [("quiz", quiz), ("assign", assign)] if link_dict]
            if assignment_types_found:
                score_result["assignments"]["score"] = 20
                sum_score += 20
            else:
                notes += "\nThiếu bài tập"
        except Exception as e:
            print(f"lỗi khi chấm điểm bài tập: {e}")
        return notes, sum_score

    def _score_forum_interaction(self, driver, score_result, forums, teacher_name):
        """Chấm điểm tiêu chí diễn đàn: mở từng diễn đàn kiểm tra giảng viên đã tương tác với sinh viên chưa.
        Chỉ cần 1 diễn đàn có tương tác là đạt tiêu chí nên dừng ngay khi tìm thấy (khác bản cũ là mở hết
        tất cả diễn đàn rồi mới xét — kết quả chấm điểm giống nhau nhưng chạy nhanh hơn khi có nhiều diễn đàn)."""
        if not forums:
            return "\nKhông có diễn đàn"

        has_teacher_interaction = False
        for name, forum_link in forums.items():
            try:
                if self.check_forum(driver, forum_link, teacher_name):
                    has_teacher_interaction = True
                    break
            except Exception as e:
                print(f"lỗi khi kiểm tra diễn đàn '{name}': {e}")

        if has_teacher_interaction:
            score_result["forum"]["score"] = "x"
            return ""
        return "\nKhông có diễn đàn tương tác với sinh viên"

    def _percent_to_float(self, value):
        """Đổi chuỗi phần trăm kiểu "90%" (dữ liệu LSA) thành số thực để so sánh."""
        return float(str(value).replace('%', ''))

    def _evaluate_final_result(self, score_result, sum_score):
        """Tính kết quả đánh giá cuối cùng (evaluation_result): Không đạt / Đáp ứng cơ bản / Đáp ứng nâng cao,
        dựa trên các tiêu chí cơ bản/nâng cao và % sinh viên, giảng viên truy cập lấy từ LSA.
        Trả về notes bổ sung (vd lý do không đạt)."""
        notes = ''
        try:
            score_basic = [
                score_result["structure"]["score"] == 15,
                score_result["course_info"]["score"] == 10,
                score_result["objectives"]["score"] == 20,
                score_result["lecture_content"]["score"] == 20,
                score_result["midterm_grades"]["score"] == 15,
                sum_score >= 50
            ]

            student_access_rate = score_result["student_access_rate"]["score"]
            teacher_access_rate = score_result["teacher_access_rate"]["score"]

            if all(score_basic):
                if student_access_rate == "lỗi" or teacher_access_rate == "lỗi":
                    score_result["evaluation_result"]["score"] = "Thiếu % để đáp ứng cơ bản"
                elif self._percent_to_float(student_access_rate) >= 70 and self._percent_to_float(teacher_access_rate) > 0:
                    score_result["evaluation_result"]["score"] = "Đáp ứng cơ bản"

                    score_advance = [
                        score_result["forum"]["score"] == "x",
                        score_result["assignments"]["score"] == 20,
                        sum_score >= 80
                    ]
                    if all(score_advance):
                        if student_access_rate == "lỗi" or teacher_access_rate == "lỗi":
                            score_result["evaluation_result"]["score"] = "Thiếu % để đáp ứng nâng cao"
                        elif self._percent_to_float(student_access_rate) >= 90 and self._percent_to_float(teacher_access_rate) > 0:
                            score_result["evaluation_result"]["score"] = "Đáp ứng nâng cao"
                else:
                    if self._percent_to_float(student_access_rate) < 70:
                        score_result["evaluation_result"]["score"] = "Không đạt vì tỷ lệ sinh viên truy cập dưới 70%"
                        notes += "\nTỷ lệ sinh viên truy cập dưới 70%"
                    elif self._percent_to_float(teacher_access_rate) == 0:
                        score_result["evaluation_result"]["score"] = "Không đạt vì giảng viên không truy cập"
                        notes += "\nGiảng viên không truy cập vào LMS"
            else:
                score_result["evaluation_result"]["score"] = "Không đạt"

        except Exception as e:
            print(f"Bị lỗi khi lấy phần trăm giảng viên và sinh viên lsa: {e}")

        return notes

    def get_results_score(self, driver, model, teacher_name, lsa_values):
        """Chấm điểm 1 môn học đang mở trên `driver` theo các tiêu chí trong SCORE_CRITERIA.
        Trả về [score_result, notes]."""
        notes = ''  # lưu các tiêu chí chưa đạt để ghi chú vào file excel
        sum_score = 0  # lưu tổng điểm của môn học, sẽ được ghi vào file excel ở cột "Tổng điểm"
        # copy từ điển điểm để áp dụng cho từng môn học, tránh bị ghi đè khi duyệt qua nhiều môn
        score_result = copy.deepcopy(self.SCORE_CRITERIA)
        score_result["lms_done"]["score"] = "x"

        try:
            lines = self._prepare_lines_for_prediction(driver)
            forums, quiz, assign = self._extract_forum_quiz_assign_links(driver)

            # xóa các text của bài tập, diễn đàn để không làm ảnh hưởng đến kết quả dự đoán
            remove_forums_assignments = set(forums.keys()).union(set(quiz.keys())).union(set(assign.keys()))
            lines = [line for line in lines if line not in remove_forums_assignments]

            results = self.predict(model, lines)
            results = list(set(results))  # lấy kết quả duy nhất để đánh giá

            notes_struct, sum_struct = self._score_structure_criteria(score_result, results)
            notes += notes_struct
            sum_score += sum_struct

            # khúc này sẽ dùng dữ liệu sinh viên đăng ký/truy cập lấy từ lsa,
            # để tính ra tỷ lệ sinh viên/giảng viên truy cập, phục vụ cho _evaluate_final_result bên dưới
            self._apply_lsa_percent(score_result, lsa_values)

            notes_bt, sum_bt = self._score_assignments(score_result, quiz, assign)
            notes += notes_bt
            sum_score += sum_bt

            score_result["total_score"]["score"] = sum_score

            notes += self._score_forum_interaction(driver, score_result, forums, teacher_name)

            notes += self._evaluate_final_result(score_result, sum_score)

        except Exception as e:
            score_result["evaluation_result"]["score"] = "Không đạt"
            notes += "\nKhông đủ thông tin để đánh giá"
            print(f"môn học không có dữ kiện để đánh giá: {e}")

        return [score_result, notes]

    @staticmethod
    def validate_columns(excel_file):
        """Kiểm tra file Excel có đủ các cột bắt buộc (REQUIRED_COLUMNS) không.
        Trả về (column_index, missing_columns):
        - column_index: dict {tên cột: vị trí cột}
        - missing_columns: list tên cột bị thiếu (rỗng nếu file hợp lệ)
        """
        wb = load_workbook(excel_file, read_only=True)
        ws = wb.active
        header_row = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        column_index = {name: idx for idx, name in enumerate(header_row)}
        missing_columns = [col for col in ScoreLMS.REQUIRED_COLUMNS if col not in column_index]
        return column_index, missing_columns

    def score_lms(self, lsa_data=None): 
        model = joblib.load(self.model_file)

        wb_score = load_workbook(self.excel_file)
        ws_score = wb_score.active
        column_current = ws_score.max_column
        col_start = column_current + 1                              # cột bắt đầu thêm các cột đánh giá
        col_result = column_current + len(self.SCORE_CRITERIA)   # cột "kết quả đánh giá" — dùng để ghi "Hoàn thành"/"Chưa soạn LMS"
        col_done = col_result + 2                                  # cột đánh dấu "x" — biết dòng nào đã xử lý xong để resume

        # lần đầu kiểm tra trong file đã có các cột để tiến hành đánh giá chưa
        # nếu chưa có thì sẽ thêm các cột này vào file,
        # nếu tiếp tục chạy thì sẽ không thêm nữa, dữ liệu cứ thế được ghi tiếp, mà không có tiêu đề cột
        column_index, missing_columns = self.validate_columns(self.excel_file)

        # nếu thiếu cột, thì sẽ không tiến hành đánh giá
        if missing_columns: 
            print("File thiếu các cột: {}".format(", ".join(missing_columns)))
        else:
            all_done = all(
                ws_score.cell(row=i, column=col_done).value == "x"
                for i in range(2, ws_score.max_row + 1)
            )
            if all_done:
                print("Tất cả các môn học đã được đánh giá. Không cần chạy lại.")
            else:
                if self.SCORE_CRITERIA["lms_done"]["name"] not in column_index:
                    column_add = col_start
                    for criteria_key in self.SCORE_CRITERIA:
                        ws_score.cell(row=1, column=column_add).value = self.SCORE_CRITERIA[criteria_key]["name"]
                        column_add += 1
                    ws_score.cell(row=1, column=column_add).value = "Ghi chú"
        
                # khởi tạo selenium để đăng nhập LMS và duyệt qua các môn học
                start_selenium = InitSelenium()

                # lấy dữ liệu lsa từ hàm _craw_lsa_dat() nếu none
                # muốn test thì truyền vào lsa_data = {} để bỏ qua bước cào dữ liệu lsa, 
                # nhưng sẽ không có dữ liệu sinh viên/giảng viên truy cập
                if lsa_data is None:
                    print("Đang cào dữ liệu LSA từ hệ thống...")
                    lsa_data = self._craw_lsa_data(start_selenium)

                driver = start_selenium.login_selenium(self.url_lms)
                #chỗ này sẽ chạy vòng lặp for để duyệt qua các môn học trong file excel
                print("Đang đánh giá các môn học trên LMS...")
                for i, row in enumerate(ws_score.iter_rows(min_row=2, values_only=True), start=2):
                    if ws_score.cell(row=i, column=col_done).value != "x":
                        # mỗi lần duyệt qua một môn học thì sẽ mở lại trang tìm kiếm môn học trên LMS
                        # để đảm bảo rằng các bước tìm kiếm và đánh giá được thực hiện chính xác cho từng môn học,
                        # tránh bị lỗi do trang web bị thay đổi khi duyệt qua nhiều môn học
                        driver.get(self.url_lms)

                        subject_id = row[column_index["Mã MH"]]
                        subject_name = row[column_index["Tên môn học"]]
                        teacher_id = row[column_index["CBGD"]]
                        teacher_name = row[column_index["Tên CBGD"]]
                        group_id = row[column_index["Nhóm"]]

                        key_lsa = f"{subject_id}_{group_id}_{teacher_id}"
                        search_subject = f"{subject_id} - ({teacher_id}-{group_id})"

                        # kiểm tra xem môn học có thuộc danh sách các môn không đánh giá điểm LMS của khoa xây dựng hay không,
                        # nếu có thì sẽ ghi "Hoàn thành" vào cột điểm số và bỏ qua các bước kiểm tra tiếp theo,
                        # vì những môn này chỉ cần soạn LMS hoàn chỉnh là được, không yêu cầu phải có đầy đủ các tiêu chí như các môn khác
                        if subject_name.lower() in [subject.lower() for subject in self.NOT_RATED]:
                            ws_score.cell(row=i, column=col_result).value = "Hoàn thành"
                            ws_score.cell(row=i, column=col_done).value = "x"
                            print(f"[Hoàn thành - Khoa xây dựng]: {subject_name} - {teacher_name} - {group_id}")
                            continue

                        try:
                            search_input = WebDriverWait(driver, 10).until(
                                EC.visibility_of_element_located((By.NAME, "q"))
                            )
                            search_input.clear()
                            search_input.send_keys(search_subject) # BADM1391 - (QT653-TN120)
                        except Exception as e:
                                print(f"Không tìm thấy ô nhập môn học, lỗi {e}")

                        try:
                            # Nhấn nút tìm kiếm
                            button_find_text = "button[type='submit']"
                            button_find = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, button_find_text))
                            )
                            button_find.click()
                        except Exception as e:
                                print(f"Không tìm thấy nút tìm môn học, lỗi {e}")

                        try:
                            # Lấy link môn học đầu tiên trong kết quả tìm kiếm (class="aalink")
                            subject_link_element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CLASS_NAME, "aalink"))
                            )
                            href_link_subject = subject_link_element.get_attribute("href")
                            # mở link môn học để kiểm tra các tiêu chí đánh giá cấu trúc môn học trên LMS
                            driver.get(href_link_subject)

                            # gọi hàm get_results_score để đánh giá các tiêu chí và lấy điểm số áp dụng cho từng tiêu chí,
                            # cũng như ghi chú các tiêu chí chưa đạt vào file excel
                            try:
                                lsa_values = lsa_data.get(key_lsa)
                                score_result, notes = self.get_results_score(driver, model, teacher_name, lsa_values)
                                print(f"{score_result['evaluation_result']['score']}: {subject_name} - {teacher_name} - {group_id}")
                            except Exception as e:
                                print(f"lỗi ở hàm score_lms: {e}")

                            col_idx = col_start
                            # score_result là dictionary — ghi điểm từng tiêu chí vào cột tương ứng,
                            # thứ tự cột phải khớp thứ tự key trong SCORE_CRITERIA (dict giữ thứ tự chèn)
                            for criteria_key in score_result:
                                ws_score.cell(row=i, column=col_idx).value = score_result[criteria_key]["score"]
                                col_idx += 1
                            ws_score.cell(row=i, column=col_idx).value = notes
                            ws_score.cell(row=i, column=col_idx).alignment = Alignment(wrap_text=True)

                            driver.get(self.url_lms) # quay lại trang tìm kiếm để tiếp tục duyệt qua môn học tiếp theo

                        except Exception as e:
                            # nếu không tìm thấy môn học trên LMS thì sẽ ghi "Chưa soạn LMS" vào cột điểm số và ghi chú vào file excel, sau đó tiếp tục chạy vòng lặp để duyệt qua môn học tiếp theo
                            ws_score.cell(row=i, column=col_result).value = "Chưa soạn LMS"
                            print(f"[Chưa soạn LMS]: {subject_name} - {teacher_name} - {group_id}")
                        ws_score.cell(row=i, column=col_done).value = "x"
                        wb_score.save(self.excel_file)
                driver.quit()
                time.sleep(2)  # Thêm delay để dừng lại 2 giây trước khi nhập thông tin môn

if __name__ == "__main__":
    test = ScoreLMS()
    test.score_lms()