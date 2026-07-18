from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from modules.init_selenium import InitSelenium
# from db.connect_mysql import get_connection
from config.settings import semester
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
import os
import time
import re
import copy
import joblib


class ScoreLMS:
    def __init__(self):
        # self.url_lms = f"https://lms.oude.edu.vn/{semester}/course/search.php"
        self.base_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        self.url_lms = f"https://lms.oude.edu.vn/251/course/search.php"
        self.link_data_ai = os.path.join(self.base_dir, "AI", "data", "du_lieu_tim_duoc_sau_khi_danh_gia.xlsx")
        # tên key đặt bằng tiếng Anh cho dễ nhớ/dễ gõ, "name" là nhãn tiếng Việt hiển thị ra cột Excel
        self.SCORE_CRITERIA = {
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
        self.CRITERIA_LABELS = {
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
        self.NOT_RATED = [
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
            "Đồ án tốt nghiệp"
        ] 
        # các từ khóa nhiễu (menu, nút bấm...) cần loại khỏi text trước khi đưa vào model dự đoán
        self.NOISE_KEYWORDS = [
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
            'Copyright © 2024. Powered by Moodle'
        ]

    def preprocess(self, text: str) -> str:
        if not isinstance(text, str):
            return ""
        text = text.lower().strip()
        # Xoá ký tự đặc biệt, giữ chữ cái và khoảng trắng
        text = re.sub(r"[^\w\sàáâãèéêìíòóôõùúýăđơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợụủứừửữựỳỵỷỹ]", " ", text)
        text = re.sub(r"\s+", " ", text).strip()
        return text
    
    def predict(self, model, texts: list) -> list:
        cleaned = [self.preprocess(t) for t in texts]
        preds = model.predict(cleaned)
        results = list()
        for label in preds:
            results.append(label)
        return results
    
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

    def _prepare_lines_for_prediction(self, driver):
        """Lấy toàn bộ text của trang môn học, bỏ dòng nhiễu (NOISE_KEYWORDS) rồi tách theo tiêu chí
        (CRITERIA_LABELS) để model dự đoán chính xác hơn. Đồng thời lưu lại các dòng này để sau này train lại model."""
        full_text = driver.find_element(By.TAG_NAME, "body").text
        lines = full_text.splitlines()
        # loại bỏ các dòng có 1 từ hoặc không có từ nào, vì những dòng này thường không mang nhiều thông tin
        lines = [item for item in lines if len(item.split()) > 1]

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

        self._save_lines_for_training(lines)
        return lines

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
        wb.save(self.link_data_ai)

    def _score_structure_criteria(self, score_result, results):
        """Chấm điểm các tiêu chí cấu trúc môn học dựa trên nhãn (1-10) mà model dự đoán được.
        Trả về (notes, sum_score) để cộng dồn ở hàm gọi."""
        notes = ''
        sum_score = 0

        score_result["structure"]["score"] = 15
        sum_score += 15

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

        return notes, sum_score

    def _apply_lsa_percent(self, score_result, lsa_values):
        """Gán dữ liệu sinh viên đăng ký/truy cập lấy từ LSA vào điểm số.
        lsa_values theo đúng thứ tự [count_total_student, count_student_access, percent_student_access, percent_teacher_access]
        (xem nơi tạo ra list này ở score_lms()/test_lsa()). Nếu không có dữ liệu LSA (lsa_values rỗng) thì đánh dấu
        "lỗi" để hàm đánh giá kết quả cuối biết mà xử lý riêng."""
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

        lines = self._prepare_lines_for_prediction(driver)

        try:
            results = self.predict(model, lines)
            results = list(set(results))  # lấy kết quả duy nhất để đánh giá

            notes_struct, sum_struct = self._score_structure_criteria(score_result, results)
            notes += notes_struct
            sum_score += sum_struct

            # khúc này sẽ dùng dữ liệu sinh viên đăng ký/truy cập lấy từ lsa,
            # để tính ra tỷ lệ sinh viên/giảng viên truy cập, phục vụ cho _evaluate_final_result bên dưới
            self._apply_lsa_percent(score_result, lsa_values)

            forums, quiz, assign = self._extract_forum_quiz_assign_links(driver)

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
        

    def test_lsa(self):

        lsa_data = {}
        start_selenium = InitSelenium()
        get_lsa_data = start_selenium.process_get_detail_lsa()
        for lsa_row in get_lsa_data:
            if "subject_id" not in lsa_row:
                continue
            key = f"{lsa_row['subject_id']}_{lsa_row['group_id']}_{lsa_row['teacher_id']}"
            lsa_values = [lsa_row["count_total_student"], lsa_row["count_student_access"], lsa_row["percent_student_access"], lsa_row["percent_teacher_access"]]
            lsa_data[key] = lsa_values
        print(lsa_data)

        excel_file = os.path.join(self.base_dir, "data", "input", "score_lms","251","test.xlsx")
        wb_score = load_workbook(excel_file)
        ws_score = wb_score.active
        for i, row in enumerate(ws_score.iter_rows(min_row=2, values_only=True), start=2):
            subject_name = row[0]  # Giả sử tên môn học nằm ở cột A
            teacher_name = row[1]  # Giả sử tên giảng viên nằm ở cột B
            subject_id = row[2]    # Giả sử mã môn học nằm ở cột C
            group_id = row[3]      # Giả sử mã nhóm môn học nằm ở cột D
            teacher_id = row[4]    # Giả sử mã giảng viên nằm ở cột E

            key_lsa = f"{subject_id}_{group_id}_{teacher_id}"

            print(lsa_data.get(key_lsa))


    def score_lms(self): 
        
        model_file = os.path.join(self.base_dir, "AI", "model", "model_vi_classification_v2.pkl")
        model = joblib.load(model_file)

        # excel_file = os.path.join(self.base_dir, "data", "input", "score_lms","251","test2.xlsx")
        excel_file = os.path.join(self.base_dir, "data", "input", "score_lms","251","2025-06-04_251_email-nhom-mh-hk.xlsx")
        wb_score = load_workbook(excel_file)
        ws_score = wb_score.active

        # lần đầu kiểm tra trong file đã có các cột để tiến hành đánh giá chưa
        # Lấy cột thực hiện lms kiểm tra (lấy bất kỳ cột nào cũng được) nhưng lấy cột thực hiện lms
        # vì cột này sẽ luôn đứng sau các cột có sẵn trong file
        headers = [ws_score.cell(row=1, column=col).value for col in range(1, ws_score.max_column + 1)]
        column_current = ws_score.max_column
        if self.SCORE_CRITERIA["lms_done"]["name"] not in headers:
            column_add = column_current + 1
            for criteria_key in self.SCORE_CRITERIA:
                ws_score.cell(row=1, column=column_add).value = self.SCORE_CRITERIA[criteria_key]["name"]
                column_add += 1
            ws_score.cell(row=1, column=column_add).value = "Ghi chú"

        start_selenium = InitSelenium()
        driver = start_selenium.login_selenium(self.url_lms)

        # tạo dictionary với key là mã môn học + mã nhóm + mã giảng viên, value là list chứa dữ liệu lsa
        lsa_data = {}
        get_lsa_data = start_selenium.process_get_detail_lsa()
        for lsa_row in get_lsa_data:
            if "subject_id" not in lsa_row:
                continue
            key = f"{lsa_row['subject_id']}_{lsa_row['group_id']}_{lsa_row['teacher_id']}"
            lsa_values = [lsa_row["count_total_student"], lsa_row["count_student_access"], lsa_row["percent_student_access"], lsa_row["percent_teacher_access"]]
            lsa_data[key] = lsa_values

        # temp = [2,2,'90%', '90%']

        #chỗ này sẽ chạy vòng lặp for để duyệt qua các môn học trong file excel
        for i, row in enumerate(ws_score.iter_rows(min_row=2, values_only=True), start=2):
            if ws_score.cell(row=i, column=column_current + len(self.SCORE_CRITERIA) + 2).value != "x":
                # mỗi lần duyệt qua một môn học thì sẽ mở lại trang tìm kiếm môn học trên LMS
                # để đảm bảo rằng các bước tìm kiếm và đánh giá được thực hiện chính xác cho từng môn học,
                # tránh bị lỗi do trang web bị thay đổi khi duyệt qua nhiều môn học
                driver.get(self.url_lms)

                subject_id = row[0]
                subject_name = row[1]
                teacher_id = row[8]
                teacher_name = row[9]
                group_id = row[6]

                # phục vụ cho việc test dữ liệu
                # subject_id = row[2]
                # subject_name = row[0]
                # teacher_id = row[4]
                # teacher_name = row[1]
                # group_id = row[3]

                key_lsa = f"{subject_id}_{group_id}_{teacher_id}"
                search_subject = f"{subject_id} - ({teacher_id}-{group_id})"

                # kiểm tra xem môn học có thuộc danh sách các môn không đánh giá điểm LMS của khoa xây dựng hay không,
                # nếu có thì sẽ ghi "Hoàn thành" vào cột điểm số và bỏ qua các bước kiểm tra tiếp theo,
                # vì những môn này chỉ cần soạn LMS hoàn chỉnh là được, không yêu cầu phải có đầy đủ các tiêu chí như các môn khác
                if subject_name.lower() in [subject.lower() for subject in self.NOT_RATED]:
                    # print(f"✅ Môn học thuộc cách chấm chỉ hoàn thành của khoa xây dựng")
                    ws_score.cell(row=i, column=column_current + len(self.SCORE_CRITERIA)).value = "Hoàn thành"
                    ws_score.cell(row=i, column=column_current + len(self.SCORE_CRITERIA) + 2).value = "x"
                    print(f"[môn học của khoa xây dựng] kết quả xử lý môn {subject_name} - {teacher_name} - {group_id}: Hoàn thành")
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
                        score_result, notes = self.get_results_score(driver, model, teacher_name, lsa_data.get(key_lsa))
                        # score_result, notes = self.get_results_score(driver, model, teacher_name, temp)
                        print(f"[Đã có môn] kết quả xử lý môn {subject_name} - {teacher_name} - {group_id}: {score_result['evaluation_result']['score']}")
                        # print(score_result)
                        # print(score_result)
                    except Exception as e:
                        print(f"lỗi ở hàm score_lms: {e}")

                    col_idx = column_current + 1
                    # score_result sẽ trả về dictionary
                    # for criteria_key in score_result: sẽ duyệt qua các key của dictionary
                    for criteria_key in score_result:
                        ws_score.cell(row=i, column=col_idx).value = score_result[criteria_key]["score"]
                        col_idx += 1
                    ws_score.cell(row=i, column=col_idx).value = notes
                    ws_score.cell(row=i, column=col_idx).alignment = Alignment(wrap_text=True)

                    driver.get(self.url_lms) # quay lại trang tìm kiếm để tiếp tục duyệt qua môn học tiếp theo

                except Exception as e:
                    # nếu không tìm thấy môn học trên LMS thì sẽ ghi "Chưa soạn LMS" vào cột điểm số và ghi chú vào file excel, sau đó tiếp tục chạy vòng lặp để duyệt qua môn học tiếp theo
                    ws_score.cell(row=i, column=column_current + len(self.SCORE_CRITERIA)).value = "Chưa soạn LMS"
                    print(f"[chưa soạn] kết quả xử lý môn {subject_name} - {teacher_name} - {group_id}: Chưa soạn LMS")
                    # print(f"no: {subject_name} - {teacher_name} - {group_id}")
                ws_score.cell(row=i, column=column_current + len(self.SCORE_CRITERIA) + 2).value = "x"
                wb_score.save(excel_file)
        driver.quit()
        time.sleep(2)  # Thêm delay để dừng lại 2 giây trước khi nhập thông tin môn
        
test = ScoreLMS()
test.score_lms()
# print(test.test_lsa())