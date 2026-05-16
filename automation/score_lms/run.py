from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from modules.init_selenium import InitSelenium
from utils.config_loader import ConfigLoader
from utils.logger import setup_logger
# from db.connect_mysql import get_connection
from config.settings import semester
import os
import time
import re
import copy
import joblib
import re

class ScoreLMS:
    def __init__(self):
        # self.url_lms = f"https://lms.oude.edu.vn/{semester}/course/search.php"
        self.url_lms = f"https://lms.oude.edu.vn/251/course/search.php"
    
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
    

    # def predict(self, model, texts: list) -> list:
    #     cleaned = [self.preprocess(t) for t in texts]
    #     preds = model.predict(cleaned)
    #     for text, label in zip(texts, preds):
    #         print(f"  [{label}] {text}")
    #     return preds.tolist()


    def check_forum(self, driver, forum_link, ten_giang_vien):
        second_tab = driver.current_window_handle # lưu tab môn học, vì khi mở diễn đàn, sẽ mở sang tab khác
        driver.execute_script("window.open(arguments[0]);", forum_link)
        driver.switch_to.window(driver.window_handles[-1])

        rows = driver.find_elements(By.CSS_SELECTOR, "tr.discussion")

        chua_reply = []
        if rows:
            for row in rows:
                try:
                    ten_sv = row.find_element(By.CSS_SELECTOR, "td.author .mb-1").text.strip()
                except:
                    ten_sv = "Không rõ"

                try:
                    ten_reply = row.find_element(By.CSS_SELECTOR, "td.text-start .mb-1").text.strip()
                except:
                    ten_reply = "Không có"

                # print(f"Tên SV: {ten_sv}, Tên GV: {ten_giang_vien}, Tên reply: {ten_reply}")
                if ten_giang_vien.lower() not in ten_reply.lower():
                    chua_reply.append(ten_sv)

        driver.close()
        driver.switch_to.window(second_tab)
        return chua_reply


    def score_lms(self):
        start_selenium = InitSelenium()
        driver = start_selenium.login_selenium(self.url_lms)

        dic_score_base = {
            "thlms": {"score": "", "name": "Thực hiện LMS"},
            "ctdd": {"score": 0, "name": "Cấu trúc đầy đủ"},
            "ttkh": {"score": 0, "name": "Thông tin về khóa học"},
            "mtht": {"score": 0, "name": "Mục tiêu, hoạt động học tập"},
            "ttbg": {"score": 0, "name": "Thông tin bài giảng"},
            "bdgk": {"score": 0, "name": "Bảng điểm giữa kỳ"},
            "cbg": {"score": "", "name": "Video ghi hình buổi giảng"},
            "tltk": {"score": "", "name": "Tài liệu tham khảo/bài đọc thêm"},
            "bt": {"score": 0, "name": "Bài tập tự luận/trắc nghiệm"},
            "dd": {"score": "", "name": "Diễn đàn"},
            "vgdtm": {"score": "", "name": "video giải đáp thắc mắc cho sinh viên"},
            "svdk": {"score": 0, "name": "Sinh viên đăng ký"},
            "svtc": {"score": 0, "name": "Sinh viên truy cập"},
            "tlsvtc": {"score": 0, "name": "Tỷ lệ sinh viên truy cập"},
            "tlgvtc": {"score": 0, "name": "Tỷ lệ giảng viên truy cập"},
            "td": {"score": 0, "name": "Tổng điểm"},
            "kqdg": {"score": "x", "name": "kết quả đánh giá"},
        }
        
        dic_criteria = {
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

        not_rated = [
            "Thí nghiệm vật liệu xây dựng",
            "Thí nghiệm cơ chất lỏng",
            "Thí nghiệm sức bền vật liệu",
            "Thực tập trắc địa",
            "Địa chất công trình + thực tập",
            "Thí nghiệm cơ học đất",
            "Nhận thức ngành",
            'Thực tập địa chất công trình',
            "Đồ án nền móng",
            "Đồ án kết cấu bê tông cốt thép 1, 2",
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

        noise_remove = [
            'rút gọn',
            "các thông báo",
            "bài tập",
            "câu hỏi thảo luận",
            "diễn đàn",
            'Chuyển tới nội dung chính',
            'LMS ĐTTX',
            'Topic outline',
            'Thu gọn toàn bộ',
            'Copyright © 2024. Powered by Moodle'
        ]

        notes = ''
        sum_score = 0

        # copy từ điển điểm để áp dụng cho từng môn học, tránh bị ghi đè khi duyệt qua nhiều môn
        dic_score_apply = copy.deepcopy(dic_score_base)

        base_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        model_file = os.path.join(base_dir, "AI", "model", "model_vi_classification.pkl")
        model = joblib.load(model_file)

        #chỗ này sễ chạy vòng lặp for để duyệt qua các môn học trong file excel
        try:
            search_input = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.NAME, "q"))
            )
            search_input.send_keys("BADM1391 - (QT653-TN120)")
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
            # Lấy tất cả link môn học trong kết quả tìm kiếm (class="aalink")
            get_list_subject = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "aalink"))
            )

            # lưu tab tìm thông tin môn học
            main_tab = driver.current_window_handle

            # QUAN TRỌNG: lưu href vào list Python trước khi mở tab mới
            # Nếu dùng sb.get_attribute("href") trực tiếp trong vòng for sau khi đã switch tab,
            # driver sẽ không tìm thấy element vì context đã thay đổi
            href_link_subject = get_list_subject.get_attribute("href")

            forums = {} # chứa thông tin diễn đàn: {ten_forum: link_forum}
            quiz = {}  # chứa thông tin bài tập quiz: {ten_quiz: link_quiz}
            assign = {} # chứa thông tin bài tập tự luận: {ten_assign: link_assign}
            
            # duyệt qua từng link môn học tìm được và mở trong tab mới
            
            driver.execute_script("window.open(arguments[0]);", href_link_subject)
            driver.switch_to.window(driver.window_handles[-1])

            dic_score_apply["thlms"]["score"] = 15
            sum_score += 15
            dic_score_apply["ctdd"]["score"] = 15
            sum_score += 15

            # lấy tát cả text có trong môn học 
            full_text = driver.find_element(By.TAG_NAME, "body").text
            lines = full_text.splitlines()
            lines = [item for item in lines if len(item.split()) > 1]
            lines = [line for line in lines if not any(noise in line.lower() for noise in noise_remove)]
            results = self.predict(model, lines)
            results = list(set(results))  # lấy kết quả duy nhất để đánh giá

            get_ttkh = [1, 2, 3, 4, 5] 
            check_ttkn = [x for x in get_ttkh if x not in set(results)]
            if not check_ttkn:
                dic_score_apply["ttkh"]["score"] = 15
                sum_score += 15
            else:
                name_not_check_ttkn = ["Thiếu " + dic_criteria[i] for i in check_ttkn]
                notes += "\n".join(name_not_check_ttkn)
            
            if 6 in results:
                dic_score_apply["mtht"]["score"] = 20
                sum_score += 20
            else:
                notes += "\nThiếu giới thiệu về mục tiêu học tập"
            if 7 in results:
                dic_score_apply["ttbg"]["score"] = 20
                sum_score += 20
            else:
                notes += "\nThiếu nội dung bài giảng"
            if 8 in results:
                dic_score_apply["tltk"]["score"] = "x"
            else:
                notes += "\nThiếu tài liệu tham khảo/bài đọc thêm"
            if 9 in results:
                dic_score_apply["cbg"]["score"] = "x"
            if 10 in results:
                dic_score_apply["bdgk"]["score"] = 15
                sum_score += 15
            else:
                notes += "\nThiếu bảng điểm giữa kỳ"

            dic_score_apply["svdk"]["score"] = 15
            dic_score_apply["svtc"]["score"] = 15
            dic_score_apply["tlsvtc"]["score"] = 0.7
            dic_score_apply["tlgvtc"]["score"] = 1

            
            dic_score_apply["td"]["score"] = sum_score






            all_links = driver.find_elements(By.TAG_NAME, "a")
            for a in all_links:
                a_href = a.get_attribute("href")
                a_text = a.text.strip()
                
                if "forum" in a_href:
                    forums[a_text] = a_href
                if "quiz" in a_href:
                    quiz[a_text] = a_href
                if "assign" in a_href:
                    assign[a_text] = a_href
            


            check_bt = [name for name, d in [("quiz", quiz), ("assign", assign)] if not d]
            if not check_bt:
                dic_score_apply["bt"]["score"] = "x"
            else:
                if check_bt == ["quiz"]:
                    notes += "\nThiếu bài tập trắc nghiệm"
                elif check_bt == ["assign"]:
                    notes += "\nThiếu bài tập tự luận"
                else:
                    notes += "\nThiếu bài tập tự luận/trắc nghiệm"
            print(check_bt)

            
            print(f"quiz links  : {quiz}")
            print(f"assign links: {assign}")

            print(f"forum links : {forums}")
            if forums:
                forum_chua_day_du = []  # lưu các diễn đàn chưa tương tác đầy đủ [(ten_forum, [sv_chua_reply])]
                for name, forum_link in forums.items():
                    chua_reply = self.check_forum(driver, forum_link, "Võ Thị Hồng Nhung")
                    if chua_reply:
                        forum_chua_day_du.append({
                            "ten_forum": name
                        })
                if not forum_chua_day_du:
                    print(f" ✅ Có diễn đàn và tương tác với tất cả sinh viên")
                    dic_score_apply["dd"]["score"] = "x"
                else:
                    print(f" ⚠️ Có diễn đàn nhưng không tương tác đầy đủ với sinh viên")
                    for f in forum_chua_day_du:
                        print(f"  📌 Forum: {f['ten_forum']}")
                        notes += f"\nDiễn đàn '{f['ten_forum']}' chưa tương tác đầy đủ với sinh viên"
            else:
                print(f" ❌ Không có diễn đàn nào")
                notes += "\nKhông có diễn đàn"    

            time.sleep(2)  # Thêm delay nếu cần thiết để tránh bị khóa tài khoản do hoạt động quá nhanh
            
            driver.close()
            driver.switch_to.window(main_tab)   
        except Exception as e:
                  print(f"Không mở được tab mới, lỗi {e}")

        print(f"\nTổng điểm: {sum_score}/100")
        print(f"\nGhi chú: {notes}")
        driver.quit()
test = ScoreLMS()
test.score_lms()