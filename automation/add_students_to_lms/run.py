# =============================================================================
# MODULE: automation/add_students_to_lms/run.py
# MỤC ĐÍCH: Tự động cập nhật danh sách sinh viên (DSSV) từ LSA vào LMS.
#
# LUỒNG CHẠY CHÍNH (update_student_lms):
#   1. Scraping LSA → lấy danh sách môn học và lms_id của từng môn
#   2. Đăng nhập LMS qua SSO
#   3. Với mỗi môn học, truy cập URL ghi danh và đọc số SV đã được cập nhật
#   4. Nếu số SV cập nhật > 0 → lưu vào MySQL (bảng automatic_work.students_lms)
#   5. Ghi log toàn bộ quá trình ra file add_students_to_lms.log
#
# URL GHI DANH LMS:
#   https://lms.oude.edu.vn/<semester>/local/enrolfromextdb/enrollnet.php?id=<lms_id>
#   Trang này tự động cập nhật DSSV và hiển thị "Đã ghi danh X/Y sinh viên"
#
# CHẠY TRỰC TIẾP:
#   python -m automation.add_students_to_lms.run
# =============================================================================

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
from db.connect_mysql import get_connection
from config import paths, runtime_settings
import os
import time
import re


class UpdateStudents:

      def __init__(self):
            """Khởi tạo URL trang tìm kiếm môn học trên LMS theo học kỳ hiện tại.

            URL này dùng để đăng nhập vào LMS (cần mở 1 trang LMS bất kỳ trước),
            sau đó từng môn sẽ được truy cập bằng URL ghi danh riêng.
            """
            # học kỳ hiện tại — lấy qua runtime_settings để sửa được trên GUI (tab "Cài đặt"),
            # không cần build lại app mỗi học kỳ
            self.semester_general = runtime_settings.load()["semester_general"]
            # URL trang search môn học trên LMS — dùng làm trang đích khi đăng nhập
            # Sau khi đăng nhập xong, driver sẽ dùng URL ghi danh riêng cho từng môn
            self.url_lms = f"https://lms.oude.edu.vn/{self.semester_general}/course/search.php"


      def update_student_lms(self):
            """Pipeline chính: cập nhật DSSV từ LSA vào LMS và lưu kết quả vào MySQL.

            Các bước xử lý:
              1. Khởi tạo logger — ghi log ra file add_students_to_lms.log
                 (file bị GHI ĐÈ mỗi lần chạy, không tích lũy)
              2. Scraping LSA → lấy danh sách môn học và lms_id
              3. Đăng nhập LMS qua SSO
              4. Với mỗi môn trong danh sách:
                 a. Truy cập URL ghi danh: enrollnet.php?id=<lms_id>
                    → Trang này tự động kích hoạt cập nhật DSSV từ hệ thống QLĐT
                 b. Đọc text trong div "Đã ghi danh X/Y sinh viên"
                    → Dùng JavaScript để lấy text của node cha (tránh lấy text của node con)
                 c. Regex tách ra [X, Y] → X là số SV vừa được ghi danh
                 d. Nếu X != '0' → INSERT vào MySQL (bảng automatic_work.students_lms)
                    Nếu X == '0' → bỏ qua (không có gì thay đổi, không cần lưu)
              5. Mỗi môn có transaction MySQL riêng (try/except/finally con.close())
                 để lỗi 1 môn không ảnh hưởng các môn còn lại

            Cấu trúc bảng MySQL (automatic_work.students_lms):
                lms_id, group_id, subject_id, subject_name,
                teacher_id, teacher_name, date_update, number_student_update, semester
            """
            start_selenium = InitSelenium()

            # Đường dẫn file log — đặt cùng thư mục với file run.py này
            file_dir = paths.ADD_STUDENTS_LOG_FILE
            # setup_logger ghi đè file log cũ mỗi lần chạy — xem utils/logger.py
            logger = setup_logger(file_dir)
            logger.info("Bắt đầu chạy chương trình cập nhật sinh viên vào LMS")

            # Scraping LSA để lấy danh sách môn học cần xử lý
            print("Bắt đầu lấy dữ liệu LSA...")
            list_subject_lsa = start_selenium.process_get_detail_lsa()
            logger.info(f"Đã lấy được {len(list_subject_lsa)} môn học từ LSA")

            # Đăng nhập LMS, driver này sẽ dùng xuyên suốt vòng lặp bên dưới
            print("Đăng nhập LMS...")
            driver = start_selenium.login_selenium(self.url_lms)

            for one_subject in list_subject_lsa:
                  # Truy cập URL ghi danh của từng môn học
                  # enrollnet.php tự động kích hoạt đồng bộ DSSV từ hệ thống QLĐT vào LMS
                  driver.get(f"https://lms.oude.edu.vn/{self.semester_general}/local/enrolfromextdb/enrollnet.php?id={one_subject["lms_id"]}")

                  try:
                        # Tìm div thông báo kết quả — text dạng: "Đã ghi danh 5/25 sinh viên vào khóa học."
                        div_annoucement = WebDriverWait(driver, 10).until(
                              EC.visibility_of_element_located((By.XPATH, "//div[contains(text(), 'Đã ghi danh')]"))
                        )

                        # Dùng JavaScript để lấy text của node cha trực tiếp (childNodes[0])
                        # Vì div này có thể chứa các node con (link, span, ...) nên không dùng .text
                        text = driver.execute_script("""
                        return arguments[0].childNodes[0].nodeValue;
                        """, div_annoucement)

                        # Làm sạch text: bỏ xuống dòng, khoảng trắng thừa
                        # Kết quả mong đợi: "Đã ghi danh 5/25 sinh viên vào khóa học."
                        text = text.strip().replace('\n', ' ')

                        # Dùng regex tách ra các số trong chuỗi → [số_SV_ghi_danh, tổng_SV]
                        # Ví dụ: "Đã ghi danh 5/25 sinh viên" → ['5', '25']
                        count_students_update = re.findall(r'\d+', text)

                        # Chỉ lưu vào DB nếu có SV được cập nhật (count_students_update[0] != '0')
                        # Nếu = '0' nghĩa là không có thay đổi → không cần ghi DB
                        if count_students_update[0] != '0':
                            con = get_connection()
                            try:
                                    with con.cursor() as cursor:
                                          # INSERT thông tin môn học vừa cập nhật vào bảng students_lms
                                          sql_insert = "INSERT INTO automatic_work.students_lms (" \
                                                      "lms_id, " \
                                                      "group_id, " \
                                                      "subject_id, " \
                                                      "subject_name, " \
                                                      "teacher_id, " \
                                                      "teacher_name, " \
                                                      "date_update, " \
                                                      "number_student_update, " \
                                                      "semester) " \
                                          "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"
                                          cursor.execute(sql_insert, (
                                                one_subject["lms_id"],
                                                one_subject["group_id"],
                                                one_subject["subject_id"],
                                                one_subject["subject_name"],
                                                one_subject["teacher_id"],
                                                one_subject["teacher_name"],
                                                time.strftime('%Y-%m-%d %H:%M:%S'),  # thời điểm chạy
                                                count_students_update[0],            # số SV được ghi danh
                                                self.semester_general
                                          ))
                                    con.commit()
                                    print(f"Đã cập nhật môn {one_subject['subject_name']} - {one_subject['group_id']} - {one_subject['teacher_name']}: số lượng {count_students_update[0]}")
                            except Exception as e:
                                  # Ghi log lỗi DB nhưng KHÔNG dừng chương trình — xử lý môn tiếp theo
                                  print(f"Lỗi môn {one_subject['subject_name']}, nhóm {one_subject['group_id']} giảng viên {one_subject['teacher_name']}: cụ thể {e}")
                                  logger.error(f"Lỗi môn {one_subject['subject_name']}, nhóm {one_subject['group_id']} giảng viên {one_subject['teacher_name']}: cụ thể {e}")
                            finally:
                                  # Luôn đóng kết nối dù thành công hay lỗi
                                  con.close()

                  except Exception as e:
                        # Không tìm thấy div "Đã ghi danh" — trang có thể lỗi hoặc timeout
                        print(f"Không xuất ra số sinh viên cập nhật, lỗi {e}")
                        logger.error(f"Không xuất ra được số sinh viên cập nhật của môn {one_subject['subject_name']} - {one_subject['group_id']}")

            logger.info("Hoàn tất chạy chương trình cập nhật sinh viên vào LMS")
            driver.quit()


if __name__ == "__main__":
      ob = UpdateStudents()
      ob.update_student_lms()