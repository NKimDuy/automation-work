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
from config.settings import semester
import os
import time
import re


class UpdateStudents:

      def __init__(self):
            """Khởi tạo URL trang tìm kiếm môn học trên LMS theo học kỳ hiện tại.

            URL này dùng để đăng nhập vào LMS (cần mở 1 trang LMS bất kỳ trước),
            sau đó từng môn sẽ được truy cập bằng URL ghi danh riêng.
            """
            # URL trang search môn học trên LMS — dùng làm trang đích khi đăng nhập
            # Sau khi đăng nhập xong, driver sẽ dùng URL ghi danh riêng cho từng môn
            self.url_lms = f"https://lms.oude.edu.vn/{semester}/course/search.php"


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
            file_dir = os.path.join(os.path.dirname(__file__), "add_students_to_lms.log")
            # setup_logger ghi đè file log cũ mỗi lần chạy — xem utils/logger.py
            logger = setup_logger(file_dir)
            logger.info("Bắt đầu chạy chương trình cập nhật sinh viên vào LMS")

            # Scraping LSA để lấy danh sách môn học cần xử lý
            list_subject_lsa = start_selenium.process_get_detail_lsa()
            logger.info(f"Đã lấy được {len(list_subject_lsa)} môn học từ LSA")

            # Đăng nhập LMS, driver này sẽ dùng xuyên suốt vòng lặp bên dưới
            driver = start_selenium.login_selenium(self.url_lms)

            for one_subject in list_subject_lsa:
                  # Truy cập URL ghi danh của từng môn học
                  # enrollnet.php tự động kích hoạt đồng bộ DSSV từ hệ thống QLĐT vào LMS
                  driver.get(f"https://lms.oude.edu.vn/{semester}/local/enrolfromextdb/enrollnet.php?id={one_subject["lms_id"]}")

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
                                              semester
                                        ))
                                  con.commit()
                            except Exception as e:
                                  # Ghi log lỗi DB nhưng KHÔNG dừng chương trình — xử lý môn tiếp theo
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


      def find_subject_lms(self):
            """Hàm thử nghiệm: tìm môn học trên LMS và thử cập nhật DSSV thủ công.

            Hàm này KHÔNG dùng trong luồng tự động — chỉ dùng để kiểm tra
            từng bước khi debug hoặc phát triển tính năng mới.

            Luồng thử nghiệm:
              1. Đăng nhập LMS
              2. Tìm kiếm môn với từ khóa cứng "TN120"
              3. Mở từng kết quả trong tab mới
              4. Click "Xem thêm" → click "Cập nhật DSSV từ hệ thống QLĐT"
              5. Đọc kết quả "Đã ghi danh" và in ra màn hình

            Lưu ý:
              - Từ khóa "TN120" được hardcode — nếu muốn dùng linh hoạt cần
                refactor để nhận tham số đầu vào.
              - Hàm mở nhiều tab → cần switch_to.window() để điều hướng đúng.
              - Link href phải được lưu vào list trước khi mở tab mới,
                vì sau khi switch tab thì driver không còn thấy element cũ.
            """
            start_selenium = InitSelenium()
            driver = start_selenium.login_selenium(self.url_lms)

            # Tìm kiếm môn học với từ khóa "TN120"
            try:
                  search_input = WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.NAME, "q"))
                  )
                  search_input.send_keys("TN120")
            except Exception as e:
                  print(f"Không tìm thấy ô nhập môn học, lỗi {e}")

            # Nhấn nút tìm kiếm
            try:
                  button_find_text = "button[type='submit']"
                  button_find = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, button_find_text))
                  )
                  button_find.click()
            except Exception as e:
                  print(f"Không tìm thấy nút tìm môn học, lỗi {e}")

            try:
                  # Lấy tất cả link môn học trong kết quả tìm kiếm (class="aalink")
                  list_subject = WebDriverWait(driver, 10).until(
                        EC.presence_of_all_elements_located((By.CLASS_NAME, "aalink"))
                  )

                  main_tab = driver.current_window_handle

                  # QUAN TRỌNG: lưu href vào list Python trước khi mở tab mới
                  # Nếu dùng sb.get_attribute("href") trực tiếp trong vòng for sau khi đã switch tab,
                  # driver sẽ không tìm thấy element vì context đã thay đổi
                  list_link_subject = [sb.get_attribute("href") for sb in list_subject]

                  for link_subject in list_link_subject:
                        # Mở từng môn trong tab mới để không mất trang danh sách
                        driver.switch_to.new_window('tab')
                        driver.get(link_subject)

                        # Click "Xem thêm" để hiển thị toàn bộ nội dung trang môn học
                        try:
                              btn_read_more = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.LINK_TEXT, "Xem thêm"))
                              )
                              btn_read_more.click()
                        except Exception as e:
                              print(f"Không mở được nút xem thêm, lỗi {e}")

                        # Click nút "Cập nhật DSSV từ hệ thống QLĐT" để kích hoạt đồng bộ
                        # Lưu ý: lần đầu tiên chạy có thể không cập nhật được môn đầu tiên
                        # (đang điều tra nguyên nhân — có thể dùng select_by_visible_text thay thế)
                        try:
                              btn_update_student = WebDriverWait(driver, 15).until(
                                    EC.element_to_be_clickable((By.LINK_TEXT, "Cập nhật DSSV từ hệ thống QLĐT"))
                              )
                              btn_update_student.click()
                        except Exception as e:
                              print(f"Không mở được nút cập nhật sinh viên, lỗi {e}")

                        # Đọc kết quả sau khi cập nhật
                        try:
                              div_annoucement = WebDriverWait(driver, 10).until(
                                    EC.visibility_of_element_located((By.XPATH, "//div[contains(text(), 'Đã ghi danh')]"))
                              )
                              # Dùng JavaScript lấy text node cha (tránh lấy text node con)
                              text = driver.execute_script("""
                              return arguments[0].childNodes[0].nodeValue;
                              """, div_annoucement)
                              text = text.strip().replace('\n', ' ')
                              numbers = re.findall(r'\d+', text)  # ['số_ghi_danh', 'tổng_SV']
                              print(numbers)
                        except Exception as e:
                              print(f"Không xuất ra số sinh viên cập nhật, lỗi {e}")

                        # Đóng tab hiện tại và quay về tab danh sách môn học
                        driver.close()
                        driver.switch_to.window(main_tab)

            except Exception as e:
                  print(f"Không mở được tab mới, lỗi {e}")


