"""Module cho tự động cập nhật danh sách sinh viên từ LSA sang LMS.

Các class và phương thức:
- UpdateStudents.update_student_lms: pipeline chính, đọc dữ liệu từ LSA, cập nhật qua LMS và lưu log+DB.
- UpdateStudents.find_subject_lms: trợ giúp kiểm tra chức năng tìm môn học trên LMS.

Yêu cầu config:
- automation/add_students_to_lms/config.yaml
- config/settings.py (credentials, captcha, etc)
"""

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
import os
import time 
import re

class UpdateStudents:
      def __init__(self):
            """Khởi tạo đối tượng, đọc cấu hình và gán URL cho trình tự chạy."""
            self.config = ConfigLoader(os.path.join(os.path.dirname(__file__), "config.yaml"))
            self.url_lms = f"https://lms.oude.edu.vn/{self.config.get_attr('SEMESTER')}/course/search.php"
            self.url_lsa = "http://lsa.ou.edu.vn/auth/login"


      def test(self):
          con = get_connection()
          cursor = con.cursor()
          print("kết nối thành công")


      def update_student_lms(self):
            """Chạy quy trình cập nhật DSSV trên LMS từ thông tin LSA.

            Bước:
            1. Đọc học kỳ từ config.
            2. Lấy danh sách môn học từ LSA.
            3. Đăng nhập LMS và lần lượt vào từng môn học.
            4. Đọc thông báo "Đã ghi danh" để lấy số SV cập nhật.
            5. Ghi log và lưu vào MySQL (bảng automatic_work.students_lms).

            Lưu ý:
            - Nếu không tìm được dữ liệu trên trang, ghi lỗi vào log.
            - Nếu DB bị lỗi, rollback với từng môn học.
            """
            start_selenium = InitSelenium()
            semester = str(self.config.get_attr("SEMESTER"))

            file_dir = os.path.join(os.path.dirname(__file__), "add_students_to_lms.log")
            # nếu file log đã tồn tại, xóa file cũ để tạo file log mới cho lần chạy hiện tại
            logger = setup_logger(file_dir)
            logger.info("Bắt đầu chạy chương trình cập nhật sinh viên vào LMS")
            list_subject_lsa = start_selenium.process_detail_subject(semester, self.url_lsa)
            logger.info(f"Đã lấy được {len(list_subject_lsa)} môn học từ LSA")

            # đăng nhập vào LMS và cập nhật sinh viên cho từng môn học
            driver = start_selenium.login_selenium(self.url_lms)
            for one_subject in list_subject_lsa:
                  driver.get(f"https://lms.oude.edu.vn/{self.config.get_attr("SEMESTER")}/admin/tool/enrolfromextdb/enrollnet.php?id={one_subject["lms_id"]}")
                  try:
                        div_annoucement = WebDriverWait(driver, 10).until(
                              EC.visibility_of_element_located((By.XPATH, "//div[contains(text(), 'Đã ghi danh')]"))
                        )
                        text = driver.execute_script("""
                        return arguments[0].childNodes[0].nodeValue;
                        """, div_annoucement) # vì trong div cần lây có các node con, chỉ cần lấy text của lớp cha
                        text = text.strip().replace('\n', ' ') #Đã ghi danh 0/25 sinh viên vào khóa học.
                        count_students_update = re.findall(r'\d+', text) #[0, 27]

                        # nếu số sinh viên được cập nhật khác 0, lưu thông tin vào database để tiện theo dõi, 
                        # nếu không sẽ không lưu vì không có gì thay đổi
                        # if count_students_update[0] != '0':
                        print(type(count_students_update[0]))
                        con = get_connection()
                        try:
                              with con.cursor() as cursor:
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
                                          time.strftime('%Y-%m-%d %H:%M:%S'),
                                          count_students_update[0],
                                          semester
                                          ))
                              con.commit()
                        except Exception as e:
                              logger.error(f"Lỗi môn {one_subject['subject_name']}, nhóm {one_subject['group_id']} giảng viên {one_subject['teacher_name']}: cụ thể {e}")
                        finally:
                              con.close()
                  except Exception as e:
                        print(f"Không xuất ra số sinh viên cập nhật, lỗi {e}")  
                        logger.error(f"Không xuất ra được số sinh viên cập nhật của môn {one_subject['subject_name']} - {one_subject['group_id']}")
            logger.info("Hoàn tất chạy chương trình cập nhật sinh viên vào LMS")


      def find_subject_lms(self):
            """Thử nghiệm chức năng tìm môn học trên LMS và in kết quả sơ bộ.

            Hàm này dùng để kiểm tra bước tìm môn học với từ khóa tĩnh TN120.
            Nếu muốn dùng general, hãy thường xuyên refactor để truyền tham số đầu vào.
            """
            start_selenium = InitSelenium()
            driver = start_selenium.login_selenium(self.url_lms)

            try:
                  search_input = WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.NAME, "q"))
                  )
                  search_input.send_keys("TN120")
            except Exception as e:
                  print(f"Không tìm thấy ô nhập môn học, lỗi {e}")
            
            try:
                  button_find_text = "button[type='submit']"
                  button_find = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, button_find_text))
                  )
                  button_find.click()
            except Exception as e:
                  print(f"Không tìm thấy nút tìm môn học, lỗi {e}")
      
            try:
                  list_subject = WebDriverWait(driver, 10).until(
                        EC.presence_of_all_elements_located((By.CLASS_NAME, "aalink"))
                  )
                  
                  main_tab = driver.current_window_handle
                  # đưa các link vào list, vì nếu dùng trực tiếp sb.get_attribute("href")
                  # trong for, khi mở tab mới sẽ không thấy link
                  list_link_subject = [sb.get_attribute("href") for sb in list_subject]
                  for link_subject in list_link_subject:
                        driver.switch_to.new_window('tab')
                        driver.get(link_subject)

                        try:
                              btn_read_more = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.LINK_TEXT, "Xem thêm"))
                              )
                              btn_read_more.click()
                        except Exception as e:
                              print(f"Không mở được nút xem thêm, lỗi {e}")  
                        
                        # xem lại chỗ này, chạy lần dầu tiên không cập nhật được môn đầu tiên
                        # xem có thể sử dụng cái này được không select_by_visible_text
                        try:
                              btn_update_student = WebDriverWait(driver, 15).until(
                                    EC.element_to_be_clickable((By.LINK_TEXT, "Cập nhật DSSV từ hệ thống QLĐT")) 
                              )
                              btn_update_student.click()
                        except Exception as e:
                              print(f"Không mở được nút cập nhật sinh viên, lỗi {e}")  

                        try:
                              div_annoucement = WebDriverWait(driver, 10).until(
                                    EC.visibility_of_element_located((By.XPATH, "//div[contains(text(), 'Đã ghi danh')]"))
                              )
                              text = driver.execute_script("""
                              return arguments[0].childNodes[0].nodeValue;
                              """, div_annoucement) # vì trong div cần lây có các node con, chỉ cần lấy text của lớp cha
                              text = text.strip().replace('\n', ' ') #Đã ghi danh 0/25 sinh viên vào khóa học.
                              numbers = re.findall(r'\d+', text) #[0, 27]
                              print(numbers)
                        except Exception as e:
                              print(f"Không xuất ra số sinh viên cập nhật, lỗi {e}")  

                        driver.close() 
                        driver.switch_to.window(main_tab) 
                  
            except Exception as e:
                  print(f"Không mở được tab mới, lỗi {e}")  

ob = UpdateStudents()
# print(ob.test())
ob.update_student_lms()