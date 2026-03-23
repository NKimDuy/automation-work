from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from modules.init_selenium import InitSelenium
from utils.config_loader import ConfigLoader
import os
import time 
import re

class UpdateStudents:
      def __init__(self):
            self.config = ConfigLoader(os.path.join(os.path.dirname(__file__), "config.yaml"))
            self.url_lms = f"https://lms.oude.edu.vn/{self.config.get_attr("SEMESTER")}/course/search.php"
            self.url_lsa = f"http://lsa.ou.edu.vn/auth/login"

      def process_detail_subject(self):
            start_selenium = InitSelenium()
            driver = start_selenium.login_selenium(self.url_lsa)

            try:
                  dropdown_semester = WebDriverWait(driver, 15).until(
                  EC.presence_of_element_located((By.ID, "moodlesiteid"))
                  )
                  select_type_semester = Select(dropdown_semester)
                  select_type_semester.select_by_visible_text(" ".join(["[LIVE] LMS ĐTTX", str(self.config.get_attr("SEMESTER"))]))
                  
            except Exception as e:
                  print(f"Không tìm thấy dropdownlist thể hiện học kỳ, lỗi {e}")

            try:
                  driver.execute_script("arguments[0].style.display='block';", driver.find_element(By.ID, "menu_1_sub"))
                  overview_link = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='usersiteoverviews']"))
                  )
                  overview_link.click()
            except Exception as e:
                  print(f"Không tìm thấy nút report, lỗi {e}")

            try:
                  table = WebDriverWait(driver, 40).until(
                        EC.presence_of_element_located((By.ID, "ourptlistcourse"))  # Thay "myTable" bằng ID thực tế
                  )
            except Exception as e:
                  print(f"Không tìm thấy bảng, lỗi {e}")

            try:
                  rows = WebDriverWait(driver, 20).until(
                        EC.presence_of_all_elements_located((By.XPATH, ".//tr"))
                  )
            except Exception as e:
                  print(f"Không tìm thấy các dòng, lỗi {e}")

            try:
                  get_subject = []
                  for row in rows:
                        cells = row.find_elements(By.XPATH, ".//td")
                        print('--------------------------')
                        for i, cell in enumerate(cells):
                              print(F"{i}: {cell.text}")
            except Exception as e:
                  print(f"Không xuất được nội dung từng ô trong bảng, lỗi {e}")


      def process_update_student(self):
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

test = UpdateStudents()
test.process_detail_subject()