from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from config.settings import user_name, password, captcha
import time
import re

class InitSelenium:
      def init_selenium(self):
            chrome_options = Options()
            chrome_options.add_argument("--headless=new")  # Chế độ headless mới
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Tắt automation
            chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36")
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option("useAutomationExtension", False)

            #driver = webdriver.Chrome(options=chrome_options)
            driver = webdriver.Chrome()
            return driver
      

      def login_selenium(self, url):
            driver = self.init_selenium()
            driver.get(url)
            
            try:
                  button_sso_text = "//button[contains(text(), 'HCMCOU-SSO')]"
                  button_sso = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, button_sso_text))
                  )
                  button_sso.click()
            except Exception as e:
                  print(f"Không thấy nút HCMCOU-SSO, lỗi: {e}")

            try:
                  dropdown = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "form-usertype"))
                  )
                  select_type_user = Select(dropdown)
                  select_type_user.select_by_visible_text("Cán bộ-Nhân viên / Giảng viên")
            except Exception as e:
                  print(f"Không thấy lựa chọn Cán bộ nhân viên, lỗi: {e}")    
   
            try:
                  username_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "form-username"))
                  )
                  username_input.send_keys(user_name)
            except Exception as e:
                  print(f"Không thấy ô nhập tài khoản, lỗi: {e}")    

            try:
                  password_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "form-password"))
                  )
                  password_input.send_keys(password)
            except Exception as e:
                  print(f"Không thấy ô nhập mật khẩu, lỗi: {e}")    

            try:
                  captcha_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "form-captcha"))
                  )
                  captcha_input.send_keys(captcha)
            except Exception as e:
                  print(f"Không tìm thấy ô để nhập Capcha, lỗi {e}")

            
            try:
                  button_login_text = "//*[self::button or self::a or self::input][contains(., 'Đăng nhập') or contains(@value, 'Đăng nhập')]"
                  button_login = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, button_login_text))
                  )
                  button_login.click()            
            except Exception as e:
                  print(f"Không tìm thấy nút đăng nhập, lỗi {e}")

            try:
                  button_allow = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn.btn-success.btn-approve"))
                  )
                  button_allow.click()
            except Exception as e:
                  print(f"Không tìm thấy nút cho phép nhấn đồng ý, lỗi {e}")

            return driver
      

      def process_detail_subject(self, semester, url_lsa):
            driver =self. login_selenium(url_lsa)

            try:
                  dropdown_semester = WebDriverWait(driver, 15).until(
                  EC.presence_of_element_located((By.ID, "moodlesiteid"))
                  )
                  select_type_semester = Select(dropdown_semester)
                  select_type_semester.select_by_visible_text(" ".join(["[LIVE] LMS ĐTTX", semester]))
                  
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
            
            get_subject = []
            try:
                  for row in rows:
                        cells = row.find_elements(By.XPATH, ".//td")
                        temp = []
                        for i, cell in enumerate(cells):
                              get_value = cell.splitlines()
                              for i, vals in enumerate(get_value):
                                    if vals.startswith("[" + semester + "]"):
                                          temp.append(get_value[i - 1]) # gắn giá trị id môn học trên lsa
                                          pattern = r"\[{0}\]\s+(\w+)\s+-.*\((\w+)-(\w+)\)".format(semester)
                                          match = re.search(pattern, vals)
                                          if match:
                                                # group1: FINA2343, group2: KT196, group3: TN303
                                                temp.extend([match.group(1), match.group(2), match.group(3)])
                                          elif vals.startswith("Giảng viên"):
                                                temp.append(vals.split(". ", 1)[-1].strip()) # tên giảng viên

                              print(F"{i}: {cell.text}")
            except Exception as e:
                  print(f"Không xuất được nội dung từng ô trong bảng, lỗi {e}")

      