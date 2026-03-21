from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from config.settings import user_name, password, captcha
import time

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
            driver.switch_to.new_window('tab')
            driver.get("https://lms.oude.edu.vn/251/course/view.php?id=608")
            # try:
            #       list_subject = WebDriverWait(driver, 10).until(
            #             EC.presence_of_all_elements_located((By.CLASS_NAME, "aalink"))
            #       )
                  
                  # main_tab = driver.current_window_handle
                  # for sb in list_subject:
                  #       print(sb.get_attribute("href"))
                  #       driver.switch_to.new_window('tab')
                  #       driver.get(sb.get_attribute("href"))

                        # btn_read_more = WebDriverWait(driver, 10).until(
                        #       EC.element_to_be_clickable((By.LINK_TEXT, "Xem thêm"))
                        # )
                        # btn_read_more.click()

                        # btn_update_student = WebDriverWait(driver, 10).until(
                        #       EC.element_to_be_clickable((By.LINK_TEXT, "Xem thêm"))
                        # )
                        # btn_update_student.click()

                        # div_annoucement = WebDriverWait(driver, 10).until(
                        #       EC.visibility_of_element_located((By.XPATH, "//div[contains(., 'Đã ghi danh')]"))
                        # )
                        # print(div_annoucement.text)
                        # driver.close() 

                  #driver.switch_to.window(main_tab) 
            # except Exception as e:
            #       print(f"Không tìm thấy nút tìm môn học, lỗi {e}")  

            
            