from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from config.settings import user_lms, password_lms, captcha
from config.settings import semester
from utils.api import APIHandler
import time
import re

class InitSelenium:
      """
      Lớp quản lý tương tác với trình duyệt Chrome thông qua Selenium
      Chức năng chính: Đăng nhập, lấy thông tin môn học từ LMS và PHDT
      """
      
      def __init__(self):
            self.url_lsa = "http://lsa.ou.edu.vn/auth/login"

      def init_selenium(self):
            """
            Khởi tạo trình duyệt Chrome với các cấu hình tối ưu
            - Chế độ headless: chạy ẩn không hiển thị giao diện
            - Vô hiệu hóa công nghệ phát hiện automation
            - Giả lập user-agent để tránh bị chặn
            Return: webdriver.Chrome instance
            """
            chrome_options = Options()
            chrome_options.add_argument("--headless=new")  # Chế độ headless mới
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Tắt công nghệ phát hiện automation
            chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36")
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option("useAutomationExtension", False)

            # Sử dụng cấu hình mặc định nếu không cần service
            driver = webdriver.Chrome()
            return driver
      

      def login_selenium(self, url):
            """
            Đăng nhập vào hệ thống LMS qua SSO
            Quy trình:
            1. Mở URL
            2. Nhấn nút SSO
            3. Chọn loại người dùng (Cán bộ-Nhân viên / Giảng viên)
            4. Nhập tài khoản, mật khẩu, captcha
            5. Nhấn đăng nhập
            6. Nhấn nút cho phép/đồng ý
            
            Args: url (str) - Đường dẫn website cần truy cập
            Return: webdriver.Chrome - driver đã đăng nhập thành công
            """
            driver = self.init_selenium()
            driver.get(url)
            
            # Bước 1: Nhấn nút HCMCOU-SSO
            try:
                  button_sso_text = "//button[contains(text(), 'HCMCOU-SSO')]"
                  button_sso = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, button_sso_text))
                  )
                  button_sso.click()
            except Exception as e:
                  print(f"Không thấy nút HCMCOU-SSO, lỗi: {e}")

            # Bước 2: Chọn loại người dùng
            try:
                  dropdown = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "form-usertype"))
                  )
                  select_type_user = Select(dropdown)
                  select_type_user.select_by_visible_text("Cán bộ-Nhân viên / Giảng viên")
            except Exception as e:
                  print(f"Không thấy lựa chọn Cán bộ nhân viên, lỗi: {e}")    
   
            # Bước 3: Nhập tài khoản
            try:
                  username_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "form-username"))
                  )
                  username_input.send_keys(user_lms)
            except Exception as e:
                  print(f"Không thấy ô nhập tài khoản, lỗi: {e}")    

            # Bước 4: Nhập mật khẩu
            try:
                  password_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "form-password"))
                  )
                  password_input.send_keys(password_lms)
            except Exception as e:
                  print(f"Không thấy ô nhập mật khẩu, lỗi: {e}")    

            # Bước 5: Nhập mã captcha
            try:
                  captcha_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "form-captcha"))
                  )
                  captcha_input.send_keys(captcha)
            except Exception as e:
                  print(f"Không tìm thấy ô để nhập Capcha, lỗi {e}")

            # Bước 6: Nhấn nút đăng nhập
            try:
                  button_login_text = "//*[self::button or self::a or self::input][contains(., 'Đăng nhập') or contains(@value, 'Đăng nhập')]"
                  button_login = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, button_login_text))
                  )
                  button_login.click()            
            except Exception as e:
                  print(f"Không tìm thấy nút đăng nhập, lỗi {e}")

            # Bước 7: Nhấn nút cho phép/đồng ý
            try:
                  button_allow = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn.btn-success.btn-approve"))
                  )
                  button_allow.click()
            except Exception as e:
                  print(f"Không tìm thấy nút cho phép nhấn đồng ý, lỗi {e}")

            return driver
      

      def process_get_detail_lsa(self):
            """
            Truy cập LMS và lấy thông tin chi tiết của từng môn học
            Thông tin lấy được:
            - lms_id: ID môn học trên LMS
            - subject_id: Mã môn học (vd: FINA2343)
            - subject_name: Tên môn học
            - teacher_id: Mã giảng viên
            - teacher_name: Tên giảng viên
            - group_id: Mã nhóm lớp
            - count_total_student: Tổng số sinh viên
            - count_student_access: Số sinh viên đã truy cập
            - percent_student_access: % sinh viên truy cập
            - percent_teacher_access: % giảng viên truy cập
            
            Args: 
              semester (str) - Ký hiệu học kỳ (vd: "202403")
              url_lsa (str) - Đường dẫn tới LMS
            Return: list - Danh sách dict chứa thông tin các môn học
            """
            driver = self.login_selenium(self.url_lsa)

            # Chọn học kỳ từ dropdown
            try:
                  dropdown_semester = WebDriverWait(driver, 15).until(
                  EC.presence_of_element_located((By.ID, "moodlesiteid"))
                  )
                  select_type_semester = Select(dropdown_semester)
                  select_type_semester.select_by_visible_text(" ".join(["[LIVE] LMS ĐTTX", semester]))
                  
            except Exception as e:
                  print(f"Không tìm thấy dropdownlist thể hiện học kỳ, lỗi {e}")

            # Mở menu và click vào report/overview
            try:
                  # Hiển thị menu ẩn
                  driver.execute_script("arguments[0].style.display='block';", driver.find_element(By.ID, "menu_1_sub"))
                  overview_link = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='usersiteoverviews']"))
                  )
                  overview_link.click()
            except Exception as e:
                  print(f"Không tìm thấy nút report, lỗi {e}")

            # Chờ bảng môn học tải
            try:
                  table = WebDriverWait(driver, 40).until(
                        EC.presence_of_element_located((By.ID, "ourptlistcourse"))
                  )
            except Exception as e:
                  print(f"Không tìm thấy bảng, lỗi {e}")

            # Lấy tất cả các dòng trong bảng
            try:
                  rows = WebDriverWait(driver, 20).until(
                        EC.presence_of_all_elements_located((By.XPATH, ".//tr"))
                  )
            except Exception as e:
                  print(f"Không tìm thấy các dòng, lỗi {e}")
            
            get_subject = []
            try:
                  # Lặp qua từng dòng để lấy dữ liệu
                  for row in rows:
                        temp = {}
                        cells = row.find_elements(By.XPATH, ".//td")
                        for idx_cell, cell in enumerate(cells):
                              # Kiểm tra nếu ô chứa thông tin môn học của học kỳ hiện tại
                              if cell.text.startswith("[" + semester + "]"):
                                    # Lấy các giá trị từ các ô liên quan
                                    temp["lms_id"] = cells[idx_cell - 1].text
                                    temp["count_total_student"] = cells[idx_cell + 2].text
                                    temp["count_student_access"] = cells[idx_cell + 4].text
                                    temp["percent_student_access"] = cells[idx_cell + 6].text
                                    temp["percent_teacher_access"] = cells[idx_cell + 7].text
                                    
                                    # Dùng regex để trích xuất thông tin từ text
                                    # Format: [Kỳ] SUBID - Tên môn (Room-GroupID) ... Giảng viên ..., Tên GV
                                    pattern = rf"\[{semester}\]\s+(\w+)\s+-\s+(.*?)\s+\((\w+)-(\w+)\).*?Giảng viên\s*\d+\.\s*([^\n]+)"
                                    match = re.search(pattern, cell.text, re.DOTALL)
                                    if match:
                                          temp["subject_id"] = match.group(1)      # FINA2343
                                          temp["subject_name"] = match.group(2)    # Tên môn
                                          temp["teacher_id"] = match.group(3)      # TN303
                                          temp["group_id"] = match.group(4)        # 01
                                          temp["teacher_name"] = match.group(5)    # Nguyễn Văn A

                                    get_subject.append(temp)
                  driver.quit()
                  return get_subject 
            except Exception as e:
                  print(f"Không xuất được nội dung từng ô trong bảng, lỗi {e}")


      def process_get_detail_phdt(self):
            """
            Truy cập PHDT để lấy thông tin giảng viên và môn học
            Thông tin lấy được:
            - Key: Mã môn học - Mã lớp
            - Value: [Mã giảng viên, Tên giảng viên]
            
            Quy trình:
            1. Đăng nhập vào PHDT
            2. Lấy danh sách đơn vị tổ chức từ API
            3. Truy cập trang thời khóa biểu của từng đơn vị với học kỳ hiện tại
            4. Chờ bảng thời khóa biểu tải và trích xuất dữ liệu
            5. Lưu thông tin vào dict get_teacher_and_subject
            6. Đóng trình duyệt và trả về dict kết quả

            Args: None
            Return: dict - Thông tin giảng viên và môn học từ PHDT
            """

            driver = self.login_selenium("https://phdt.ou.edu.vn/")
            api_handler = APIHandler()
            # Lấy danh sách tất cả các đơn vị tổ chức
            list_unit = api_handler.get_unit()
            get_teacher_and_subject = {}

            # Lặp qua từng đơn vị
            for unit in list_unit:
                  # Truy cập trang thời khóa biểu của đơn vị với học kỳ hiện tại
                  driver.get(f"https://phdt.ou.edu.vn/admin/tidt/sodaubai/admin/tkbcoso_listtkb?nhhk={"".join(["20", semester])}&madp={unit[0]}")
                  
                  # Chờ bảng thời khóa biểu tải
                  try:
                        table = WebDriverWait(driver, 40).until(
                              EC.presence_of_element_located((By.ID, "table_tkbcs"))
                        )
                  except Exception as e:
                        print(f"Không tìm thấy bảng, lỗi {e}")            

                  # Lấy tất cả các dòng trong bảng
                  try:
                        rows = WebDriverWait(driver, 20).until(
                              EC.presence_of_all_elements_located((By.XPATH, ".//tr"))
                        )
                  except Exception as e:
                        print(f"Không tìm thấy các dòng, lỗi {e}")

                  # Trích xuất dữ liệu từ bảng
                  try:
                        for row in rows:
                              cells = row.find_elements(By.XPATH, ".//td")
                              # Đảm bảo cells không rỗng trước khi truy cập
                              if cells:
                                    # for idx_cell, cell in enumerate(cells):
                                          # Key: Mã môn học - Mã lớp
                                          # Value: [Mã giảng viên, Tên giảng viên]
                                    get_teacher_and_subject["-".join([cells[3].text, cells[2].text])] = [cells[8].text, cells[9].text]
                  except Exception as e:
                        print(f"Không in được dữ liệu, lỗi {e}")
                  
            driver.quit()
            return get_teacher_and_subject
