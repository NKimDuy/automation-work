# =============================================================================
# MODULE: modules/init_selenium.py
# MỤC ĐÍCH: Wrapper trung tâm cho toàn bộ thao tác Selenium trong dự án.
#           Mọi workflow đều khởi tạo InitSelenium() và dùng các hàm ở đây.
#
# CÁC CHỨC NĂNG CHÍNH:
#   - init_selenium()            : tạo Chrome driver (chế độ headless, ẩn automation)
#   - login_selenium(url)        : đăng nhập SSO vào LMS hoặc PHDT
#   - process_get_detail_lsa()   : scraping bảng môn học từ LSA
#   - process_get_detail_phdt()  : scraping thời khóa biểu từ PHDT để lấy thông tin GV
#
# THÔNG TIN ĐĂNG NHẬP:
#   Lấy từ config/settings.py: user_lms, password_lms, captcha (đã gitignore)
# =============================================================================

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
      """Lớp quản lý tương tác với trình duyệt Chrome thông qua Selenium.

      Chức năng chính: đăng nhập SSO, lấy thông tin môn học từ LSA và PHDT.
      Tất cả workflow trong dự án đều dùng class này làm điểm khởi đầu
      để tương tác với các hệ thống web.
      """

      def __init__(self):
            """Khởi tạo URL đăng nhập LSA."""
            # URL trang đăng nhập LSA — dùng trong process_get_detail_lsa()
            self.url_lsa = "http://lsa.ou.edu.vn/auth/login"


      def init_selenium(self):
            """Tạo và trả về Chrome WebDriver với cấu hình tối ưu cho automation.

            Cấu hình bao gồm:
              - Headless mode: chạy ẩn, không hiển thị cửa sổ trình duyệt
              - Tắt AutomationControlled: giảm khả năng bị phát hiện là bot
              - User-agent giả: giả lập trình duyệt thật để tránh bị chặn
              - Kích thước cửa sổ 1920x1080: đảm bảo hiển thị đúng layout

            Trả về:
                webdriver.Chrome: driver đã được cấu hình, sẵn sàng dùng.
            """
            chrome_options = Options()
            # Headless mode mới (stable từ Chrome 112+): chạy ẩn không mở cửa sổ
            chrome_options.add_argument("--headless=new")
            # Tắt flag báo hiệu đây là trình duyệt automation → tránh bị website chặn
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            # Giả lập user-agent của Chrome thật trên Windows để tránh bị detect
            chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36")
            # Kích thước cửa sổ ảo — quan trọng vì headless mode có thể render khác
            chrome_options.add_argument("--window-size=1920,1080")
            # Loại bỏ flag "enable-automation" khỏi danh sách switch của Chrome
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            # Tắt extension automation của Chrome
            chrome_options.add_experimental_option("useAutomationExtension", False)

            # Khởi tạo driver với Chrome mặc định của hệ thống (không cần chỉ định service)
            driver = webdriver.Chrome()
            return driver


      def login_selenium(self, url):
            """Đăng nhập vào hệ thống web qua luồng SSO (Single Sign-On) của trường OU.

            Luồng đăng nhập SSO gồm 7 bước:
              1. Mở URL đích (LMS hoặc PHDT)
              2. Nhấn nút "HCMCOU-SSO" để chuyển sang trang SSO
              3. Chọn loại người dùng "Cán bộ-Nhân viên / Giảng viên" từ dropdown
              4. Nhập tên đăng nhập (user_lms từ settings.py)
              5. Nhập mật khẩu (password_lms từ settings.py)
              6. Nhập mã CAPTCHA (captcha từ settings.py — cần cập nhật thủ công)
              7. Nhấn nút "Đăng nhập" và xử lý popup cho phép

            Lưu ý về CAPTCHA:
              - Giá trị captcha được hardcode trong settings.py.
              - Nếu CAPTCHA hết hạn hoặc thay đổi, cần cập nhật lại settings.py thủ công.
              - Mỗi bước đều được bọc trong try/except riêng để không dừng hẳn nếu 1 bước lỗi.

            Tham số:
                url (str): URL trang cần truy cập.
                           Ví dụ: "https://lms.oude.edu.vn/251/course/search.php"
                                  "https://phdt.ou.edu.vn/"

            Trả về:
                webdriver.Chrome: driver đã đăng nhập thành công, sẵn sàng thao tác tiếp.
            """
            driver = self.init_selenium()
            driver.get(url)

            # Bước 1: Nhấn nút "HCMCOU-SSO" để chuyển sang trang đăng nhập SSO
            try:
                  button_sso_text = "//button[contains(text(), 'HCMCOU-SSO')]"
                  button_sso = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, button_sso_text))
                  )
                  button_sso.click()
            except Exception as e:
                  print(f"Không thấy nút HCMCOU-SSO, lỗi: {e}")

            # Bước 2: Chọn loại người dùng từ dropdown (ID="form-usertype")
            # Chọn "Cán bộ-Nhân viên / Giảng viên" để phân biệt với tài khoản sinh viên
            try:
                  dropdown = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "form-usertype"))
                  )
                  select_type_user = Select(dropdown)
                  select_type_user.select_by_visible_text("Cán bộ-Nhân viên / Giảng viên")
            except Exception as e:
                  print(f"Không thấy lựa chọn Cán bộ nhân viên, lỗi: {e}")

            # Bước 3: Nhập tên đăng nhập vào ô input (ID="form-username")
            try:
                  username_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "form-username"))
                  )
                  username_input.send_keys(user_lms)
            except Exception as e:
                  print(f"Không thấy ô nhập tài khoản, lỗi: {e}")

            # Bước 4: Nhập mật khẩu vào ô input (ID="form-password")
            try:
                  password_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "form-password"))
                  )
                  password_input.send_keys(password_lms)
            except Exception as e:
                  print(f"Không thấy ô nhập mật khẩu, lỗi: {e}")

            # Bước 5: Nhập CAPTCHA — giá trị lấy từ settings.py, cần cập nhật thủ công
            try:
                  captcha_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "form-captcha"))
                  )
                  captcha_input.send_keys(captcha)
            except Exception as e:
                  print(f"Không tìm thấy ô để nhập Capcha, lỗi {e}")

            # Bước 6: Nhấn nút "Đăng nhập" — selector rộng để bắt cả button, a, input
            try:
                  button_login_text = "//*[self::button or self::a or self::input][contains(., 'Đăng nhập') or contains(@value, 'Đăng nhập')]"
                  button_login = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, button_login_text))
                  )
                  button_login.click()
            except Exception as e:
                  print(f"Không tìm thấy nút đăng nhập, lỗi {e}")

            # Bước 7: Nhấn nút "Cho phép/Đồng ý" nếu hệ thống yêu cầu xác nhận quyền truy cập
            # Timeout 5s (ngắn hơn) vì bước này không phải lúc nào cũng xuất hiện
            try:
                  button_allow = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn.btn-success.btn-approve"))
                  )
                  button_allow.click()
            except Exception as e:
                  print(f"Không tìm thấy nút cho phép nhấn đồng ý, lỗi {e}")

            return driver


      def process_get_detail_lsa(self):
            """Đăng nhập LSA và scraping bảng danh sách môn học trong học kỳ hiện tại.

            Luồng xử lý:
              1. Đăng nhập vào LSA qua SSO
              2. Chọn học kỳ từ dropdown "[LIVE] LMS ĐTTX <semester>"
              3. Mở menu ẩn và click vào link report/overview
              4. Chờ bảng môn học (ID="ourptlistcourse") tải xong
              5. Duyệt từng dòng trong bảng, tìm ô bắt đầu bằng "[<semester>]"
              6. Dùng regex để trích xuất: mã môn, tên môn, mã GV, nhóm, tên GV
              7. Đóng driver và trả về danh sách

            Regex pattern dùng để parse text ô môn học:
              "[251] FINA2343 - Tài chính doanh nghiệp (TN303-01) ... Giảng viên 1. Nguyễn Văn A"
              → subject_id = "FINA2343"
              → subject_name = "Tài chính doanh nghiệp"
              → teacher_id = "TN303"   (đây thực ra là room/phòng, không phải mã GV)
              → group_id = "01"
              → teacher_name = "Nguyễn Văn A"

            Trả về:
                list[dict]: danh sách môn học, mỗi dict gồm:
                    - lms_id               : ID môn trên LMS (số nguyên dạng string)
                    - subject_id           : mã môn học, ví dụ "FINA2343"
                    - subject_name         : tên môn học
                    - teacher_id           : mã phòng học / giảng viên
                    - group_id             : mã nhóm, ví dụ "01"
                    - teacher_name         : tên giảng viên
                    - count_total_student  : tổng số sinh viên
                    - count_student_access : số SV đã truy cập LMS
                    - percent_student_access : % SV truy cập
                    - percent_teacher_access : % GV truy cập
            """
            driver = self.login_selenium(self.url_lsa)

            # Bước 1: Chọn học kỳ từ dropdown (ID="moodlesiteid")
            # Text option dạng: "[LIVE] LMS ĐTTX 251"
            try:
                  dropdown_semester = WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.ID, "moodlesiteid"))
                  )
                  select_type_semester = Select(dropdown_semester)
                  select_type_semester.select_by_visible_text(" ".join(["[LIVE] LMS ĐTTX", semester]))
            except Exception as e:
                  print(f"Không tìm thấy dropdownlist thể hiện học kỳ, lỗi {e}")

            # Bước 2: Hiển thị menu ẩn bằng JavaScript rồi click vào link overview
            # Menu bị ẩn bằng CSS display:none → cần execute_script để hiển thị
            try:
                  driver.execute_script("arguments[0].style.display='block';", driver.find_element(By.ID, "menu_1_sub"))
                  overview_link = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='usersiteoverviews']"))
                  )
                  overview_link.click()
            except Exception as e:
                  print(f"Không tìm thấy nút report, lỗi {e}")

            # Bước 3: Chờ bảng môn học tải — timeout 40s vì bảng có thể rất lớn
            try:
                  table = WebDriverWait(driver, 40).until(
                        EC.presence_of_element_located((By.ID, "ourptlistcourse"))
                  )
            except Exception as e:
                  print(f"Không tìm thấy bảng, lỗi {e}")

            # Bước 4: Lấy tất cả dòng <tr> trong bảng
            try:
                  rows = WebDriverWait(driver, 20).until(
                        EC.presence_of_all_elements_located((By.XPATH, ".//tr"))
                  )
            except Exception as e:
                  print(f"Không tìm thấy các dòng, lỗi {e}")

            get_subject = []
            try:
                  for row in rows:
                        temp = {}
                        # Lấy tất cả ô <td> trong dòng hiện tại
                        cells = row.find_elements(By.XPATH, ".//td")
                        for idx_cell, cell in enumerate(cells):
                              # Tìm ô nào bắt đầu bằng "[251]" — đó là ô chứa thông tin môn học
                              if cell.text.startswith("[" + semester + "]"):
                                    # Các ô xung quanh chứa số liệu thống kê:
                                    # idx_cell - 1 : lms_id
                                    # idx_cell + 2 : tổng số SV
                                    # idx_cell + 4 : số SV đã truy cập
                                    # idx_cell + 6 : % SV truy cập
                                    # idx_cell + 7 : % GV truy cập
                                    temp["lms_id"] = cells[idx_cell - 1].text
                                    temp["count_total_student"] = cells[idx_cell + 2].text
                                    temp["count_student_access"] = cells[idx_cell + 4].text
                                    temp["percent_student_access"] = cells[idx_cell + 6].text
                                    temp["percent_teacher_access"] = cells[idx_cell + 7].text

                                    # Regex parse nội dung ô môn học
                                    # Format: "[251] FINA2343 - Tên môn (Room-Nhóm) ... Giảng viên 1. Tên GV"
                                    pattern = rf"\[{semester}\]\s+(\w+)\s+-\s+(.*?)\s+\((\w+)-(\w+)\).*?Giảng viên\s*\d+\.\s*([^\n]+)"
                                    match = re.search(pattern, cell.text, re.DOTALL)
                                    if match:
                                          temp["subject_id"] = match.group(1)    # mã môn, ví dụ "FINA2343"
                                          temp["subject_name"] = match.group(2)  # tên môn học
                                          temp["teacher_id"] = match.group(3)    # mã phòng/nhóm, ví dụ "TN303"
                                          temp["group_id"] = match.group(4)      # nhóm, ví dụ "01"
                                          temp["teacher_name"] = match.group(5)  # tên giảng viên

                                    get_subject.append(temp)
                  driver.quit()
                  return get_subject
            except Exception as e:
                  print(f"Không xuất được nội dung từng ô trong bảng, lỗi {e}")


      def process_get_detail_phdt(self):
            """Đăng nhập PHDT và scraping thời khóa biểu để lấy thông tin giảng viên.

            PHDT (Phòng Hành Chính - Đào Tạo) có trang thời khóa biểu theo từng đơn vị,
            chứa thông tin liên kết giữa nhóm-môn học và giảng viên phụ trách.

            Luồng xử lý:
              1. Đăng nhập PHDT qua SSO
              2. Gọi API lấy danh sách tất cả đơn vị (get_unit())
              3. Với mỗi đơn vị, truy cập URL TKB:
                 https://phdt.ou.edu.vn/.../tkbcoso_listtkb?nhhk=20251&madp=<unit_id>
              4. Chờ bảng (ID="table_tkbcs") tải xong
              5. Duyệt từng dòng, lấy: nhóm (cells[3]), mã môn (cells[2]),
                 mã GV (cells[8]), tên GV (cells[9])
              6. Lưu vào dict với key = "nhóm-mã_môn"
              7. Nếu lỗi, thử lại tối đa 3 lần trước khi bỏ qua đơn vị đó

            Cấu trúc bảng PHDT (index cột):
              cells[0]: STT
              cells[1]: Tên môn học
              cells[2]: Mã môn học    → dùng làm key
              cells[3]: Nhóm môn      → dùng làm key
              cells[4-7]: Thông tin thời gian/phòng
              cells[8]: Mã giảng viên → value[0]
              cells[9]: Tên giảng viên → value[1]

            Trả về:
                dict: { "nhóm-mã_môn": [mã_GV, tên_GV], ... }
                      Ví dụ: { "SG001-ACCO4331": ["TX001", "Nguyễn Văn A"] }
            """
            driver = self.login_selenium("https://phdt.ou.edu.vn/")
            api_handler = APIHandler()
            # Lấy toàn bộ danh sách đơn vị để duyệt qua từng đơn vị
            list_unit = api_handler.get_unit()
            get_teacher_and_subject = {}

            for unit in list_unit:
                  # Thử lại tối đa 3 lần nếu bảng không tải được (mạng chậm hoặc trang lỗi)
                  for i in range(3):
                        try:
                              # URL TKB của đơn vị: nhhk = "20" + semester (ví dụ "20251")
                              driver.get(f"https://phdt.ou.edu.vn/admin/tidt/sodaubai/admin/tkbcoso_listtkb?nhhk={"".join(["20", semester])}&madp={unit[0]}")

                              # Chờ bảng TKB tải — timeout 40s vì dữ liệu có thể lớn
                              try:
                                    table = WebDriverWait(driver, 40).until(
                                          EC.presence_of_element_located((By.ID, "table_tkbcs"))
                                    )
                              except Exception as e:
                                    print(f"Không tìm thấy bảng, lỗi {e}")

                              # Lấy tất cả dòng <tr> trong bảng
                              try:
                                    rows = WebDriverWait(driver, 20).until(
                                          EC.presence_of_all_elements_located((By.XPATH, ".//tr"))
                                    )
                              except Exception as e:
                                    print(f"Không tìm thấy các dòng, lỗi {e}")

                              # Duyệt từng dòng, bỏ qua dòng header (không có <td>)
                              try:
                                    for row in rows:
                                          cells = row.find_elements(By.XPATH, ".//td")
                                          if cells:
                                                # Key: "nhóm-mã_môn", ví dụ "SG001-ACCO4331"
                                                # Value: [mã_GV, tên_GV]
                                                get_teacher_and_subject["-".join([cells[3].text, cells[2].text])] = [cells[8].text, cells[9].text]
                              except Exception as e:
                                    print(f"Không in được dữ liệu, lỗi {e}")

                              break  # Thành công → thoát vòng thử lại, sang đơn vị tiếp theo

                        except Exception as e:
                              print(f"Lần {i + 1} thất bại")
                              time.sleep(2)  # Đợi 2 giây trước khi thử lại

            driver.quit()
            return get_teacher_and_subject
