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

class ScoreLMS:
    def __init__(self):
        # self.url_lms = f"https://lms.oude.edu.vn/{semester}/course/search.php"
        self.url_lms = f"https://lms.oude.edu.vn/251/course/search.php"
    
    def check_forum(self, driver, forum_link, ten_giang_vien):
        second_tab = driver.current_window_handle
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
                    ten_reply = ""


                print(f"Tên SV: {ten_sv}, Tên GV: {ten_giang_vien}, Tên reply: {ten_reply}")
                if ten_giang_vien.lower() not in ten_reply.lower():
                    chua_reply.append(ten_sv)

            driver.close()
            driver.switch_to.window(second_tab)
            return chua_reply

    def score_lms(self):
        start_selenium = InitSelenium()
        driver = start_selenium.login_selenium(self.url_lms)

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

            forums = {}
            quiz = {}
            assign = {}
            
            for link_subject in list_link_subject:
                driver.execute_script("window.open(arguments[0]);", link_subject)
                driver.switch_to.window(driver.window_handles[-1])

                full_text = driver.find_element(By.TAG_NAME, "body").text
                print(full_text)

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
                       
                print(f"forum links : {forums}")
                # print(f"quiz links  : {quiz}")
                # print(f"assign links: {assign}")

                if forums:
                    forum_chua_day_du = []  # [(ten_forum, [sv_chua_reply])]
                    for name, forum_link in forums.items():
                        chua_reply = self.check_forum(driver, forum_link, "Võ Thị Hồng Nhung")
                        if chua_reply:
                            forum_chua_day_du.append({
                                "ten_forum": name,
                                "sv_chua_reply": chua_reply
                            })
                    if not forum_chua_day_du:
                        print(f" ✅ Có diễn đàn và tương tác với tất cả sinh viên")
                    else:
                        print(f" ⚠️ Có diễn đàn nhưng không tương tác đầy đủ với sinh viên")
                        for f in forum_chua_day_du:
                            print(f"  📌 Forum: {f['ten_forum']}")
                            print(f"     SV chưa được reply: {f['sv_chua_reply']}")
                else:
                    print(f" ❌ Không có diễn đàn nào")    

                time.sleep(2)  # Thêm delay nếu cần thiết để tránh bị khóa tài khoản do hoạt động quá nhanh
                
                driver.close()
                # driver.get(link_subject)
                driver.switch_to.window(main_tab)
                break
                

        except Exception as e:
                  print(f"Không mở được tab mới, lỗi {e}")

test = ScoreLMS()
test.score_lms()