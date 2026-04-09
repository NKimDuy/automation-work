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
from utils.api import APIHandler
from config.settings import semester
from datetime import datetime
from openpyxl import Workbook, load_workbook
import os
import time 
import re

class ReportPerformLMS:
      def test(self):
            # start_selenium = InitSelenium()
            # start_selenium.process_get_detail_phdt()
            api_handler = APIHandler()
            print(api_handler.get_subject_from_api("251", "SG"))
            # list_unit = api_handler.get_unit()
            # for unit in list_unit:
            #       list_subject = api_handler.get_subject_from_api(semester, unit[0])
            #       for subject in list_subject:
            #             print(subject)
            

      def get_department_of_subject(self):
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
            file = os.path.join(base_dir, "config", "excel", "mon_hoc_khoa.xlsx")
            wb = load_workbook(file)
            ws = wb.active
            dict_subject_department = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                  subject_code = row[0]
                  department = row[1]
                  dict_subject_department[subject_code] = department
            return dict_subject_department
            # return file

      def report_perform_lms(self, from_day, to_day):
            start_selenium = InitSelenium()
            get_info_teacher = start_selenium.process_get_detail_phdt()

            api_handler = APIHandler()
            list_unit = api_handler.get_unit()

            from_day = datetime.strptime(from_day, "%Y-%m-%d")
            to_day = datetime.strptime(to_day, "%Y-%m-%d")

            list_report_lms = {}

            for unit in list_unit:
                  list_subject = api_handler.get_subject_from_api(semester, unit[0])
                  for subject in list_subject:
                        if subject["TUNGAYTKB"] is not None:
                              if from_day <= datetime.strptime(subject["TUNGAYTKB"], "%Y-%m-%d") <= to_day:
                                    key = "-".join([subject["NhomTo"], subject["MaMH"]])
                                    if key not in list_report_lms:
                                          list_report_lms[key] = [
                                                self.get_department_of_subject().get(subject["MaMH"], "Không xác định"),
                                                subject["NhomTo"],
                                                subject["MaMH"],
                                                subject["TenMH"],
                                                subject["MaLop"],
                                                subject["MaDP"],
                                                subject["TenDP"],
                                                get_info_teacher[key][0],
                                                get_info_teacher[key][1]
                                          ]
                                    else:
                                          list_report_lms[key][3] = ",".join([list_report_lms[key][3], subject["MaLop"]])  

            list_lsa = start_selenium.process_get_detail_lsa()
            get_group_subject_in_lsa = []
            for item in list_lsa:
                  key = "-".join([item["group_id"], item["subject_id"]])
                  get_group_subject_in_lsa.append(key)

            for key_sb, value_sb in list_report_lms.items():
                  if key_sb in get_group_subject_in_lsa:
                        list_report_lms[key_sb].append("x")
                  else:
                        list_report_lms[key_sb].append("")
            
            return list_report_lms



ob = ReportPerformLMS()
# ob.test()
# print(ob.test())
print(ob.report_perform_lms("2026-03-16", "2026-03-22"))
# print(ob.get_department_of_subject())