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
import os
import time 
import re

class ReportPerformLMS:
      def test(self):
            start_selenium = InitSelenium()
            start_selenium.process_get_detail_phdt()
            #start_selenium.test_init()
            

      def report_perform_lms(self, from_day, to_day):
            list_report_lms = []
            api_handler = APIHandler()
            list_unit = api_handler.get_unit()
            for unit in list_unit:
                  list_subject = api_handler.get_subject_from_api(semester, unit[0])
                  for subject in list_subject:
                        if subject["TUNGAYTKB"] is not None:
                              if from_day <= datetime.strptime(subject["TUNGAYTKB"], "%Y-%m-%d") <= to_day:
                                    list_report_lms.extend([
                                          subject["NhomTo"],
                                          subject["MaMH"],
                                          subject["TenMH"],
                                          subject["MaLop"],
                                          subject["MaDP"],
                                          subject["TenDP"]
                                    ])

ob = ReportPerformLMS()
# print(ob.test())
ob.test()