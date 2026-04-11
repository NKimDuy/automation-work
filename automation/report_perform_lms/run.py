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
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import time 
import re

class ReportPerformLMS:
      def test(self):
            start_selenium = InitSelenium()
            list_lsa = start_selenium.process_get_detail_lsa()
            try:
                  for ls in list_lsa:
                        if ls["group_id"] and ls["subject_id"]:
                              print("-".join([ls["group_id"], ls["subject_id"]]))   
            except Exception as e:
                  print(f"An error occurred: {e}")


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
            try:
                  for item in list_lsa:
                        key = "-".join([item["group_id"], item["subject_id"]])
                        get_group_subject_in_lsa.append(key)
            except Exception as e:
                  print(f"Không có dữ liệu môn học: {e}")

            for key_sb, value_sb in list_report_lms.items():
                  if key_sb in get_group_subject_in_lsa:
                        list_report_lms[key_sb].append("x")
                  else:
                        list_report_lms[key_sb].append("")
            
            return list_report_lms


      def decor_report_perform_lms(self):
            wb = Workbook()
            if "Sheet" in wb.sheetnames:
                  wb.remove(wb["Sheet"])

            sheet_general = wb.create_sheet("Tổng quan")
            
            data = [
                  {'department': "TX.LALA", 'group': "SG001", 'id_subject': "ACCO4331", 'name_subject': "Quản trị học",'id_class': "TM123456", 'id_unit': "TM", 'name_unit': "Đồng Tháp Mười Long An", 'id_teacher': "TX001", 'name_teacher': "Nguyễn Kim Duy", 'has_lms': "x"},
                  {'department': "TX.QTQT","group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "id_class": "TM123456", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", "has_lms": "x"},
                  {'department': "TX.LALA","group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "id_class": "TM123456", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", "has_lms": ""},
                  {'department': "TX.LALA","group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "id_class": "TM123456", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", "has_lms": "x"},
                  {'department': "TX.LALA","group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "id_class": "TM123456", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", "has_lms": ""},
                  {'department': "TX.QTQT","group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "id_class": "TM123456", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", "has_lms": "x"},
                  {'department': "TX.QTQT","group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "id_class": "TM123456", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", "has_lms": ""}
            ]
            summary_report = []
            department_stats = {}
            footer_sheet_general = ["Tổng cộng", 0, 0, 0]
            for dic in data:
                  dept = dic['department']
                  has_lms = dic['has_lms']
                  
                  if dept not in department_stats:
                        department_stats[dept] = {'count': 0, 'has_x': 0, 'no_x': 0}
                  
                  department_stats[dept]['count'] += 1
                  if has_lms == 'x':
                        department_stats[dept]['has_x'] += 1
                  else:
                        department_stats[dept]['no_x'] += 1

            for dept, stats in department_stats.items():
                  summary_report.append([
                        dept,
                        stats['count'],
                        stats['has_x'],
                        stats['no_x']
                  ])

            for department in summary_report:
                  footer_sheet_general[1] += department[1]
                  footer_sheet_general[2] += department[2]
                  footer_sheet_general[3] += department[3]
            
            header_font = Font(name="Times New Roman", size=12, bold=True)
            footer_font = Font(name="Times New Roman", size=12, bold=True)
            header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            header_fill = PatternFill(start_color="97FFFF", end_color="97FFFF", fill_type="solid")
            data_font = Font(name="Times New Roman", size=11)
            data_alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
            border_style = Border(
                  left = Side(style="thin"),
                  right = Side(style="thin"),
                  top = Side(style="thin"),
                  bottom = Side(style="thin")
            )

            header_sheet_general = ["Khoa", "Tổng số nhóm môn học", "Đã soạn LMS", "Chưa soạn LMS"]
            sheet_general.append(header_sheet_general)
            for col_num, header in enumerate(header_sheet_general, 1):
                  cell = sheet_general.cell(row=1, column=col_num)
                  cell.font = header_font
                  cell.fill = header_fill
                  cell.alignment = header_alignment
                  cell.border = border_style

            for row_num, row_data in enumerate(summary_report, 2):
                  for col_num, value in enumerate(row_data, 1):
                        cell = sheet_general.cell(row=row_num, column=col_num)
                        cell.value = value
                        cell.font = data_font
                        cell.alignment = data_alignment
                        cell.border = border_style
            
            sheet_general.append(footer_sheet_general)
            for col_num, footer in enumerate(footer_sheet_general, 1):
                  cell = sheet_general.cell(row=len(summary_report) + 2, column=col_num)
                  cell.font = footer_font
                  cell.alignment = data_alignment
                  cell.border = border_style


            for col_num in range(1, len(header_sheet_general) + 1):
                  column_letter = get_column_letter(col_num)
                  max_length = 0
                  
                  # Duyệt qua tất cả ô trong cột
                  for row in sheet_general[column_letter]:
                        try:
                              if len(str(row.value)) > max_length:
                                    max_length = len(str(row.value))
                        except:
                              pass
                  
                  # Set độ rộng với padding
                  adjusted_width = min(max_length + 10, 100)  # Tối đa 100 ký tự
                  sheet_general.column_dimensions[column_letter].width = adjusted_width

            sheet_detail = wb.create_sheet("Chi tiết")

            wb.save(os.path.join(os.getcwd(), "report_perform_lms.xlsx"))

ob = ReportPerformLMS()
ob.decor_report_perform_lms()
# ob.test()
# print(ob.test())
# print(ob.report_perform_lms("2026-02-01", "2026-02-25"))
# print(ob.get_department_of_subject())