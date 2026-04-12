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
from datetime import datetime, timedelta
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


      def get_last_week_range(self):
            today = datetime.today()

            # Lùi về thứ 2 tuần này (weekday: 0=T2, 6=CN)
            days_since_monday = today.weekday()  # hôm nay là thứ mấy tính từ T2

            last_monday = today - timedelta(days=days_since_monday + 7)  # T2 tuần trước
            last_sunday = last_monday + timedelta(days=6)                # CN tuần trước

            return last_monday.strftime("%Y-%m-%d"), last_sunday.strftime("%Y-%m-%d")


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

            list_lsa = start_selenium.process_get_detail_lsa()
            get_group_subject_in_lsa = []
            try:
                  for item in list_lsa:
                        key = "-".join([item["group_id"], item["subject_id"]])
                        get_group_subject_in_lsa.append(key)
            except Exception as e:
                  print(f"Không có dữ liệu môn học: {e}")

            api_handler = APIHandler()
            list_unit = api_handler.get_unit()

            from_day = datetime.strptime(from_day, "%Y-%m-%d")
            to_day = datetime.strptime(to_day, "%Y-%m-%d")

            department_dic = {
                  "TX.NNNN": "Ngoại ngữ",
                  "TX.LALA": "Luật",
                  "TX.LA": "Luật",
                  "TX.CBML": "Cơ bản",
                  "TX.CBCB": "Cơ bản",
                  "TX.XHXH": "XHH-CTXH-ĐNA",
                  "TX.KKKK": "Kế toán - Kiểm toán",
                  "TX.QTQT": "Quản Trị Kinh Doanh",
                  "TX.TCTC": "Tài chính - Ngân hàng",
                  "TX.SHSH": "Công Nghệ Sinh Học",
                  "TX.KIKI": "Kinh tế và quản lý công",
                  "TX.KTKT": "Xây dựng",
            }

            list_report_lms = {}

            for unit in list_unit:
                  list_subject = api_handler.get_subject_from_api(semester, unit[0])
                  for subject in list_subject:
                        if subject["TUNGAYTKB"] is not None:
                              if from_day <= datetime.strptime(subject["TUNGAYTKB"], "%Y-%m-%d") <= to_day:
                                    key_department = self.get_department_of_subject().get(subject["MaMH"], "Không xác định")
                                    value_department = department_dic.get(key_department, "Không xác định")
                                    key = "-".join([subject["NhomTo"], subject["MaMH"]])
                                    if key not in list_report_lms:
                                          list_report_lms[key] = [
                                                value_department,
                                                subject["NhomTo"],
                                                subject["MaMH"],
                                                subject["TenMH"],
                                                subject["MaLop"],
                                                subject["MaDP"],
                                                subject["TenDP"],
                                                get_info_teacher[key][0],
                                                get_info_teacher[key][1],
                                                from_day.strftime("%d-%m-%Y")
                                          ]
                                    else:
                                          list_report_lms[key][3] = ",".join([list_report_lms[key][3], subject["MaLop"]])  

            for key_sb, value_sb in list_report_lms.items():
                  if key_sb in get_group_subject_in_lsa:
                        list_report_lms[key_sb].append("x")
                  else:
                        list_report_lms[key_sb].append("")
            
            list_lms_add_to_report = []
            for key, value in list_report_lms.items():
                  dict_subject = {
                        "department": value[0],
                        "group": value[1],
                        "id_subject": value[2],
                        "name_subject": value[3],
                        "id_class": value[4],
                        "id_unit": value[5],
                        "name_unit": value[6],
                        "id_teacher": value[7],
                        "name_teacher": value[8],
                        "from_day": value[9],
                        "has_lms": value[10]
                  }
                  list_lms_add_to_report.append(dict_subject)
            self.decor_report_perform_lms(list_lms_add_to_report, from_day.strftime("%d-%m-%Y"), to_day.strftime("%d-%m-%Y"))


      def set_dimension_column(self, name_sheet, title_header):
            for col_num in range(1, len(title_header) + 1):
                  column_letter = get_column_letter(col_num)
                  max_length = 0
                  
                  # Duyệt qua tất cả ô trong cột
                  for row in name_sheet[column_letter]:
                        try:
                              if len(str(row.value)) > max_length:
                                    max_length = len(str(row.value))
                        except:
                              pass
                  
                  # Set độ rộng với padding
                  adjusted_width = min(max_length + 10, 100)  # Tối đa 100 ký tự
                  name_sheet.column_dimensions[column_letter].width = adjusted_width


      def decor_report_perform_lms(self, data, from_day, to_day):
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
            folder_track_lms = os.path.join(base_dir, "data", "output", "report_perform_lms", semester)
            os.makedirs(folder_track_lms, exist_ok=True)
            flle_track_no_lms = os.path.join(folder_track_lms, "__danh_sach_khong_thuc_hien_lms.xlsx")
            name_file = f"BC Tình hình soạn thảo LMS HK {semester} từ ngày ({from_day} đến ngày {to_day}).xlsx"
            # name_file = f"test.xlsx"
            file_report_lms = os.path.join(folder_track_lms, name_file)
            
            wb_report_lms = Workbook()
            if "Sheet" in wb_report_lms.sheetnames:
                  wb_report_lms.remove(wb_report_lms["Sheet"])

            sheet_general = wb_report_lms.create_sheet("Tổng quan")
            
            # data = [
            #       {'department': "TX.LALA", 'group': "SG001", 'id_subject': "ACCO4331", 'name_subject': "Quản trị học", 'id_class': "TM123456", 'id_unit': "TM", 'name_unit': "Đồng Tháp Mười Long An", 'id_teacher': "TX001", 'name_teacher':  "Nguyễn Kim Duy",'from_day': '2026-06-06', 'has_lms': "x"},
            #       {'department': "TX.QTQT", "group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "id_class": "TM123456", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", 'from_day': '2026-06-06', "has_lms": "x"},
            #       {'department': "TX.LALA", "group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "id_class": "TM123456", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", 'from_day': '2026-06-06', "has_lms": ""},
            #       {'department': "TX.LALA", "group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "id_class": "TM123456", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", 'from_day': '2026-06-06', "has_lms": "x"},
            #       {'department': "TX.LALA", "group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "id_class": "TM123456", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", 'from_day': '2026-06-06', "has_lms": ""},
            #       {'department': "TX.QTQT", "group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "id_class": "TM123456", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", 'from_day': '2026-06-06', "has_lms": "x"},
            #       {'department': "TX.QTQT", "group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "id_class": "TM123456", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", 'from_day': '2026-06-06', "has_lms": ""},
            # ]
            department_stats = {}
            summary_report = []
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
            
            self.set_dimension_column(sheet_general, header_sheet_general)
            
            sheet_detail = wb_report_lms.create_sheet("Chi tiết")

            header_sheet_detail = ["STT", "Khoa", "Nhóm môn học", "Mã môn học", "Tên môn học", "Mã lớp", "Mã đơn vị", "Tên đơn vị", "Mã giảng viên", "Tên giảng viên", "Ngày bắt đầu TKB", "Đã soạn LMS"]
            sheet_detail.append(header_sheet_detail)
            for col_num, header in enumerate(header_sheet_detail, 1):
                  cell = sheet_detail.cell(row=1, column=col_num)
                  cell.font = header_font
                  cell.fill = header_fill
                  cell.alignment = header_alignment
                  cell.border = border_style

            
            if not os.path.exists(flle_track_no_lms):
                  wb_new_file_track_no_lms = Workbook()
                  ws_new_file_track_no_lms = wb_new_file_track_no_lms.active
                  ws_new_file_track_no_lms.append(header_sheet_detail[1:-1])
                  wb_new_file_track_no_lms.save(flle_track_no_lms)

            wb_file_track_no_lms = load_workbook(flle_track_no_lms)
            ws_file_track_no_lms = wb_file_track_no_lms.active
            row_in_sheet_no_lms = ws_file_track_no_lms.max_row + 1

            row_in_sheet_detail = 2
            for subject in data:
                  has_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid") # mặc định các dòng trong
                  if subject['has_lms'] != 'x':
                        has_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # nếu đã soạn LMS tô màu xanh lá

                        cell_no_lms_1 = ws_file_track_no_lms.cell(row=row_in_sheet_no_lms, column=1)
                        cell_no_lms_1.value = subject['department']

                        cell_no_lms_2 = ws_file_track_no_lms.cell(row=row_in_sheet_no_lms , column=2)
                        cell_no_lms_2.value = subject['group']

                        cell_no_lms_3 = ws_file_track_no_lms.cell(row=row_in_sheet_no_lms, column=3)
                        cell_no_lms_3.value = subject['id_subject']

                        cell_no_lms_4 = ws_file_track_no_lms.cell(row=row_in_sheet_no_lms, column=4)
                        cell_no_lms_4.value = subject['name_subject']

                        cell_no_lms_5 = ws_file_track_no_lms.cell(row=row_in_sheet_no_lms, column=5)
                        cell_no_lms_5.value = subject['id_class']

                        cell_no_lms_6 = ws_file_track_no_lms.cell(row=row_in_sheet_no_lms, column=6)
                        cell_no_lms_6.value = subject['id_unit']

                        cell_no_lms_7 = ws_file_track_no_lms.cell(row=row_in_sheet_no_lms, column=7)
                        cell_no_lms_7.value = subject['name_unit']

                        cell_no_lms_8 = ws_file_track_no_lms.cell(row=row_in_sheet_no_lms, column=8)
                        cell_no_lms_8.value = subject['id_teacher']

                        cell_no_lms_9 = ws_file_track_no_lms.cell(row=row_in_sheet_no_lms, column=9)
                        cell_no_lms_9.value = subject['name_teacher']

                        cell_no_lms_10 = ws_file_track_no_lms.cell(row=row_in_sheet_no_lms, column=10)
                        cell_no_lms_10.value = subject['from_day']

                        # wb_file_track_no_lms.save(flle_track_no_lms)
                        print(f"Đã thêm môn học {subject['id_subject']} - {subject['name_subject']} - {subject['id_teacher']} - {subject['name_teacher']}  vào file theo dõi không thực hiện LMS.")
                        row_in_sheet_no_lms += 1


                  cell = sheet_detail.cell(row=row_in_sheet_detail, column=1)
                  cell.value = row_in_sheet_detail - 1
                  cell.font = data_font
                  cell.alignment = data_alignment
                  cell.border = border_style
                  cell.fill = has_fill

                  col_in_sheet_derail = 2
                  for key, value in subject.items():
                        cell = sheet_detail.cell(row=row_in_sheet_detail, column=col_in_sheet_derail)
                        cell.value = value
                        cell.font = data_font
                        cell.alignment = data_alignment
                        cell.border = border_style
                        cell.fill = has_fill

                        col_in_sheet_derail += 1
                  row_in_sheet_detail += 1

            wb_file_track_no_lms.save(flle_track_no_lms)

            self.set_dimension_column(sheet_detail, header_sheet_detail)

            wb_report_lms.save(file_report_lms)

ob = ReportPerformLMS()
# ob.decor_report_perform_lms()
# ob.test()
# print(ob.test())
ob.report_perform_lms("2026-04-06", "2026-04-12")
# print(ob.get_department_of_subject())