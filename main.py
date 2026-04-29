"""Lịch chạy tự động các workflow.

- Cập nhật sinh viên LMS: mỗi ngày lúc 17:00
- Báo cáo tình hình LMS:  mỗi thứ 2 lúc 08:00 (dữ liệu tuần trước)

Chạy:
    python main.py

Yêu cầu:
    pip install schedule
"""

import schedule
import time
import logging
from datetime import datetime
from automation.add_students_to_lms.run import UpdateStudents
from automation.report_perform_lms.run import ReportPerformLMS


def job_update_students():
    today = datetime.today()
    print(f"{today} Bắt đầu job cập nhật sinh viên LMS")
    try:
        ob = UpdateStudents()
        ob.update_student_lms()
        print(f"{today} Hoàn tất job cập nhật sinh viên LMS")
    except Exception as e:
        print(f"{today} Lỗi job cập nhật sinh viên LMS: {e}")


def job_report_lms():
    today = datetime.today()
    print(f"{today} Bắt đầu job báo cáo tình hình LMS")
    try:
        ob = ReportPerformLMS()
        from_day, to_day = ob.get_last_week_range()
        print(f"{today} Chạy báo cáo LMS từ {from_day} đến {to_day}")
        ob.report_perform_lms(from_day, to_day)
        print(f"{today} Hoàn tất job báo cáo tình hình LMS")
    except Exception as e:
        print(f"{today} Lỗi job báo cáo tình hình LMS: {e}")


# Cập nhật sinh viên mỗi ngày lúc 17:00
schedule.every().day.at("17:00").do(job_update_students)

# Báo cáo tình hình LMS mỗi thứ 2 lúc 08:00
schedule.every().monday.at("07:00").do(job_report_lms)


if __name__ == "__main__":
    today = datetime.today()
    print(f"{today} Khởi động lịch chạy tự động...")
    print(f"{today} Cập nhật sinh viên LMS: mỗi ngày lúc 17:00")
    print(f"{today} Báo cáo tình hình LMS: mỗi thứ 2 lúc 07:00")

    while True:
        schedule.run_pending()
        time.sleep(60)
