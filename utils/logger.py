# =============================================================================
# MODULE: utils/logger.py
# MỤC ĐÍCH: Tạo và cấu hình logger ghi log ra file.
#
# CÁCH DÙNG:
#   from utils.logger import setup_logger
#   logger = setup_logger("/path/to/file.log")
#   logger.info("Thông báo bình thường")
#   logger.error("Thông báo lỗi")
#
# LƯU Ý:
#   - Mỗi lần gọi setup_logger với cùng một file_path, nội dung log CŨ sẽ bị
#     GHI ĐÈ (do dùng basicConfig mặc định với filemode='w').
#   - Nếu muốn ghi tiếp (append) thay vì ghi đè, cần thêm filemode='a'.
# =============================================================================

import logging
import os


def setup_logger(log_dir):
      """Khởi tạo và trả về module logging đã được cấu hình ghi vào file.

      Định dạng log: "YYYY-MM-DD HH:MM:SS,mmm - LEVEL - Nội dung thông điệp"
      Ví dụ: "2026-04-14 17:00:05,123 - INFO - Bắt đầu chạy chương trình"

      Mức log mặc định là INFO, tức là sẽ ghi tất cả log từ INFO trở lên:
        - INFO    : thông tin quá trình chạy bình thường
        - WARNING : cảnh báo, chương trình vẫn chạy được
        - ERROR   : lỗi xảy ra tại một bước cụ thể
        - CRITICAL: lỗi nghiêm trọng (hiếm dùng)

      Tham số:
          log_dir (str): đường dẫn tuyệt đối đến file log cần ghi.
                         Ví dụ: "automation/add_students_to_lms/add_students_to_lms.log"

      Trả về:
          module logging: dùng trực tiếp logging.info(), logging.error(), v.v.
      """
      logging.basicConfig(
            filename=log_dir,      # ghi log vào file tại đường dẫn này
            level=logging.INFO,    # ghi tất cả log từ mức INFO trở lên
            format="%(asctime)s - %(levelname)s - %(message)s",  # định dạng mỗi dòng log
            encoding="utf-8"       # dùng UTF-8 để hỗ trợ tiếng Việt trong log
      )
      return logging
