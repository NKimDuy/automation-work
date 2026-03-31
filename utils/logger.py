import logging
import os

# Thiết lập logger (ghi log vào file) với định dạng: thời gian - mức độ log - thông điệp log
def setup_logger(log_dir):
      logging.basicConfig(
            filename=log_dir,
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
            encoding="utf-8"
      )
      return logging