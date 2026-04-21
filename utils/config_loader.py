# =============================================================================
# MODULE: utils/config_loader.py
# MỤC ĐÍCH: Đọc file cấu hình YAML và cung cấp interface lấy giá trị theo key.
#
# CÁCH DÙNG:
#   from utils.config_loader import ConfigLoader
#   config = ConfigLoader("path/to/config.yaml")
#   semester = config.get_attr("SEMESTER")
#
# LƯU Ý:
#   - File YAML phải dùng encoding UTF-8.
#   - Hiện tại class này chủ yếu được dùng để đọc config.yaml trong từng workflow,
#     nhưng đã được thay thế phần lớn bởi config/settings.py trực tiếp.
# =============================================================================

import yaml


class ConfigLoader:
      """Lớp đọc và quản lý cấu hình từ file YAML.

      Khi khởi tạo, toàn bộ file YAML được load vào bộ nhớ một lần.
      Sau đó dùng get_attr() để lấy từng giá trị theo tên key.
      """

      def __init__(self, path):
            """Đọc file YAML và lưu vào self.config.

            Tham số:
                path (str): đường dẫn đến file .yaml cần đọc.
                            Ví dụ: "automation/add_students_to_lms/config.yaml"
            """
            with open(path, encoding="utf-8") as f:
                  # yaml.safe_load đọc file YAML và trả về dict Python
                  # safe_load an toàn hơn load vì không cho phép thực thi code tùy ý
                  self.config = yaml.safe_load(f)

      def get_attr(self, key, default=None):
            """Lấy giá trị cấu hình theo tên key.

            Tham số:
                key     (str): tên key cần lấy, ví dụ "SEMESTER"
                default     : giá trị trả về nếu key không tồn tại (mặc định None)

            Trả về:
                Giá trị tương ứng với key trong file YAML,
                hoặc default nếu key không tồn tại.

            Ví dụ:
                config.get_attr("SEMESTER")           → "251"
                config.get_attr("TIMEOUT", default=30) → 30 (nếu TIMEOUT không có trong file)
            """
            return self.config.get(key, default)
