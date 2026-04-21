# =============================================================================
# MODULE: utils/api.py
# MỤC ĐÍCH: Gọi REST API của hệ thống api.ou.edu.vn để lấy danh sách đơn vị
#           và danh sách môn học theo học kỳ.
#
# API SỬ DỤNG:
#   - GET https://api.ou.edu.vn/api/v1/hdmdp       → danh sách đơn vị (khoa/trung tâm)
#   - GET https://api.ou.edu.vn/api/v1/tkblopdp    → danh sách môn học theo học kỳ + đơn vị
#
# CÁCH DÙNG:
#   from utils.api import APIHandler
#   api = APIHandler()
#   list_unit    = api.get_unit()
#   list_subject = api.get_subject_from_api("20251", "SG")
#
# XÁC THỰC:
#   Dùng Bearer Token, lấy từ config/settings.py (biến authorization).
# =============================================================================

from config.settings import authorization
import requests


class APIHandler:
      """Lớp xử lý các lời gọi API đến hệ thống api.ou.edu.vn.

      Tất cả request đều dùng chung headers Authorization (Bearer Token)
      được cấu hình một lần trong __init__.
      """

      def __init__(self):
            """Khởi tạo URL endpoint và headers xác thực dùng chung cho mọi request."""
            # Endpoint lấy danh sách đơn vị (khoa, trung tâm, ...)
            self.api_unit = "https://api.ou.edu.vn/api/v1/hdmdp"
            # Endpoint lấy danh sách môn học theo học kỳ và đơn vị
            self.api_subject = "https://api.ou.edu.vn/api/v1/tkblopdp"
            # Header dùng Bearer Token để xác thực, lấy từ config/settings.py
            self.headers = {
                  "Authorization": authorization,
                  "Content-Type": "application/json"
            }


      def get_subject_from_api(self, semester, unit_id):
            """Lấy danh sách môn học theo học kỳ và mã đơn vị từ API.

            Gọi API với params nhhk (học kỳ) và madp (mã đơn vị),
            rồi trích xuất các trường cần thiết từ response.

            Tham số:
                semester (str): mã học kỳ dạng "20251"
                                (= "20" + semester config, ví dụ semester="251" → "20251")
                unit_id  (str): mã đơn vị, ví dụ "SG", "TX", "LA"

            Trả về:
                list[dict]: danh sách môn học, mỗi phần tử là dict gồm:
                    - NhomTo    : nhóm môn học, ví dụ "SG001"
                    - MaMH      : mã môn học, ví dụ "ACCO4331"
                    - TenMH     : tên môn học, ví dụ "Kế toán tài chính"
                    - TUNGAYTKB : ngày bắt đầu thời khóa biểu, dạng "YYYY-MM-DD" hoặc None
                    - MaLop     : mã lớp, ví dụ "TM123456"
                    - MaDP      : mã đơn vị, ví dụ "SG"
                    - TenDP     : tên đơn vị, ví dụ "Sài Gòn"
                Trả về [] nếu API lỗi.
            """
            subjects = []
            # Tham số truy vấn gửi lên API
            params = {
                  "nhhk": semester,   # học kỳ, ví dụ "20251"
                  "madp": unit_id     # mã đơn vị, ví dụ "SG"
            }
            call_api = requests.get(self.api_subject, headers=self.headers, params=params)

            if call_api.status_code == 200:
                  data = call_api.json()
                  # data["data"] là list các môn học trả về từ API
                  for subject in data.get("data", []):
                        # Chỉ lấy các trường cần thiết, bỏ qua các trường thừa từ API
                        subjects.append({
                              'NhomTo'    : subject['NhomTo'],
                              'MaMH'      : subject['MaMH'],
                              'TenMH'     : subject['TenMH'],
                              'TUNGAYTKB' : subject['TUNGAYTKB'],  # có thể là None nếu chưa xếp TKB
                              'MaLop'     : subject['MaLop'],
                              'MaDP'      : subject['MaDP'],
                              'TenDP'     : subject['TenDP']
                        })
                  return subjects
            else:
                  print(f"API môn học trả về lỗi: {call_api.status_code}")
                  return []


      def get_unit(self):
            """Lấy danh sách tất cả đơn vị (khoa/trung tâm) từ API.

            Không cần tham số — API trả về toàn bộ đơn vị trong hệ thống.
            Kết quả này được dùng để lặp qua từng đơn vị và gọi get_subject_from_api().

            Trả về:
                list[list]: danh sách đơn vị, mỗi phần tử là [MaDP, TenDP].
                    - MaDP  : mã đơn vị, ví dụ "SG"
                    - TenDP : tên đơn vị, ví dụ "Sài Gòn"
                Trả về [] nếu API lỗi.

            Ví dụ kết quả:
                [["SG", "Sài Gòn"], ["TX", "Tây Nam Bộ"], ["LA", "Long An"], ...]
            """
            units = []
            call_api = requests.get(self.api_unit, headers=self.headers)

            if call_api.status_code == 200:
                  data = call_api.json()
                  # Chỉ lấy MaDP và TenDP, bỏ qua các trường thừa
                  for unit in data.get("data", []):
                        units.append([unit['MaDP'], unit['TenDP']])
                  return units
            else:
                  print(f"API đơn vị trả về lỗi: {call_api.status_code}")
                  return []
