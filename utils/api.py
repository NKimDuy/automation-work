from config.settings import authorization
import requests
import base64
import bz2
import json

class APIHandler:
      def __init__(self):
            self.api_unit = "https://api.ou.edu.vn/api/v1/hdmdp"
            self.api_subject = "https://api.ou.edu.vn/api/v1/tkblopdp"
            self.headers = {
                  "Authorization": authorization,
                  "Content-Type": "application/json"
            }


      def get_subject_from_api(self, semester, unit_id):
            """Hàm lấy thông tin môn học từ API dựa trên học kỳ và mã đơn vị.

            Args:
                  semester (str): Học kỳ cần lấy thông tin, ví dụ "20251".
                  unit_id (str): Mã đơn vị cần lấy thông tin, ví dụ "SG".

            Returns:
                  list: Danh sách các môn học với thông tin chi tiết.
                  [
                  
                        [NhomTo, MaMH, TenMH, TUNGAYTKB, MaLop, MaDP, TenDP],
                        ....
                  ]
            """
            subjects = []
            params = {
                  "nhhk": semester,
                  "madp": unit_id
            }
            call_api = requests.get(self.api_subject, headers=self.headers, params=params)
            if call_api.status_code == 200:
                  data = call_api.json()
                  for subject in data.get("data", []):
                        subjects.append({
                              'NhomTo': subject['NhomTo'],
                              'MaMH': subject['MaMH'],
                              'TenMH': subject['TenMH'],
                              'TUNGAYTKB': subject['TUNGAYTKB'],
                              'MaLop': subject['MaLop'],
                              'MaDP': subject['MaDP'],
                              'TenDP': subject['TenDP']
                        })
                  return subjects
            else:
                  print(f"API môn học trả về lỗi: {call_api.status_code}")
                  return []


      def get_unit(self):
            """Hàm lấy thông tin đơn vị từ API.

            Returns:
                  list: Danh sách các đơn vị với thông tin chi tiết.
            """
            units = []
            call_api = requests.get(self.api_unit, headers=self.headers)
            if call_api.status_code == 200:
                  data = call_api.json()
                  for unit in data.get("data", []):
                        units.append([unit['MaDP'], unit['TenDP']])
                  return units
                  
            else:
                  print(f"API đơn vị trả về lỗi: {call_api.status_code}")
                  return []