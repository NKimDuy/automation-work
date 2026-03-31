import yaml

# Hàm ConfigLoader để đọc file config.yaml, lấy các giá trị cấu hình cần thiết cho chương trình, 
# ví dụ như SEMESTER để lấy link truy cập vào LMS tương ứng với học kỳ hiện
class ConfigLoader:
      def __init__(self, path):
            with open(path, encoding="utf-8") as f:
                  self.config = yaml.safe_load(f)
      
      def get_attr(self, key, default=None):
            return self.config.get(key, default)