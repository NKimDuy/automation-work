# =============================================================================
# MODULE: db/connect_mysql.py
# MỤC ĐÍCH: Cung cấp kết nối đến cơ sở dữ liệu MySQL.
#
# CÁCH DÙNG:
#   from db.connect_mysql import get_connection
#   con = get_connection()
#   try:
#       # dùng con để thực hiện query
#   finally:
#       con.close()  # LUÔN đóng kết nối sau khi dùng xong
#
# LƯU Ý:
#   - Hàm này trả về connection, người gọi tự chịu trách nhiệm đóng kết nối.
#   - Thông tin kết nối (host, user, password, database) lấy từ config/settings.py (đã gitignore).
#   - Bảng chính đang dùng: automatic_work.students_lms
# =============================================================================

import mysql.connector
from config.settings import host, user_db, password_db, database


def get_connection():
      """Tạo và trả về một kết nối MySQL mới.

      Thông tin kết nối được lấy từ config/settings.py:
        - host       : địa chỉ server MySQL
        - user_db    : tên đăng nhập database
        - password_db: mật khẩu database
        - database   : tên database (automatic_work)

      Trả về:
          mysql.connector.connection.MySQLConnection: đối tượng kết nối MySQL.

      Lưu ý:
          Người gọi hàm này PHẢI tự đóng kết nối bằng con.close()
          sau khi dùng xong, tốt nhất dùng trong khối try/finally.
      """
      conn = mysql.connector.connect(
            host=host,
            user=user_db,
            password=password_db,
            database=database
      )
      return conn
