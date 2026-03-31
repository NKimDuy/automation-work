# automation-work

## Mục đích
Dự án tự động hóa cập nhật danh sách sinh viên từ hệ thống LSA sang LMS, với Selenium và MySQL.

## Cấu trúc chính
- `automation/`: các module tự động hóa theo từng workflow.
- `config/`: cấu hình chung, đường dẫn, tham số hệ thống.
- `data/`: input/output files như CSV/Excel.
- `db/`: module kết nối database.
- `modules/`: helper modules (Selenium, v.v.).
- `utils/`: loader config, logging, v.v.

## Cách chạy
1. `pip install -r requirements.txt`
2. Tạo config trong `config/settings.py` (user_lms, password_lms, captcha, ...)
3. Chạy `python -m automation.add_students_to_lms.run`

## Tài liệu thêm
- `automation/add_students_to_lms/README.md`
- `config/README.md`
- `data/input/README.md`
- `data/output/README.md`
- `db/README.md`
