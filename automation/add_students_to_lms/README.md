# automation/add_students_to_lms

## Mục đích
Tập hợp module và script dùng để đồng bộ dữ liệu sinh viên giữa LSA (thông tin môn học) và LMS (Moodle).

## File chính
- `run.py`: entry point với class `UpdateStudents`.
- `config.yaml`: thông số cấu hình riêng.

## Luồng thực hiện
1. `InitSelenium.process_detail_subject`: lấy danh sách môn học từ LSA.
2. `UpdateStudents.update_student_lms`: duyệt danh sách môn, vào trang cập nhật, lấy số sinh viên, insert vào DB.
3. `UpdateStudents.find_subject_lms`: chức năng phụ test kiểm tra tìm môn học.

## Cấu hình cần có
- `SEMESTER` trong `config.yaml`.
- `config/settings.py` chứa `user_lms`, `password_lms`, `captcha`.
- DB: `config.py` trong `db/connect_mysql` cung cấp thông tin kết nối.

## Ghi chú tốt
- Không để hardcode account cá nhân trong code, nên dùng biến môi trường hoặc .env khi nâng cấp.
- Node `add_students_to_lms.log` ghi chi tiết thời gian chạy và lỗi.
