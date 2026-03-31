# db/

## Mục đích
Chứa module kết nối database và hướng dẫn schema dùng cho bảng `students_lms`.

## File chính
- `connect_mysql.py`: hàm `get_connection()` trả connection MySQL.

## Schema ví dụ (MySQL)
```sql
CREATE TABLE automatic_work.students_lms (
    id INT AUTO_INCREMENT PRIMARY KEY,
    lms_id VARCHAR(50),
    group_id VARCHAR(20),
    subject_id VARCHAR(50),
    subject_name VARCHAR(255),
    teacher_id VARCHAR(50),
    teacher_name VARCHAR(255),
    date_update DATETIME,
    number_student_update INT,
    semester VARCHAR(20)
);
```

## Ghi chú
- Thêm chỉ mục (`INDEX`) lên `semester` và `lms_id` nếu query thường xuyên.
