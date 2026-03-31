# data/input/

## Mục đích
Chứa file dữ liệu nguồn cần xử lý (CSV/XLSX) như danh sách sinh viên, mã môn học, ...

## Định dạng mẫu đề xuất
- `student_id`: mã sinh viên
- `subject_id`: mã môn
- `group_id`: mã nhóm
- `semester`: học kỳ

## Ghi chú
- Kiểm tra encoding UTF-8 để tránh lỗi vỡ tiếng việt.
- Đặt tên file rõ ràng `input_students_{yyyyMMdd}.csv`.
