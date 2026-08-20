# =============================================================================
# MODULE: gui/log_stream.py
# MỤC ĐÍCH: Cầu nối giữa luồng (thread) chạy nền và giao diện Tkinter (luồng chính).
#
# VÌ SAO CẦN FILE NÀY:
#   Các hàm workflow hiện tại (update_student_lms, report_perform_lms, score_lms)
#   đều báo tiến trình bằng print(). Khi GUI chạy các hàm này trong 1 thread nền
#   (để không làm đơ cửa sổ), Tkinter không cho phép thread nền tự ý cập nhật
#   widget trực tiếp — mọi cập nhật giao diện phải do luồng chính (main thread)
#   thực hiện.
#
#   QueueWriter đóng vai trò "giả làm sys.stdout": mọi lệnh print() trong lúc
#   thread nền chạy sẽ được gom vào 1 hàng đợi (queue.Queue) thay vì in ra
#   console. Luồng chính (gui/app.py) định kỳ lấy dữ liệu từ hàng đợi này ra
#   và mới là bên thật sự ghi lên khung log.
# =============================================================================

import queue


class QueueWriter:
    """File-like object dùng thay cho sys.stdout/sys.stderr trong lúc chạy nền.

    "File-like object" nghĩa là 1 đối tượng có đủ 2 hàm write()/flush() để
    Python coi nó như 1 file thật — vì print() nội bộ gọi sys.stdout.write(),
    nên chỉ cần định nghĩa đúng 2 hàm này là "đánh lừa" được print().
    """

    def __init__(self, log_queue: queue.Queue):
        self.log_queue = log_queue

    def write(self, text):
        if text:
            self.log_queue.put(text)

    def flush(self):
        # print() thỉnh thoảng gọi flush() — không cần làm gì vì Queue không cần "đẩy" dữ liệu đi đâu cả
        pass
