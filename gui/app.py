# =============================================================================
# MODULE: gui/app.py
# MỤC ĐÍCH: Giao diện Tkinter cho 3 chức năng chính:
#   1. Báo cáo soạn thảo LMS hàng tuần (report_perform_lms)
#   2. Cập nhật sinh viên vào LMS (add_students_to_lms)
#   3. Chấm điểm LMS (score_lms) — có chọn model, chọn file Excel, validate cột
#
# CƠ CHẾ CHẠY NỀN:
#   Cả 3 workflow đều chạy khá lâu (Selenium mở trình duyệt, chờ tải trang...).
#   Nếu gọi trực tiếp trong hàm xử lý sự kiện nút bấm, cửa sổ sẽ bị "đơ" suốt
#   thời gian chạy. Nên mỗi workflow được chạy trong 1 threading.Thread riêng
#   (luồng nền), tách khỏi luồng chính của Tkinter (luồng chính chỉ lo vẽ giao diện).
#
#   Trong lúc luồng nền chạy, print() của nó được "hứng" qua QueueWriter
#   (xem gui/log_stream.py) rồi luồng chính định kỳ lấy ra hiển thị lên khung log.
#
# CHẠY TRỰC TIẾP:
#   python run_gui.py
# =============================================================================

import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox

from config import paths, settings
from automation.add_students_to_lms.run import UpdateStudents
from automation.report_perform_lms.run import ReportPerformLMS
from automation.score_lms.run import ScoreLMS
from gui.log_stream import QueueWriter


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automatic Work - LMS")
        self.geometry("820x620")
        self.minsize(700, 500)

        # hàng đợi nhận log từ thread nền, luồng chính sẽ đọc định kỳ (_poll_log_queue)
        self.log_queue = queue.Queue()
        # chỉ cho chạy 1 workflow tại 1 thời điểm (cả 3 đều dùng chung Selenium/Chrome)
        self.is_running = False

        # đường dẫn model/file cho "Chấm điểm LMS" — mặc định lấy theo settings,
        # người dùng có thể đổi qua nút "Chọn..."
        self.score_model_path = tk.StringVar(
            value=paths.model_file_dir(settings.model_filename)
        )
        self.score_excel_path = tk.StringVar(
            value=paths.score_lms_input_file_dir(settings.semester_score_lms, settings.score_lms_input_filename)
        )

        self._build_widgets()
        self.after(100, self._poll_log_queue)

    # -------------------------------------------------------------------
    # DỰNG GIAO DIỆN
    # -------------------------------------------------------------------
    def _build_widgets(self):
        pad = {"padx": 10, "pady": 6}

        # ---- Khối 1: Báo cáo soạn thảo LMS hàng tuần ----
        frame_report = ttk.LabelFrame(self, text="Báo cáo soạn thảo LMS hàng tuần")
        frame_report.pack(fill="x", **pad)
        self.btn_report = ttk.Button(
            frame_report, text="Chạy", command=self._on_click_run_report
        )
        self.btn_report.pack(anchor="w", padx=10, pady=8)

        # ---- Khối 2: Cập nhật sinh viên vào LMS ----
        frame_students = ttk.LabelFrame(self, text="Cập nhật sinh viên vào LMS")
        frame_students.pack(fill="x", **pad)
        self.btn_students = ttk.Button(
            frame_students, text="Chạy", command=self._on_click_run_add_students
        )
        self.btn_students.pack(anchor="w", padx=10, pady=8)

        # ---- Khối 3: Chấm điểm LMS ----
        frame_score = ttk.LabelFrame(self, text="Chấm điểm LMS")
        frame_score.pack(fill="x", **pad)

        row_model = ttk.Frame(frame_score)
        row_model.pack(fill="x", padx=10, pady=4)
        ttk.Label(row_model, text="Model:", width=10).pack(side="left")
        ttk.Entry(row_model, textvariable=self.score_model_path, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        ttk.Button(row_model, text="Chọn...", command=self._choose_model_file).pack(side="left")

        row_excel = ttk.Frame(frame_score)
        row_excel.pack(fill="x", padx=10, pady=4)
        ttk.Label(row_excel, text="File môn học:", width=10).pack(side="left")
        ttk.Entry(row_excel, textvariable=self.score_excel_path, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        ttk.Button(row_excel, text="Chọn...", command=self._choose_excel_file).pack(side="left")

        self.btn_score = ttk.Button(
            frame_score, text="Chạy", command=self._on_click_run_score
        )
        self.btn_score.pack(anchor="w", padx=10, pady=8)

        # ---- Khung log dùng chung cho cả 3 khối ----
        ttk.Label(self, text="Log:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(self, state="disabled", wrap="word", height=18)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        self.status_label = ttk.Label(self, text="Sẵn sàng")
        self.status_label.pack(anchor="w", padx=10, pady=(0, 8))

    # -------------------------------------------------------------------
    # CHỌN FILE (chỉ dùng cho khối Chấm điểm LMS)
    # -------------------------------------------------------------------
    def _choose_model_file(self):
        path = filedialog.askopenfilename(
            title="Chọn file model",
            initialdir=paths.MODEL_DIR,
            filetypes=[("Model file", "*.pkl"), ("Tất cả", "*.*")],
        )
        if path:
            self.score_model_path.set(path)

    def _choose_excel_file(self):
        current_dir = os.path.dirname(self.score_excel_path.get())
        path = filedialog.askopenfilename(
            title="Chọn file Excel danh sách môn học cần đánh giá",
            initialdir=current_dir if os.path.isdir(current_dir) else paths.BASE_DIR,
            filetypes=[("Excel file", "*.xlsx"), ("Tất cả", "*.*")],
        )
        if path:
            self.score_excel_path.set(path)

    # -------------------------------------------------------------------
    # XỬ LÝ SỰ KIỆN NÚT "CHẠY" — kiểm tra nhanh trước (nếu cần) rồi mới chạy nền
    # -------------------------------------------------------------------
    def _on_click_run_report(self):
        self._start_job(self._run_report)

    def _on_click_run_add_students(self):
        self._start_job(self._run_add_students)

    def _on_click_run_score(self):
        excel_path = self.score_excel_path.get()

        # kiểm tra cột NGAY khi bấm "Chạy" — việc này nhẹ (chỉ đọc dòng tiêu đề),
        # nên gọi thẳng trong luồng chính, không cần chờ chạy nền mới biết sai
        try:
            _, missing_columns = ScoreLMS.validate_columns(excel_path)
        except Exception as e:
            messagebox.showerror("Không đọc được file", f"Không mở được file Excel đã chọn:\n{e}")
            return

        if missing_columns:
            messagebox.showerror(
                "Thiếu cột",
                "File Excel đang chọn thiếu các cột bắt buộc:\n- " + "\n- ".join(missing_columns)
            )
            return

        self._start_job(self._run_score_lms)

    # -------------------------------------------------------------------
    # NỘI DUNG THỰC SỰ CỦA TỪNG WORKFLOW — chạy bên trong thread nền
    # -------------------------------------------------------------------
    def _run_report(self):
        ReportPerformLMS().report_perform_lms()

    def _run_add_students(self):
        UpdateStudents().update_student_lms()

    def _run_score_lms(self):
        ob = ScoreLMS(
            model_file=self.score_model_path.get(),
            excel_file=self.score_excel_path.get(),
        )
        ob.score_lms()

    # -------------------------------------------------------------------
    # CƠ CHẾ CHẠY NỀN + HIỂN THỊ LOG
    # -------------------------------------------------------------------
    def _all_buttons(self):
        return [self.btn_report, self.btn_students, self.btn_score]

    def _start_job(self, target):
        """Khóa các nút lại, chạy `target` (1 trong 3 hàm _run_xxx) trong thread nền."""
        if self.is_running:
            return
        self.is_running = True
        for btn in self._all_buttons():
            btn.state(["disabled"])
        self.status_label.config(text="Đang chạy...")
        threading.Thread(target=self._run_in_thread, args=(target,), daemon=True).start()

    def _run_in_thread(self, target):
        # chuyển hướng print() sang hàng đợi trong lúc chạy — xem gui/log_stream.py
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.log_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            target()
            self.log_queue.put("\n[Hoàn tất]\n")
        except Exception as e:
            self.log_queue.put(f"\n[Lỗi] {e}\n")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            # tín hiệu báo cho luồng chính biết đã chạy xong, để mở khóa lại các nút
            self.log_queue.put("__DONE__")

    def _poll_log_queue(self):
        """Chạy lặp lại mỗi 100ms trên luồng chính — lấy log từ hàng đợi, đổ lên khung hiển thị."""
        try:
            while True:
                line = self.log_queue.get_nowait()
                if line == "__DONE__":
                    self.is_running = False
                    for btn in self._all_buttons():
                        btn.state(["!disabled"])
                    self.status_label.config(text="Sẵn sàng")
                    continue
                self.log_text.configure(state="normal")
                self.log_text.insert("end", line)
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._poll_log_queue)
