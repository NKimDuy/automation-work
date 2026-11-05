# =============================================================================
# MODULE: gui/app.py
# MỤC ĐÍCH: Giao diện Tkinter cho 3 chức năng chính + 1 tab cài đặt:
#   1. Báo cáo soạn thảo LMS hàng tuần (report_perform_lms)
#   2. Cập nhật sinh viên vào LMS (add_students_to_lms)
#   3. Chấm điểm LMS (score_lms) — có chọn model, chọn file Excel, validate cột
#   4. Cài đặt — sửa học kỳ (semester_general, semester_score_lms) ngay trên GUI,
#      lưu vào runtime_settings.json, không cần mở code / build lại app
#
# LAYOUT: dùng ttk.Notebook (tab) để mỗi chức năng nằm gọn trong 1 tab riêng
#   thay vì show hết cùng lúc. Khung log dùng chung, nằm dưới các tab (vì log
#   là của "workflow đang chạy", không thuộc riêng tab nào).
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
# THỨ TỰ CÁC HÀM TRONG FILE NÀY: xếp theo đúng chuỗi gọi thực tế, đọc từ trên
# xuống là theo được luồng chạy:
#   __init__ → _build_widgets → _build_tab_xxx → _choose_xxx_file /
#   _on_click_save_settings → _on_click_run_xxx → _start_job → _run_in_thread →
#   _run_xxx → _all_buttons → _poll_log_queue
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

from config import paths, settings, runtime_settings
from automation.add_students_to_lms.run import UpdateStudents
from automation.report_perform_lms.run import ReportPerformLMS
from automation.score_lms.run import ScoreLMS
from gui.log_stream import QueueWriter


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automatic Work - LMS")
        self.geometry("780x560")
        self.minsize(680, 460)

        # hàng đợi nhận log từ thread nền, luồng chính sẽ đọc định kỳ (_poll_log_queue)
        self.log_queue = queue.Queue()
        # chỉ cho chạy 1 workflow tại 1 thời điểm (cả 3 đều dùng chung Selenium/Chrome)
        self.is_running = False

        current = runtime_settings.load()

        # đường dẫn model/file cho tab "Chấm điểm LMS" — mặc định lấy theo cài đặt hiện tại,
        # người dùng có thể đổi qua nút "Chọn..."
        self.score_model_path = tk.StringVar(
            value=paths.model_file_dir(settings.model_filename)
        )
        self.score_excel_path = tk.StringVar(
            value=paths.score_lms_input_file_dir(current["semester_score_lms"], settings.score_lms_input_filename)
        )

        # giá trị hiển thị trên tab "Cài đặt"
        self.semester_general_var = tk.StringVar(value=current["semester_general"])
        self.semester_score_lms_var = tk.StringVar(value=current["semester_score_lms"])

        self._build_widgets()
        self.after(100, self._poll_log_queue)

    # -------------------------------------------------------------------
    # DỰNG GIAO DIỆN — tạo notebook (tab) + khung log dùng chung bên dưới
    # -------------------------------------------------------------------
    def _build_widgets(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="x", padx=10, pady=10)

        tab_report = ttk.Frame(notebook)
        tab_students = ttk.Frame(notebook)
        tab_score = ttk.Frame(notebook)
        tab_settings = ttk.Frame(notebook)

        notebook.add(tab_report, text="Báo cáo LMS")
        notebook.add(tab_students, text="Cập nhật sinh viên")
        notebook.add(tab_score, text="Chấm điểm LMS")
        notebook.add(tab_settings, text="Cài đặt")

        self._build_tab_report(tab_report)
        self._build_tab_students(tab_students)
        self._build_tab_score(tab_score)
        self._build_tab_settings(tab_settings)

        # ---- Khung log dùng chung cho cả 3 chức năng ----
        ttk.Label(self, text="Log:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(self, state="disabled", wrap="word", height=16)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        self.status_label = ttk.Label(self, text="Sẵn sàng")
        self.status_label.pack(anchor="w", padx=10, pady=(0, 8))

    def _build_tab_report(self, tab):
        ttk.Label(
            tab, text="Tạo báo cáo tình hình soạn thảo LMS của tuần trước.",
            wraplength=620, justify="left"
        ).pack(anchor="w", padx=10, pady=10)
        self.btn_report = ttk.Button(tab, text="Chạy", command=self._on_click_run_report)
        self.btn_report.pack(anchor="w", padx=10, pady=(0, 10))

    def _build_tab_students(self, tab):
        ttk.Label(
            tab, text="Cập nhật danh sách sinh viên từ LSA vào LMS.",
            wraplength=620, justify="left"
        ).pack(anchor="w", padx=10, pady=10)
        self.btn_students = ttk.Button(tab, text="Chạy", command=self._on_click_run_add_students)
        self.btn_students.pack(anchor="w", padx=10, pady=(0, 10))

    def _build_tab_score(self, tab):
        row_model = ttk.Frame(tab)
        row_model.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(row_model, text="Model:", width=10).pack(side="left")
        ttk.Entry(row_model, textvariable=self.score_model_path, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        ttk.Button(row_model, text="Chọn...", command=self._choose_model_file).pack(side="left")

        row_excel = ttk.Frame(tab)
        row_excel.pack(fill="x", padx=10, pady=4)
        ttk.Label(row_excel, text="File môn học:", width=10).pack(side="left")
        ttk.Entry(row_excel, textvariable=self.score_excel_path, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        ttk.Button(row_excel, text="Chọn...", command=self._choose_excel_file).pack(side="left")

        self.btn_score = ttk.Button(tab, text="Chạy", command=self._on_click_run_score)
        self.btn_score.pack(anchor="w", padx=10, pady=10)

    def _build_tab_settings(self, tab):
        ttk.Label(
            tab,
            text="Sửa học kỳ ở đây — áp dụng ngay cho lần bấm \"Chạy\" tiếp theo, "
                 "không cần sửa code hay khởi động lại app.",
            wraplength=620, justify="left"
        ).pack(anchor="w", padx=10, pady=(10, 10))

        row_general = ttk.Frame(tab)
        row_general.pack(fill="x", padx=10, pady=4)
        ttk.Label(row_general, text="Học kỳ chung:", width=24).pack(side="left")
        ttk.Entry(row_general, textvariable=self.semester_general_var, width=15).pack(side="left")

        row_score = ttk.Frame(tab)
        row_score.pack(fill="x", padx=10, pady=4)
        ttk.Label(row_score, text="Học kỳ chấm điểm LMS:", width=24).pack(side="left")
        ttk.Entry(row_score, textvariable=self.semester_score_lms_var, width=15).pack(side="left")

        ttk.Button(tab, text="Lưu", command=self._on_click_save_settings).pack(anchor="w", padx=10, pady=12)

    # -------------------------------------------------------------------
    # Được nút "Chọn..." trong _build_tab_score() gọi tới
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
    # Được nút "Lưu" trong _build_tab_settings() gọi tới
    # -------------------------------------------------------------------
    def _on_click_save_settings(self):
        semester_general = self.semester_general_var.get().strip()
        semester_score_lms = self.semester_score_lms_var.get().strip()

        if not semester_general or not semester_score_lms:
            messagebox.showerror("Thiếu thông tin", "Học kỳ không được để trống.")
            return

        runtime_settings.save(
            semester_general=semester_general,
            semester_score_lms=semester_score_lms,
        )
        messagebox.showinfo("Đã lưu", "Đã lưu học kỳ mới — áp dụng từ lần chạy tiếp theo.")

    # -------------------------------------------------------------------
    # Được nút "Chạy" trong từng tab gọi tới — kiểm tra nhanh (nếu cần)
    # rồi giao việc chạy thật cho _start_job()
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
    # Được 3 hàm _on_click_run_xxx() ở trên gọi tới — lo việc xóa log cũ,
    # khóa nút, tạo thread nền
    # -------------------------------------------------------------------
    def _start_job(self, target):
        """Khóa các nút lại, xóa log cũ, chạy `target` (1 trong 3 hàm _run_xxx bên dưới) trong thread nền."""
        if self.is_running:
            return
        self.is_running = True
        for btn in self._all_buttons():
            btn.state(["disabled"])

        # xóa log của lần chạy trước — mỗi lần bấm "Chạy" là 1 log mới, không nối tiếp log cũ
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        self.status_label.config(text="Đang chạy...")
        threading.Thread(target=self._run_in_thread, args=(target,), daemon=True).start()

    # -------------------------------------------------------------------
    # Được _start_job() giao cho thread nền chạy — lo việc chuyển hướng print()
    # -------------------------------------------------------------------
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

    # -------------------------------------------------------------------
    # Nội dung thật sự của từng workflow — được _run_in_thread() gọi qua tham số `target`
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
    # Hàm phụ nhỏ, dùng chung ở cả _start_job() và _poll_log_queue()
    # -------------------------------------------------------------------
    def _all_buttons(self):
        return [self.btn_report, self.btn_students, self.btn_score]

    # -------------------------------------------------------------------
    # Được __init__() hẹn giờ chạy lặp lại qua self.after(...)
    # -------------------------------------------------------------------
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
