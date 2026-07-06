# -*- coding: utf-8 -*-
"""
Công cụ lọc & in hàng loạt file PDF / Word / Excel theo mã công ty
====================================================================

YÊU CẦU CÀI ĐẶT:
    pip install pywin32

TÙY CHỌN (khuyến nghị để in PDF im lặng, nhanh, không mở app):
    Cài SumatraPDF (miễn phí): https://www.sumatrapdfreader.org/download-free-pdf-viewer
    Nếu không có, chương trình sẽ dùng lệnh "print" mặc định của Windows
    (thường sẽ mở Adobe Reader thoáng qua rồi tự in).

CÁCH DÙNG:
    1. Chạy: python mass_print_tool.py
    2. Bấm "Chọn thư mục" -> chọn folder chứa ~300 file
    3. Nhập mã công ty vào ô tìm kiếm (ví dụ: XEM001), có thể nhập nhiều mã
       cách nhau bằng dấu phẩy (ví dụ: XEM001, XEM002)
    4. Chọn loại file muốn lọc (PDF / DOC / XLS / Tất cả) bằng nút bấm
    5. Danh sách file khớp sẽ hiện ra, click vào dòng để tick/bỏ tick chọn in
       (hoặc bấm "Chọn tất cả"). Click vào tiêu đề cột để sắp xếp.
       Double-click vào 1 dòng để mở file xem trước (không in).
    6. Chọn máy in ở dropdown "Máy in" (mặc định dùng máy in mặc định hệ thống)
    7. Bấm "IN HÀNG LOẠT" để in toàn bộ file đã chọn. Có thể bấm "Dừng in"
       để dừng giữa batch. Log chi tiết được lưu trong thư mục "print_logs"
       cạnh file script này.
"""

import os
import time
import threading
import subprocess
import datetime
import json
import re
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

if HAS_DND:
    class BaseApp(TkinterDnD.Tk):
        pass
else:
    class BaseApp(tk.Tk):
        pass

# ----------------------------------------------------------------------
# Cấu hình loại file
# ----------------------------------------------------------------------
FILE_TYPE_GROUPS = {
    "PDF": [".pdf"],
    "DOC": [".doc", ".docx"],
    "XLS": [".xls", ".xlsx"],
}
ALL_EXTENSIONS = [ext for exts in FILE_TYPE_GROUPS.values() for ext in exts]

# Các vị trí cài đặt SumatraPDF phổ biến trên Windows
SUMATRA_CANDIDATES = [
    r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
    r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    os.path.expanduser(r"~\AppData\Local\SumatraPDF\SumatraPDF.exe"),
]


def find_sumatra():
    for path in SUMATRA_CANDIDATES:
        if os.path.isfile(path):
            return path
    return None


def list_printers():
    """Trả về (danh sách tên máy in, tên máy in mặc định)."""
    try:
        import win32print
        printers = [
            p[2] for p in win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            )
        ]
        default = win32print.GetDefaultPrinter()
        return printers, default
    except Exception:
        return [], None


def natural_sort_key(text):
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', text)]


# ----------------------------------------------------------------------
# Logic in file (Windows)
# ----------------------------------------------------------------------
class Printer:
    """Quản lý việc in Word / Excel / PDF, tái sử dụng 1 instance Word/Excel
    cho toàn bộ batch để chạy nhanh hơn."""

    def __init__(self, log_callback, printer_name=None):
        self.log = log_callback
        self._word_app = None
        self._excel_app = None
        self.sumatra_path = find_sumatra()
        # None hoặc chuỗi rỗng nghĩa là dùng máy in mặc định của Windows
        self.printer_name = printer_name or None
        self._original_default_printer = None

    def _get_word(self):
        if self._word_app is None:
            import win32com.client
            self._word_app = win32com.client.DispatchEx("Word.Application")
            self._word_app.Visible = False
            self._word_app.DisplayAlerts = 0  # wdAlertsNone
            try:
                # msoAutomationSecurityForceDisable = 3
                # Ép tắt cảnh báo macro/security để không bị treo khi in ngầm
                # (Visible=False nên sẽ không ai bấm được popup này)
                self._word_app.AutomationSecurity = 3
            except Exception:
                pass
        return self._word_app

    def _get_excel(self):
        if self._excel_app is None:
            import win32com.client
            self._excel_app = win32com.client.DispatchEx("Excel.Application")
            self._excel_app.Visible = False
            self._excel_app.DisplayAlerts = False
        return self._excel_app

    def print_word(self, filepath):
        word = self._get_word()
        if self.printer_name:
            try:
                word.ActivePrinter = self.printer_name
            except Exception as e:
                self.log(f"[CẢNH BÁO] Không đặt được máy in cho Word: {e}")
        doc = word.Documents.Open(filepath, ReadOnly=True)
        try:
            doc.PrintOut()
            time.sleep(1)  # chờ spool trước khi đóng
        finally:
            doc.Close(False)

    def print_excel(self, filepath):
        excel = self._get_excel()
        if self.printer_name:
            try:
                excel.ActivePrinter = self.printer_name
            except Exception as e:
                self.log(f"[CẢNH BÁO] Không đặt được máy in cho Excel: {e}")
        wb = excel.Workbooks.Open(filepath, ReadOnly=True)
        try:
            wb.PrintOut()
            time.sleep(1)
        finally:
            wb.Close(False)

    def print_pdf(self, filepath):
        if self.sumatra_path:
            if self.printer_name:
                cmd = [self.sumatra_path, "-print-to", self.printer_name, "-silent", filepath]
            else:
                cmd = [self.sumatra_path, "-print-to-default", "-silent", filepath]
            subprocess.run(cmd, timeout=60)
        else:
            # Không có SumatraPDF: dùng verb "print" mặc định của Windows shell.
            # Verb này luôn in vào máy in MẶC ĐỊNH của hệ thống, nên nếu người
            # dùng chọn máy in khác, ta tạm đổi máy in mặc định, in xong rồi
            # trả lại máy in mặc định ban đầu.
            self._switch_default_printer_if_needed()
            os.startfile(filepath, "print")
            time.sleep(3)  # cho app mặc định kịp xử lý lệnh in

    def _switch_default_printer_if_needed(self):
        if not self.printer_name:
            return
        try:
            import win32print
            if self._original_default_printer is None:
                self._original_default_printer = win32print.GetDefaultPrinter()
            if self._original_default_printer != self.printer_name:
                win32print.SetDefaultPrinter(self.printer_name)
        except Exception as e:
            self.log(f"[CẢNH BÁO] Không đổi được máy in mặc định tạm thời: {e}")

    def restore_default_printer(self):
        if self._original_default_printer:
            try:
                import win32print
                win32print.SetDefaultPrinter(self._original_default_printer)
            except Exception:
                pass

    def _reset_word(self):
        """Ép tạo lại instance Word mới, phòng trường hợp app cũ đã bị
        treo/crash gây lỗi lặp lại ở các lần retry tiếp theo."""
        try:
            if self._word_app is not None:
                self._word_app.Quit()
        except Exception:
            pass
        self._word_app = None

    def _reset_excel(self):
        """Ép tạo lại instance Excel mới, tương tự _reset_word."""
        try:
            if self._excel_app is not None:
                self._excel_app.Quit()
        except Exception:
            pass
        self._excel_app = None

    def print_file(self, filepath):
        if not os.path.exists(filepath):
            self.log(f"FAILED\n{os.path.basename(filepath)} (File không tồn tại)")
            return False

        ext = os.path.splitext(filepath)[1].lower()

        # Retry logic: thử tối đa 3 lần (1 lần chạy + 2 lần retry)
        for attempt in range(3):
            try:
                if ext == ".pdf":
                    self.print_pdf(filepath)
                elif ext in (".doc", ".docx"):
                    self.print_word(filepath)
                elif ext in (".xls", ".xlsx"):
                    self.print_excel(filepath)
                else:
                    self.log(f"FAILED\n{os.path.basename(filepath)} (Loại file không hỗ trợ)")
                    return False

                self.log(f"SUCCESS\n{os.path.basename(filepath)}")
                return True
            except Exception as e:
                if attempt < 2:
                    self.log(f"[CẢNH BÁO] Lỗi in {os.path.basename(filepath)} (Thử lại {attempt+1}/2). Lỗi: {e}")
                    # Ép tạo lại instance Word/Excel mới cho lần retry, phòng
                    # trường hợp app cũ đã bị treo/crash gây lỗi lặp lại
                    if ext in (".doc", ".docx"):
                        self._reset_word()
                    elif ext in (".xls", ".xlsx"):
                        self._reset_excel()
                    time.sleep(2)
                else:
                    self.log(f"FAILED\n{os.path.basename(filepath)} (Lỗi sau 3 lần thử: {e})")
                    return False
        return False

    def cleanup(self):
        try:
            if self._word_app is not None:
                self._word_app.Quit()
        except Exception:
            pass
        try:
            if self._excel_app is not None:
                self._excel_app.Quit()
        except Exception:
            pass
        self.restore_default_printer()


# ----------------------------------------------------------------------
# Giao diện chính
# ----------------------------------------------------------------------
class MassPrintApp(BaseApp):
    CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

    def __init__(self):
        super().__init__()
        self.title("Lọc & In hàng loạt file theo mã công ty")
        self.geometry("880x600")

        self.folder_path = tk.StringVar()
        self.keyword = tk.StringVar()
        self.active_types = {k: tk.BooleanVar(value=True) for k in FILE_TYPE_GROUPS}
        self.all_files = []          # toàn bộ file trong folder (quét 1 lần)
        self.checked_files = set()   # lưu filepath của các file được tick chọn
        self.cancel_requested = threading.Event()  # cờ báo hiệu "Dừng in"
        self.sort_state = {}         # ghi nhớ chiều sắp xếp của từng cột

        self._filter_timer = None    # Timer cho Debounce ô tìm kiếm

        # Load config
        self.config = self._load_config()

        available_printers, default_printer = list_printers()
        self.available_printers = available_printers

        saved_printer = self.config.get("last_printer", "")
        default_p_label = f"(Mặc định: {default_printer})" if default_printer else "(Mặc định hệ thống)"
        if not saved_printer or saved_printer not in available_printers + [default_p_label]:
            saved_printer = default_p_label

        self.printer_choice = tk.StringVar(value=saved_printer)

        # File log: lưu cạnh script, đặt tên theo thời gian chạy chương trình
        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "print_logs")
        os.makedirs(log_dir, exist_ok=True)
        self.log_file_path = os.path.join(
            log_dir, f"print_log_{datetime.datetime.now():%Y%m%d_%H%M%S}.txt"
        )

        self._build_ui()

        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self._on_drop)

        # Tự động quét lại thư mục gần nhất
        last_folder = self.config.get("last_folder", "")
        if last_folder and os.path.isdir(last_folder):
            self.folder_path.set(last_folder)
            self._start_folder_scan(last_folder)

    def _load_config(self):
        try:
            if os.path.exists(self.CONFIG_FILE):
                with open(self.CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def _save_config(self):
        try:
            config = {
                "last_folder": self.folder_path.get(),
                "last_printer": self.printer_choice.get()
            }
            with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

    # ---------------- UI ----------------
    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}

        # Hàng chọn thư mục
        frm_top = ttk.Frame(self)
        frm_top.pack(fill="x", **pad)
        ttk.Button(frm_top, text="Chọn thư mục...", command=self.choose_folder).pack(side="left")
        ttk.Label(frm_top, textvariable=self.folder_path, foreground="gray").pack(side="left", padx=8)

        # Hàng tìm kiếm + loại file
        frm_filter = ttk.Frame(self)
        frm_filter.pack(fill="x", **pad)
        ttk.Label(frm_filter, text="Mã công ty (cách nhau bằng dấu phẩy):").pack(side="left")
        entry = ttk.Entry(frm_filter, textvariable=self.keyword, width=30)
        entry.pack(side="left", padx=6)

        # SỬ DỤNG DEBOUNCE TÌM KIẾM
        entry.bind("<KeyRelease>", self.on_search_key_release)

        for label in FILE_TYPE_GROUPS:
            cb = ttk.Checkbutton(
                frm_filter, text=label, variable=self.active_types[label],
                command=self.apply_filter
            )
            cb.pack(side="left", padx=4)

        ttk.Button(frm_filter, text="Lọc", command=self.apply_filter).pack(side="left", padx=10)

        # Bảng danh sách file
        columns = ("checked", "name", "type", "modified", "path")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("checked", text="✔")
        self.tree.heading("name", text="Tên file")
        self.tree.heading("type", text="Loại")
        self.tree.heading("modified", text="Ngày sửa")
        self.tree.heading("path", text="Đường dẫn")
        self.tree.column("checked", width=40, anchor="center")
        self.tree.column("name", width=280)
        self.tree.column("type", width=60, anchor="center")
        self.tree.column("modified", width=140, anchor="center")
        self.tree.column("path", width=280)
        self.tree.pack(fill="both", expand=True, padx=8, pady=6)
        self.tree.bind("<Button-1>", self.on_row_click)
        self.tree.bind("<Double-1>", self.on_row_double_click)
        for col in columns:
            self.tree.heading(col, command=lambda c=col: self.sort_by_column(c))

        # Hàng chọn máy in
        frm_printer = ttk.Frame(self)
        frm_printer.pack(fill="x", padx=8, pady=(0, 4))
        ttk.Label(frm_printer, text="Máy in:").pack(side="left")
        printer_values = list(self.available_printers)
        if not printer_values:
            printer_values = ["(Không tìm thấy máy in - dùng mặc định hệ thống)"]
        # Loại bỏ trùng lặp: nếu máy in đã lưu (last_printer) trùng tên với
        # 1 máy in thật trong danh sách, không thêm nó 2 lần vào dropdown
        combo_values = [self.printer_choice.get()] + [
            p for p in printer_values if p != self.printer_choice.get()
        ]
        self.printer_combo = ttk.Combobox(
            frm_printer, textvariable=self.printer_choice,
            values=combo_values,
            state="readonly", width=45
        )
        self.printer_combo.pack(side="left", padx=6)

        # Hàng nút hành động
        frm_actions = ttk.Frame(self)
        frm_actions.pack(fill="x", **pad)
        ttk.Button(frm_actions, text="Chọn tất cả", command=self.select_all).pack(side="left")
        ttk.Button(frm_actions, text="Bỏ chọn tất cả", command=self.deselect_all).pack(side="left", padx=6)
        self.count_label = ttk.Label(frm_actions, text="Đã lọc: 0 file | Đã chọn: 0")
        self.count_label.pack(side="left", padx=20)

        self.btn_print = ttk.Button(
            frm_actions, text="IN HÀNG LOẠT", command=self.start_print
        )
        self.btn_print.pack(side="right")

        self.btn_copy_move = ttk.Button(
            frm_actions, text="SAO CHÉP / DI CHUYỂN", command=self.start_copy_move
        )
        self.btn_copy_move.pack(side="right", padx=6)

        self.btn_cancel = ttk.Button(
            frm_actions, text="Dừng in", command=self.cancel_print, state="disabled"
        )
        self.btn_cancel.pack(side="right", padx=6)

        # Thanh tiến trình + nhãn trạng thái
        frm_progress = ttk.Frame(self)
        frm_progress.pack(fill="x", padx=8, pady=(0, 4))
        self.progress = ttk.Progressbar(frm_progress, mode="determinate")
        self.progress.pack(fill="x", side="left", expand=True)
        self.progress_label = ttk.Label(frm_progress, text="", width=14)
        self.progress_label.pack(side="left", padx=8)

        # Log
        self.log_box = tk.Text(self, height=8, state="disabled", bg="#111", fg="#0f0")
        self.log_box.pack(fill="x", padx=8, pady=6)

    # ---------------- Logic ----------------
    def log(self, msg):
        timestamped = f"[{time.strftime('%H:%M:%S')}]\n{msg}"
        self.log_box.configure(state="normal")
        self.log_box.insert("end", timestamped + "\n\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        try:
            with open(self.log_file_path, "a", encoding="utf-8") as f:
                f.write(timestamped + "\n\n")
        except Exception:
            pass  # không để lỗi ghi log làm gián đoạn việc in

    def _on_drop(self, event):
        path = event.data
        if path.startswith('{') and path.endswith('}'):
            path = path[1:-1]
        if os.path.isdir(path):
            self.folder_path.set(path)
            self._save_config()
            self._start_folder_scan(path)
        else:
            messagebox.showerror("Lỗi", "Chỉ chấp nhận kéo thả thư mục.")

    def choose_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_path.set(path)
            self._save_config()
            self._start_folder_scan(path)

    def _start_folder_scan(self, path):
        self.log(f"--- Đang quét thư mục: {path} ---")
        self.count_label.config(text="Đang quét...")
        self.tree.delete(*self.tree.get_children())
        self.all_files = []
        self.checked_files.clear()

        self.btn_print.config(state="disabled")
        if hasattr(self, 'btn_copy_move'):
            self.btn_copy_move.config(state="disabled")

        threading.Thread(target=self._scan_worker, args=(path,), daemon=True).start()

    def _scan_worker(self, path):
        scanned_files = []
        file_count = 0
        for root, _, files in os.walk(path):
            for f in files:
                if f.startswith("~$"):
                    continue
                ext = os.path.splitext(f)[1].lower()
                if ext in ALL_EXTENSIONS:
                    full = os.path.join(root, f)
                    try:
                        mtime = time.strftime(
                            "%d/%m/%Y %H:%M", time.localtime(os.path.getmtime(full))
                        )
                    except OSError:
                        mtime = ""
                    scanned_files.append({"name": f, "path": full, "ext": ext, "mtime": mtime})
                    file_count += 1
                    if file_count % 50 == 0:
                        self.after(0, lambda c=file_count: self.count_label.config(text=f"Đang quét {c} file..."))

        self.after(0, self._on_scan_finished, scanned_files)

    def _on_scan_finished(self, scanned_files):
        self.all_files = scanned_files
        self.log(f"Đã quét xong {len(self.all_files)} file (PDF/DOC/XLS) trong thư mục.")
        self.btn_print.config(state="normal")
        if hasattr(self, 'btn_copy_move'):
            self.btn_copy_move.config(state="normal")
        self.apply_filter()

    def on_search_key_release(self, event):
        if self._filter_timer is not None:
            self.after_cancel(self._filter_timer)
        self._filter_timer = self.after(300, self.apply_filter)

    def apply_filter(self):
        raw_keyword = self.keyword.get().strip().lower()
        keywords = [k.strip() for k in raw_keyword.split(",") if k.strip()]
        allowed_exts = []
        for label, exts in FILE_TYPE_GROUPS.items():
            if self.active_types[label].get():
                allowed_exts.extend(exts)

        self.tree.delete(*self.tree.get_children())

        def name_matches(name_lower):
            if not keywords:
                return True
            return any(k in name_lower for k in keywords)

        matched = [
            f for f in self.all_files
            if f["ext"] in allowed_exts and name_matches(f["name"].lower())
        ]

        type_label = {".pdf": "PDF", ".doc": "DOC", ".docx": "DOC", ".xls": "XLS", ".xlsx": "XLS"}
        for f in matched:
            check_mark = "☑" if f["path"] in self.checked_files else "☐"
            self.tree.insert(
                "", "end",
                values=(check_mark, f["name"], type_label.get(f["ext"], f["ext"]), f["mtime"], f["path"])
            )
        self.update_count()

    def on_row_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        if not row:
            return
        if col == "#1":  # cột checkbox
            self.toggle_check(row)

    def on_row_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row = self.tree.identify_row(event.y)
        if not row:
            return
        filepath = self.tree.item(row, "values")[4]
        try:
            os.startfile(filepath)
            self.log(f"[XEM TRƯỚC] Đã mở: {os.path.basename(filepath)}")
        except Exception as e:
            messagebox.showerror("Không mở được file", str(e))

    def sort_by_column(self, col):
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        reverse = self.sort_state.get(col, False)

        if col == "modified":
            def parse_date(date_str):
                try:
                    return datetime.datetime.strptime(date_str, "%d/%m/%Y %H:%M")
                except ValueError:
                    return datetime.datetime.min
            items.sort(key=lambda t: parse_date(t[0]), reverse=reverse)
        elif col in ("name", "path"):
            items.sort(key=lambda t: natural_sort_key(t[0]), reverse=reverse)
        else:
            items.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)

        for index, (_, k) in enumerate(items):
            self.tree.move(k, "", index)
        self.sort_state[col] = not reverse

    def toggle_check(self, item_id):
        vals = list(self.tree.item(item_id, "values"))
        filepath = vals[4]
        if filepath in self.checked_files:
            self.checked_files.remove(filepath)
            vals[0] = "☐"
        else:
            self.checked_files.add(filepath)
            vals[0] = "☑"
        self.tree.item(item_id, values=vals)
        self.update_count()

    def select_all(self):
        for item_id in self.tree.get_children():
            vals = list(self.tree.item(item_id, "values"))
            vals[0] = "☑"
            self.tree.item(item_id, values=vals)
            self.checked_files.add(vals[4])
        self.update_count()

    def deselect_all(self):
        for item_id in self.tree.get_children():
            vals = list(self.tree.item(item_id, "values"))
            vals[0] = "☐"
            self.tree.item(item_id, values=vals)
        self.checked_files.clear()
        self.update_count()

    def start_copy_move(self):
        if not self.checked_files:
            messagebox.showwarning("Chưa chọn file", "Vui lòng chọn ít nhất một file.")
            return

        dest_dir = filedialog.askdirectory(title="Chọn thư mục đích")
        if not dest_dir:
            return

        ans = messagebox.askyesnocancel("Xác nhận", "Bạn có muốn xóa file gốc sau khi copy không?\\n\\nYes = Di chuyển (Cut)\\nNo = Chỉ Copy\\nCancel = Hủy")
        if ans is None:
            return

        is_move = ans

        files_to_copy = sorted(list(self.checked_files), key=lambda x: natural_sort_key(os.path.basename(x)))

        self.cancel_requested.clear()
        self.progress.config(maximum=len(files_to_copy), value=0)
        self.progress_label.config(text=f"0 / {len(files_to_copy)}")
        self.btn_print.config(state="disabled")
        self.btn_copy_move.config(state="disabled")
        self.btn_cancel.config(state="normal")

        threading.Thread(
            target=self._copy_worker, args=(files_to_copy, dest_dir, is_move), daemon=True
        ).start()

    def _get_company_code(self, filename):
        match = re.search(r'([a-zA-Z]+\\d+)', filename)
        if match:
            return match.group(1).upper()
        return "Khác"

    def _ask_conflict_action(self, filename):
        result_container = []
        apply_to_all_container = [False]
        event = threading.Event()

        def _show_dialog():
            dlg = tk.Toplevel(self)
            dlg.title("File đã tồn tại")
            dlg.geometry("300x180")
            dlg.resizable(False, False)
            dlg.grab_set()

            ttk.Label(dlg, text=f"File {filename} đã tồn tại.\\nBạn muốn xử lý thế nào?").pack(pady=10)

            def set_res(act):
                result_container.append(act)
                apply_to_all_container[0] = chk_var.get()
                dlg.destroy()
                event.set()

            chk_var = tk.BooleanVar(value=False)
            ttk.Checkbutton(dlg, text="Áp dụng cho tất cả", variable=chk_var).pack(pady=5)

            btn_frame = ttk.Frame(dlg)
            btn_frame.pack(pady=10)

            ttk.Button(btn_frame, text="Đổi tên tự động", command=lambda: set_res("rename")).pack(side="left", padx=5)
            ttk.Button(btn_frame, text="Ghi đè", command=lambda: set_res("overwrite")).pack(side="left", padx=5)
            ttk.Button(btn_frame, text="Bỏ qua", command=lambda: set_res("skip")).pack(side="left", padx=5)

            dlg.protocol("WM_DELETE_WINDOW", lambda: set_res("skip"))

        self.after(0, _show_dialog)
        event.wait()
        return result_container[0], apply_to_all_container[0]

    def _get_auto_renamed_path(self, target_dir, filename):
        base, ext = os.path.splitext(filename)
        counter = 1
        while True:
            new_name = f"{base} ({counter}){ext}"
            new_path = os.path.join(target_dir, new_name)
            if not os.path.exists(new_path):
                return new_path
            counter += 1

    def _copy_worker(self, files, dest_dir, is_move):
        total = len(files)
        action_name = "Di chuyển" if is_move else "Sao chép"
        self.log(f"--- Bắt đầu {action_name.lower()} {total} file ---")

        success_copy = 0
        success_move = 0
        skipped = 0
        failed = 0
        processed = 0
        cancelled = False

        global_conflict_action = None

        for path in files:
            if self.cancel_requested.is_set():
                cancelled = True
                self.log(f"--- Đã dừng theo yêu cầu. Đã {action_name.lower()} {processed}/{total} file trước khi dừng. ---")
                break

            filename = os.path.basename(path)
            company_code = self._get_company_code(filename)
            target_dir = os.path.join(dest_dir, company_code)
            os.makedirs(target_dir, exist_ok=True)
            
            dest_path = os.path.join(target_dir, filename)

            try:
                if os.path.exists(dest_path):
                    if global_conflict_action:
                        action = global_conflict_action
                    else:
                        action, apply_all = self._ask_conflict_action(filename)
                        if apply_all:
                            global_conflict_action = action
                    
                    if action == "skip":
                        skipped += 1
                        self.log(f"[SKIPPED]\n{filename}")
                        processed += 1
                        self.after(0, self._update_progress, processed, total)
                        continue
                    elif action == "rename":
                        dest_path = self._get_auto_renamed_path(target_dir, filename)
                    elif action == "overwrite":
                        pass

                if is_move:
                    if os.path.exists(dest_path):
                        os.replace(path, dest_path)
                    else:
                        shutil.move(path, dest_path)
                    success_move += 1
                    self.log(f"[MOVE]\n{filename}")
                else:
                    if os.path.exists(dest_path):
                        os.remove(dest_path)
                    shutil.copy2(path, dest_path)
                    success_copy += 1
                    self.log(f"[COPY]\n{filename}")

            except Exception as e:
                failed += 1
                self.log(f"[FAILED]\n{filename} - {str(e)}")

            processed += 1
            self.after(0, self._update_progress, processed, total)

        if not cancelled:
            self.log(f"--- Hoàn tất {action_name.lower()} ---")

        self.after(0, self._on_copy_finished, success_copy, success_move, skipped, failed, is_move)

    def _on_copy_finished(self, success_copy, success_move, skipped, failed, is_move):
        self.btn_print.config(state="normal")
        self.btn_copy_move.config(state="normal")
        self.btn_cancel.config(state="disabled")

        summary = (
            "====================\n"
            "Copy thành công:\n"
            f"{success_copy}\n\n"
            "Move thành công:\n"
            f"{success_move}\n\n"
            "Bỏ qua:\n"
            f"{skipped}\n\n"
            "Lỗi:\n"
            f"{failed}\n"
            "===================="
        )
        messagebox.showinfo("Tổng kết", summary)
        
        if is_move:
            self._start_folder_scan(self.folder_path.get())

    def update_count(self):
        total = len(self.tree.get_children())
        self.count_label.config(
            text=f"Đã lọc: {total} file | Đã chọn: {len(self.checked_files)}"
        )

    def start_print(self):
        if not self.checked_files:
            messagebox.showwarning("Chưa chọn file", "Vui lòng chọn ít nhất 1 file để in.")
            return

        # Lấy danh sách in và sort natural theo tên file
        files_to_print = sorted(list(self.checked_files), key=lambda x: natural_sort_key(os.path.basename(x)))

        warn_threshold = 50
        confirm_msg = f"Sẽ in {len(files_to_print)} file lên máy in mặc định.\nBạn có chắc chắn muốn tiếp tục?"
        if len(files_to_print) >= warn_threshold:
            confirm_msg = (
                f"⚠ Bạn sắp in SỐ LƯỢNG LỚN: {len(files_to_print)} file.\n"
                "Việc này có thể mất nhiều thời gian và tốn giấy/mực.\n"
                "Bạn có chắc chắn muốn tiếp tục?"
            )
        if not messagebox.askyesno("Xác nhận in", confirm_msg):
            return

        self._save_config()

        self.cancel_requested.clear()
        self.progress.config(maximum=len(files_to_print), value=0)
        self.progress_label.config(text=f"0 / {len(files_to_print)}")
        self.btn_print.config(state="disabled")
        self.btn_cancel.config(state="normal")

        selected_printer = self.printer_choice.get()
        if selected_printer not in self.available_printers:
            selected_printer = None

        threading.Thread(
            target=self._print_worker, args=(files_to_print, selected_printer), daemon=True
        ).start()

    def cancel_print(self):
        self.cancel_requested.set()
        self.log("--- Đã nhận yêu cầu DỪNG. Sẽ dừng sau khi in xong file hiện tại... ---")
        self.btn_cancel.config(state="disabled")

    def _print_worker(self, files, printer_name=None):
        total = len(files)
        self.log(f"--- Bắt đầu in {total} file (máy in: {printer_name or 'mặc định hệ thống'}) ---")
        printer = Printer(self.log, printer_name=printer_name)

        success = 0
        failed = 0
        processed = 0
        cancelled = False

        for path in files:
            if self.cancel_requested.is_set():
                cancelled = True
                self.log(f"--- Đã dừng theo yêu cầu. Đã in {processed}/{total} file trước khi dừng. ---")
                break

            if printer.print_file(path):
                success += 1
            else:
                failed += 1

            processed += 1
            self.after(0, self._update_progress, processed, total)
            time.sleep(1)

        printer.cleanup()

        if not cancelled:
            self.log(f"--- Hoàn tất: {success}/{total} file gửi in thành công ---")

        unprocessed = total - success - failed
        self.after(0, self._on_print_finished, success, failed, unprocessed)

    def _update_progress(self, processed, total):
        self.progress.config(value=processed)
        self.progress_label.config(text=f"{processed} / {total}")

    def _on_print_finished(self, success, failed, unprocessed):
        self.btn_print.config(state="normal")
        self.btn_cancel.config(state="disabled")

        summary = (
            "==================\n"
            "Đã in thành công:\n"
            f"{success} file\n\n"
            "Lỗi:\n"
            f"{failed} file\n\n"
            "Bị hủy:\n"
            f"{unprocessed} file\n"
            "=================="
        )
        messagebox.showinfo("Báo cáo In", summary)


if __name__ == "__main__":
    app = MassPrintApp()
    app.mainloop()