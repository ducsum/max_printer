import os
import sys
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

from config import AppConfig
from logger import AppLogger
from printer import Printer
from copy_manager import CopyManager
from scanner import FolderScanner
from utils import FILE_TYPE_GROUPS, natural_sort_key, format_size, format_time
import dialogs

logging.basicConfig(level=logging.INFO)

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

if HAS_DND:
    class BaseApp(TkinterDnD.Tk):
        pass
else:
    class BaseApp(tk.Tk):
        pass

class MassPrintApp(BaseApp):
    CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

    def __init__(self):
        super().__init__()
        self.title("Lọc & In hàng loạt file theo mã công ty (Production Edition)")
        
        self.cfg = AppConfig(self.CONFIG_FILE)
        geo = self.cfg.get("window_geometry", "880x600")
        self.geometry(geo)

        # Thiết lập icon cho ứng dụng
        self.icon_path = resource_path("icon.ico")
        if os.path.exists(self.icon_path):
            try:
                self.iconbitmap(self.icon_path)
            except Exception:
                pass


        self.folder_path = tk.StringVar()
        self.keyword = tk.StringVar(value=self.cfg.get("last_keyword", ""))
        self.active_types = {}
        last_types = self.cfg.get("last_active_types", {})
        for k in FILE_TYPE_GROUPS:
            self.active_types[k] = tk.BooleanVar(value=last_types.get(k, True))

        self.all_files = []
        self.file_by_path = {}
        self.checked_files = set()
        
        self.print_cancel_requested = threading.Event()
        self.copy_cancel_requested = threading.Event()
        
        self.sort_column = self.cfg.get("sort_column", "name")
        self.sort_reverse = self.cfg.get("sort_reverse", False)

        self._filter_timer = None
        self._progress_start_time = 0

        available_printers, default_printer = list_printers()
        self.available_printers = available_printers

        saved_printer = self.cfg.get("last_printer", "")
        default_p_label = f"(Mặc định: {default_printer})" if default_printer else "(Mặc định hệ thống)"
        if not saved_printer or saved_printer not in available_printers + [default_p_label]:
            saved_printer = default_p_label

        self.printer_choice = tk.StringVar(value=saved_printer)

        self.logger = AppLogger()
        self._build_ui()
        self.logger.log_box = self.log_box

        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self._on_drop)
            
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        last_folder = self.cfg.get("last_folder", "")
        if last_folder and os.path.isdir(last_folder):
            self.folder_path.set(last_folder)
            self._start_folder_scan(last_folder)
        else:
            self.status_var.set("Ready")

    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}

        frm_top = ttk.Frame(self)
        frm_top.pack(fill="x", **pad)
        self.btn_select_folder = ttk.Button(frm_top, text="Chọn thư mục...", command=self.choose_folder)
        self.btn_select_folder.pack(side="left")
        ttk.Label(frm_top, textvariable=self.folder_path, foreground="gray").pack(side="left", padx=8)

        # Donate Button
        self.btn_donate = ttk.Button(frm_top, text="🎁 Donate", command=self.show_donate_popup)
        self.btn_donate.pack(side="right")

        frm_filter = ttk.Frame(self)
        frm_filter.pack(fill="x", **pad)
        ttk.Label(frm_filter, text="Mã công ty (cách nhau dấu phẩy):").pack(side="left")
        self.entry_keyword = ttk.Entry(frm_filter, textvariable=self.keyword, width=30)
        self.entry_keyword.pack(side="left", padx=6)
        self.entry_keyword.bind("<KeyRelease>", self.on_search_key_release)

        self.type_cbs = []
        for label in FILE_TYPE_GROUPS:
            cb = ttk.Checkbutton(
                frm_filter, text=label, variable=self.active_types[label],
                command=self.apply_filter
            )
            cb.pack(side="left", padx=4)
            self.type_cbs.append(cb)

        self.btn_filter = ttk.Button(frm_filter, text="Lọc", command=self.apply_filter)
        self.btn_filter.pack(side="left", padx=10)

        columns = ("checked", "stt", "name", "type", "size", "modified", "path")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("checked", text="✔")
        self.tree.heading("stt", text="STT")
        self.tree.heading("name", text="Tên file", command=lambda: self.sort_by_column("name"))
        self.tree.heading("type", text="Loại", command=lambda: self.sort_by_column("type"))
        self.tree.heading("size", text="Kích thước", command=lambda: self.sort_by_column("size"))
        self.tree.heading("modified", text="Ngày sửa", command=lambda: self.sort_by_column("modified"))
        self.tree.heading("path", text="Đường dẫn", command=lambda: self.sort_by_column("path"))
        
        self.tree.column("checked", width=40, anchor="center")
        self.tree.column("stt", width=50, anchor="center")
        self.tree.column("name", width=250)
        self.tree.column("type", width=60, anchor="center")
        self.tree.column("size", width=80, anchor="e")
        self.tree.column("modified", width=140, anchor="center")
        self.tree.column("path", width=250)
        self.tree.pack(fill="both", expand=True, padx=8, pady=6)
        
        self.tree.bind("<Button-1>", self.on_row_click)
        self.tree.bind("<Double-1>", self.on_row_double_click)
        self.tree.bind("<Button-3>", self.show_context_menu)
        
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Mở thư mục chứa file", command=self.open_file_location)
        self.context_menu.add_command(label="Copy đường dẫn", command=self.copy_file_path)

        frm_printer = ttk.Frame(self)
        frm_printer.pack(fill="x", padx=8, pady=(0, 4))
        ttk.Label(frm_printer, text="Máy in:").pack(side="left")
        printer_values = list(self.available_printers)
        if not printer_values:
            printer_values = ["(Không tìm thấy máy in - dùng mặc định hệ thống)"]
        combo_values = [self.printer_choice.get()] + [p for p in printer_values if p != self.printer_choice.get()]
        
        self.printer_combo = ttk.Combobox(
            frm_printer, textvariable=self.printer_choice,
            values=combo_values, state="readonly", width=45
        )
        self.printer_combo.pack(side="left", padx=6)

        frm_actions = ttk.Frame(self)
        frm_actions.pack(fill="x", **pad)
        self.btn_select_all = ttk.Button(frm_actions, text="Chọn tất cả", command=self.select_all)
        self.btn_select_all.pack(side="left")
        self.btn_deselect_all = ttk.Button(frm_actions, text="Bỏ chọn tất cả", command=self.deselect_all)
        self.btn_deselect_all.pack(side="left", padx=6)
        self.count_label = ttk.Label(frm_actions, text="Đã lọc: 0 | Đã chọn: 0 (0 B)")
        self.count_label.pack(side="left", padx=20)

        self.btn_print = ttk.Button(frm_actions, text="IN HÀNG LOẠT", command=self.start_print)
        self.btn_print.pack(side="right")
        self.btn_copy_move = ttk.Button(frm_actions, text="SAO CHÉP / DI CHUYỂN", command=self.start_copy_move)
        self.btn_copy_move.pack(side="right", padx=6)
        self.btn_cancel = ttk.Button(frm_actions, text="Dừng in", command=self.cancel_print, state="disabled")
        self.btn_cancel.pack(side="right", padx=6)

        frm_progress = ttk.Frame(self)
        frm_progress.pack(fill="x", padx=8, pady=(0, 4))
        self.progress = ttk.Progressbar(frm_progress, mode="determinate")
        self.progress.pack(fill="x", side="left", expand=True)
        self.progress_label = ttk.Label(frm_progress, text="", width=25)
        self.progress_label.pack(side="left", padx=8)

        self.log_box = tk.Text(self, height=8, state="disabled", bg="#111", fg="#0f0")
        self.log_box.pack(fill="x", padx=8, pady=6)
        
        # Status Bar
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(side="bottom", fill="x")

    def _lock_ui(self, lock: bool):
        state = "disabled" if lock else "normal"
        self.btn_select_folder.config(state=state)
        self.entry_keyword.config(state=state)
        self.btn_filter.config(state=state)
        self.btn_select_all.config(state=state)
        self.btn_deselect_all.config(state=state)
        self.printer_combo.config(state="disabled" if lock else "readonly")
        
        for cb in self.type_cbs:
            cb.config(state=state)

    def on_close(self):
        self.cfg.set("window_geometry", self.geometry())
        self.cfg.set("last_keyword", self.keyword.get())
        
        types = {k: v.get() for k, v in self.active_types.items()}
        self.cfg.set("last_active_types", types)
        
        self.destroy()

    def _on_drop(self, event):
        path = event.data
        if path.startswith('{') and path.endswith('}'):
            path = path[1:-1]
        if os.path.isdir(path):
            self.folder_path.set(path)
            self.cfg.set("last_folder", path)
            self._start_folder_scan(path)
        else:
            messagebox.showerror("Lỗi", "Chỉ chấp nhận kéo thả thư mục.")

    def choose_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_path.set(path)
            self.cfg.set("last_folder", path)
            self._start_folder_scan(path)

    def _start_folder_scan(self, path):
        self._lock_ui(True)
        self.btn_print.config(state="disabled")
        self.btn_copy_move.config(state="disabled")
        
        self.scan_cancel_requested = threading.Event()
        self.btn_cancel.config(command=self.cancel_scan, state="normal", text="Dừng quét")
        
        self.status_var.set("Scanning...")
        self.tree.delete(*self.tree.get_children())
        self.all_files = []
        self.file_by_path = {}
        self.checked_files.clear()
        self.update_count()
        
        scanner = FolderScanner(
            path, 
            progress_callback=lambda f, fd, e, s: self.after(0, self._update_scan_progress, f, fd, e, s),
            finished_callback=lambda files, size: self.after(0, self._on_scan_finished, files),
            cancel_event=self.scan_cancel_requested
        )
        threading.Thread(target=scanner.scan, daemon=True).start()

    def cancel_scan(self):
        if hasattr(self, 'scan_cancel_requested'):
            self.scan_cancel_requested.set()
        self.btn_cancel.config(state="disabled")

    def _update_scan_progress(self, files, folders, elapsed, speed):
        self.status_var.set(f"Scanning... Files: {files} | Folders: {folders} | Elapsed: {format_time(elapsed)} | Speed: {speed:.1f} f/s")

    def _on_scan_finished(self, scanned_files):
        self.all_files = scanned_files
        self.file_by_path = {f["path"]: f for f in scanned_files}
        
        self.status_var.set("Ready")
        msg = f"Đã quét xong {len(self.all_files)} file."
        if hasattr(self, 'scan_cancel_requested') and self.scan_cancel_requested.is_set():
            msg += " (Đã dừng)"
        self.logger.log(msg, after_callback=self.after)
        
        self.btn_cancel.config(text="Dừng in", state="disabled")
        self._lock_ui(False)
        self.btn_print.config(state="normal")
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

        type_label = {".pdf": "PDF", ".doc": "DOC", ".docx": "DOC", ".xls": "XLS", ".xlsx": "XLS"}
        
        # Sử dụng Generator thay vì tạo List trung gian giúp tiết kiệm RAM và giảm độ trễ
        matched_gen = (
            f for f in self.all_files
            if f["ext"] in allowed_exts and name_matches(f["name_lower"])
        )

        for f in matched_gen:
            check_mark = "☑" if f["path"] in self.checked_files else "☐"
            self.tree.insert(
                "", "end",
                values=(check_mark, "", f["name"], type_label.get(f["ext"], f["ext"]), 
                        format_size(f["size"]), f["mtime"], f["path"])
            )
        
        self.sort_by_column(self.sort_column, self.sort_reverse)
        self.update_count()

    def sort_by_column(self, col, force_reverse=None):
        self.sort_column = col
        
        if force_reverse is not None:
            reverse = force_reverse
        else:
            reverse = not self.sort_reverse
            self.sort_reverse = reverse
            self.cfg.set("sort_column", col)
            self.cfg.set("sort_reverse", reverse)

        items = [(self.tree.set(k, "path"), k) for k in self.tree.get_children("")]

        keywords = [k.strip() for k in self.keyword.get().strip().lower().split(",") if k.strip()]
        def get_search_score(path):
            if not keywords:
                return 3
            f = self.file_by_path.get(path)
            if not f:
                return 3
            name_lower = f["name_lower"]
            best = 3
            for k in keywords:
                if k == name_lower: best = min(best, 0)
                elif name_lower.startswith(k): best = min(best, 1)
                elif k in name_lower: best = min(best, 2)
            return best

        if col == "name":
            items.sort(key=lambda t: (get_search_score(t[0]), natural_sort_key(self.file_by_path.get(t[0], {}).get("name", ""))), reverse=reverse)
        elif col == "size":
            items.sort(key=lambda t: (get_search_score(t[0]), self.file_by_path.get(t[0], {}).get("size", 0)), reverse=reverse)
        elif col == "modified":
            items.sort(key=lambda t: (get_search_score(t[0]), self.file_by_path.get(t[0], {}).get("mtime", "")), reverse=reverse)
        elif col == "type":
            items.sort(key=lambda t: (get_search_score(t[0]), self.file_by_path.get(t[0], {}).get("ext", "")), reverse=reverse)
        else:
            items.sort(key=lambda t: (get_search_score(t[0]), t[0]), reverse=reverse)

        for index, (_, k) in enumerate(items):
            self.tree.move(k, "", index)
            
        self._update_stt()

    def _update_stt(self):
        children = self.tree.get_children()
        # Ngăn chặn Tkinter UI Freeze (O(N) update event) nếu số lượng quá lớn
        if len(children) > 10000:
            return
            
        for index, item_id in enumerate(children, start=1):
            vals = list(self.tree.item(item_id, "values"))
            vals[1] = str(index)
            self.tree.item(item_id, values=vals)

    def on_row_click(self, event):
        if self.btn_print.cget("state") == "disabled":
            return # locked
            
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        if not row:
            return
        if col == "#1":
            self.toggle_check(row)

    def on_row_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row = self.tree.identify_row(event.y)
        if not row:
            return
        filepath = self.tree.item(row, "values")[6]
        try:
            os.startfile(filepath)
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    def show_context_menu(self, event):
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
            self.context_menu.post(event.x_root, event.y_root)

    def open_file_location(self):
        selected = self.tree.selection()
        if selected:
            filepath = self.tree.item(selected[0], "values")[6]
            os.startfile(os.path.dirname(filepath))
            
    def copy_file_path(self):
        selected = self.tree.selection()
        if selected:
            filepath = self.tree.item(selected[0], "values")[6]
            self.clipboard_clear()
            self.clipboard_append(filepath)

    def toggle_check(self, item_id):
        vals = list(self.tree.item(item_id, "values"))
        filepath = vals[6]
        if filepath in self.checked_files:
            self.checked_files.remove(filepath)
            vals[0] = "☐"
        else:
            self.checked_files.add(filepath)
            vals[0] = "☑"
        self.tree.item(item_id, values=vals)
        self.update_count()

    def _batch_update_checkbox(self, children, symbol, start_idx):
        chunk_size = 1000
        end_idx = min(start_idx + chunk_size, len(children))
        for i in range(start_idx, end_idx):
            item_id = children[i]
            vals = list(self.tree.item(item_id, "values"))
            vals[0] = symbol
            self.tree.item(item_id, values=vals)
            
        if end_idx < len(children):
            self.after(10, self._batch_update_checkbox, children, symbol, end_idx)

    def select_all(self):
        children = self.tree.get_children()
        # Update underlying data fast
        for item_id in children:
            self.checked_files.add(self.tree.item(item_id, "values")[6])
        self.update_count()
        # Non-blocking batch update for UI
        self._batch_update_checkbox(children, "☑", 0)

    def deselect_all(self):
        children = self.tree.get_children()
        # Update underlying data fast
        for item_id in children:
            self.checked_files.discard(self.tree.item(item_id, "values")[6])
        self.update_count()
        # Non-blocking batch update for UI
        self._batch_update_checkbox(children, "☐", 0)
        
    def update_count(self):
        total_items = len(self.tree.get_children())
        total_selected_size = sum(self.file_by_path.get(p, {}).get("size", 0) for p in self.checked_files)
                
        self.count_label.config(
            text=f"Đã lọc: {total_items} | Đã chọn: {len(self.checked_files)} ({format_size(total_selected_size)})"
        )

    def cancel_print(self):
        self.print_cancel_requested.set()
        self.logger.log("--- Đã nhận yêu cầu DỪNG IN. Sẽ dừng sau file hiện tại... ---", after_callback=self.after)
        self.btn_cancel.config(state="disabled")
        
    def cancel_copy(self):
        self.copy_cancel_requested.set()
        self.logger.log("--- Đã nhận yêu cầu DỪNG COPY/MOVE. Sẽ dừng sau file hiện tại... ---", after_callback=self.after)
        self.btn_cancel.config(state="disabled")

    def _update_progress(self, processed, total):
        self.progress.config(value=processed)
        
        elapsed = time.time() - self._progress_start_time
        speed = processed / elapsed if elapsed > 0 else 0
        
        if speed > 0 and processed < total:
            eta = (total - processed) / speed
            eta_str = f"ETA: {format_time(eta)}"
        else:
            eta_str = ""
            
        self.progress_label.config(text=f"{processed}/{total} | {speed:.1f} file/s | {eta_str}")

    def start_print(self):
        if not self.checked_files:
            messagebox.showwarning("Chưa chọn file", "Vui lòng chọn ít nhất 1 file để in.")
            return

        files_to_print = sorted(list(self.checked_files), key=lambda x: natural_sort_key(os.path.basename(x)))

        if not messagebox.askyesno("Xác nhận in", f"Sẽ in {len(files_to_print)} file.\nTiếp tục?"):
            return

        self.cfg.set("last_printer", self.printer_choice.get())

        self.print_cancel_requested.clear()
        self._progress_start_time = time.time()
        self.progress.config(maximum=len(files_to_print), value=0)
        self.progress_label.config(text=f"0 / {len(files_to_print)}")
        
        self._lock_ui(True)
        self.btn_print.config(state="disabled")
        self.btn_copy_move.config(state="disabled")
        self.btn_cancel.config(command=self.cancel_print, state="normal")
        self.status_var.set("Printing...")

        printer_name = self.printer_choice.get()
        if printer_name not in self.available_printers:
            printer_name = None

        threading.Thread(
            target=self._print_worker, args=(files_to_print, printer_name), daemon=True
        ).start()

    def _print_worker(self, files, printer_name):
        total = len(files)
        printer = Printer(self.logger, printer_name=printer_name)
        
        self.logger.log(f"--- Bắt đầu in {total} file ---", after_callback=self.after)

        success, failed, processed = 0, 0, 0
        cancelled = False
        last_ui_update = time.time()

        for path in files:
            if self.print_cancel_requested.is_set():
                cancelled = True
                break
                
            is_ok, dur = printer.print_file(path)
            if is_ok:
                success += 1
            else:
                failed += 1
                
            processed += 1
            current_time = time.time()
            if current_time - last_ui_update >= 0.1:
                self.after(0, self._update_progress, processed, total)
                last_ui_update = current_time

        printer.cleanup()
        
        self.after(0, self._update_progress, processed, total)
        
        if not cancelled:
            self.logger.log(f"--- Hoàn tất gửi {success}/{total} file ---", after_callback=self.after)
            
        self.after(0, self._on_print_finished, success, failed, total - success - failed)

    def _on_print_finished(self, success, failed, unprocessed):
        self._lock_ui(False)
        self.btn_print.config(state="normal")
        self.btn_copy_move.config(state="normal")
        self.btn_cancel.config(state="disabled")
        self.status_var.set("Ready")

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

    def ask_conflict(self, filename):
        event = threading.Event()
        res = [None, False]
        
        def callback(action, apply_all):
            res[0] = action
            res[1] = apply_all
            event.set()
            
        self.after(0, dialogs.ask_conflict_action, self, filename, callback)
        event.wait()
        return res[0], res[1]

    def start_copy_move(self):
        if not self.checked_files:
            messagebox.showwarning("Chưa chọn file", "Vui lòng chọn ít nhất một file.")
            return

        initial_dir = self.cfg.get("last_destination_folder", "")
        dest_dir = filedialog.askdirectory(title="Chọn thư mục đích", initialdir=initial_dir if os.path.exists(initial_dir) else "")
        if not dest_dir:
            return

        self.cfg.set("last_destination_folder", dest_dir)

        ans = messagebox.askyesnocancel("Xác nhận", "Bạn có muốn xóa file gốc sau khi copy không?\n\nYes = Di chuyển (Cut)\nNo = Chỉ Copy\nCancel = Hủy")
        if ans is None:
            return

        is_move = ans
        files_to_copy = sorted(list(self.checked_files), key=lambda x: natural_sort_key(os.path.basename(x)))

        if is_move:
            if not messagebox.askyesno("Xác nhận di chuyển", f"Bạn sắp DI CHUYỂN {len(files_to_copy)} file.\nSau khi hoàn thành,\nfile gốc sẽ bị xóa.\nBạn có chắc chắn muốn tiếp tục?"):
                return

        self.copy_cancel_requested.clear()
        self._progress_start_time = time.time()
        self.progress.config(maximum=len(files_to_copy), value=0)
        self.progress_label.config(text=f"0 / {len(files_to_copy)}")
        
        self._lock_ui(True)
        self.btn_print.config(state="disabled")
        self.btn_copy_move.config(state="disabled")
        self.btn_cancel.config(command=self.cancel_copy, state="normal")
        self.status_var.set("Copying..." if not is_move else "Moving...")

        cm = CopyManager(self)
        threading.Thread(
            target=cm.run, args=(files_to_copy, dest_dir, is_move), daemon=True
        ).start()

    def _on_copy_finished(self, success_copy, success_move, skipped, failed, is_move, dest_dir):
        self._lock_ui(False)
        self.btn_print.config(state="normal")
        self.btn_copy_move.config(state="normal")
        self.btn_cancel.config(state="disabled")
        self.status_var.set("Ready")

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

        if messagebox.askyesno("Đã hoàn thành", "Đã hoàn thành.\nBạn có muốn mở thư mục đích không?"):
            try:
                os.startfile(dest_dir)
            except Exception:
                pass

    def show_donate_popup(self):
        if not HAS_PIL: return
        popup = tk.Toplevel(self)
        popup.title("Donate ❤️")
        popup.geometry("290x370")
        popup.transient(self)
        popup.grab_set()
        img_path = resource_path("QR.jpg")
        if os.path.exists(img_path):
            img = Image.open(img_path).resize((270, 348), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            lbl = tk.Label(popup, image=photo, bg="#f0f0f0")
            lbl.image = photo
            lbl.pack(pady=10)


if __name__ == "__main__":
    app = MassPrintApp()
    app.mainloop()