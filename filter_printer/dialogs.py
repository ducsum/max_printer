import tkinter as tk
from tkinter import ttk

def ask_conflict_action(parent, filename: str, callback):
    """
    Hiển thị hộp thoại xử lý conflict (Ghi đè, Bỏ qua, Đổi tên).
    callback(action: str, apply_to_all: bool)
    """
    dlg = tk.Toplevel(parent)
    dlg.title("File đã tồn tại")
    dlg.geometry("300x180")
    dlg.resizable(False, False)
    dlg.grab_set()

    ttk.Label(dlg, text=f"File {filename} đã tồn tại.\nBạn muốn xử lý thế nào?").pack(pady=10)

    chk_var = tk.BooleanVar(value=False)
    
    def set_res(act):
        apply_all = chk_var.get()
        dlg.destroy()
        callback(act, apply_all)

    ttk.Checkbutton(dlg, text="Áp dụng cho tất cả", variable=chk_var).pack(pady=5)

    btn_frame = ttk.Frame(dlg)
    btn_frame.pack(pady=10)

    ttk.Button(btn_frame, text="Đổi tên tự động", command=lambda: set_res("rename")).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="Ghi đè", command=lambda: set_res("overwrite")).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="Bỏ qua", command=lambda: set_res("skip")).pack(side="left", padx=5)

    dlg.protocol("WM_DELETE_WINDOW", lambda: set_res("skip"))
