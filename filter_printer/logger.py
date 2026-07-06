import os
import csv
import time
import datetime
from typing import Optional

class AppLogger:
    def __init__(self, log_box=None, log_dir: str = "print_logs"):
        self.log_box = log_box
        self.log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), log_dir)
        os.makedirs(self.log_dir, exist_ok=True)
        
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        self.txt_path = os.path.join(self.log_dir, f"log_{timestamp}.txt")
        self.csv_path = os.path.join(self.log_dir, f"log_{timestamp}.csv")
        
        # Initialize CSV header
        try:
            with open(self.csv_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(["Timestamp", "Action", "Filename", "Source", "Destination", "Status", "Duration", "Error"])
        except Exception:
            pass

    def log(self, msg: str, after_callback=None):
        """Ghi log chung (chỉ ghi txt và hiển thị UI)."""
        timestamped = f"[{time.strftime('%H:%M:%S')}]\n{msg}"
        
        if self.log_box and after_callback:
            def _update_ui():
                try:
                    self.log_box.configure(state="normal")
                    self.log_box.insert("end", timestamped + "\n\n")
                    self.log_box.see("end")
                    self.log_box.configure(state="disabled")
                except Exception:
                    pass
            after_callback(0, _update_ui)
            
        try:
            with open(self.txt_path, "a", encoding="utf-8") as f:
                f.write(timestamped + "\n\n")
        except Exception:
            pass

    def log_action(self, action: str, filename: str, source: str, destination: str, 
                   status: str, duration: float = 0.0, error: str = "", after_callback=None):
        """Ghi log chi tiết cho 1 file, bao gồm cả CSV và TXT."""
        timestamp = time.strftime('%H:%M:%S')
        
        # Format text message
        msg_parts = [f"[{status}] {filename}"]
        if source:
            msg_parts.append(f"FROM: {source}")
        if destination:
            msg_parts.append(f"TO: {destination}")
        if duration > 0:
            msg_parts.append(f"Thời gian: {duration:.2f}s")
        if error:
            msg_parts.append(f"Lỗi: {error}")
            
        txt_msg = "\n".join(msg_parts)
        self.log(txt_msg, after_callback)
        
        # Log to CSV
        try:
            with open(self.csv_path, 'a', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow([timestamp, action, filename, source, destination, status, f"{duration:.2f}", error])
        except Exception:
            pass
