import os
import csv
import time
import datetime
import logging
from typing import Optional

class AppLogger:
    def __init__(self, log_box=None, log_dir: str = "print_logs"):
        self.log_box = log_box
        self.log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), log_dir)
        os.makedirs(self.log_dir, exist_ok=True)
        
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        self.txt_path = os.path.join(self.log_dir, f"log_{timestamp}.txt")
        self.csv_path = os.path.join(self.log_dir, f"log_{timestamp}.csv")
        
        # Initialize standard logger
        self._logger = logging.getLogger("MassPrintApp")
        self._logger.setLevel(logging.DEBUG)
        
        # File handler
        file_handler = logging.FileHandler(self.txt_path, encoding="utf-8")
        file_formatter = logging.Formatter('[%(asctime)s] %(levelname)s\n%(message)s\n', datefmt='%H:%M:%S')
        file_handler.setFormatter(file_formatter)
        self._logger.addHandler(file_handler)
        
        # Initialize CSV handle
        self._csv_file = None
        self._csv_writer = None
        try:
            self._csv_file = open(self.csv_path, 'w', newline='', encoding='utf-8-sig')
            self._csv_writer = csv.writer(self._csv_file)
            self._csv_writer.writerow(["Timestamp", "Action", "Filename", "Source", "Destination", "Status", "Duration", "Error"])
            self._csv_file.flush()
        except Exception as e:
            self._logger.error(f"Failed to create CSV log file: {e}")

    def __del__(self):
        if hasattr(self, '_csv_file') and self._csv_file:
            try:
                self._csv_file.close()
            except Exception:
                pass

    def log(self, msg: str, after_callback=None, level=logging.INFO):
        """Ghi log chung (chỉ ghi txt và hiển thị UI)."""
        if level >= logging.ERROR:
            self._logger.error(msg, exc_info=True)
        elif level >= logging.WARNING:
            self._logger.warning(msg)
        else:
            self._logger.info(msg)
            
        timestamped = f"[{time.strftime('%H:%M:%S')}]\n{msg}"
        
        if self.log_box and after_callback:
            def _update_ui():
                try:
                    self.log_box.configure(state="normal")
                    self.log_box.insert("end", timestamped + "\n\n")
                    self.log_box.see("end")
                    self.log_box.configure(state="disabled")
                except Exception as e:
                    self._logger.error(f"Failed to update UI log: {e}")
            after_callback(0, _update_ui)

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
        # Use log for TXT and UI
        level = logging.ERROR if status == "FAILED" or error else logging.INFO
        self.log(txt_msg, after_callback, level=level)
        
        # Log to CSV (tái sử dụng handle)
        if self._csv_writer and self._csv_file:
            try:
                self._csv_writer.writerow([timestamp, action, filename, source, destination, status, f"{duration:.2f}", error])
                self._csv_file.flush() # Vẫn an toàn dữ liệu, nhưng bỏ qua overhead HĐH open/close
            except Exception as e:
                self._logger.error(f"Failed to write to CSV log: {e}")
