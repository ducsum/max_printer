import os
import time
import subprocess
from typing import Optional
from constants import PRINT_RETRY_COUNT, PRINT_RETRY_DELAY, SUMATRA_TIMEOUT

def find_sumatra() -> Optional[str]:
    SUMATRA_CANDIDATES = [
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
        os.path.expanduser(r"~\AppData\Local\SumatraPDF\SumatraPDF.exe"),
    ]
    for path in SUMATRA_CANDIDATES:
        if os.path.isfile(path):
            return path
    return None

class Printer:
    """Quản lý việc in Word / Excel / PDF, tái sử dụng 1 instance Word/Excel cho toàn bộ batch."""

    def __init__(self, logger, printer_name=None):
        self.logger = logger
        self._word_app = None
        self._excel_app = None
        self.sumatra_path = find_sumatra()
        self.printer_name = printer_name or None
        self._original_default_printer = None

    def _get_word(self):
        if self._word_app is None:
            try:
                import win32com.client
                self._word_app = win32com.client.DispatchEx("Word.Application")
                self._word_app.Visible = False
                self._word_app.DisplayAlerts = 0
                self._word_app.AutomationSecurity = 3
            except Exception as e:
                self.logger.log(f"[CẢNH BÁO] Không khởi tạo được Word: {e}", after_callback=None)
        return self._word_app

    def _get_excel(self):
        if self._excel_app is None:
            try:
                import win32com.client
                self._excel_app = win32com.client.DispatchEx("Excel.Application")
                self._excel_app.Visible = False
                self._excel_app.DisplayAlerts = False
            except Exception as e:
                self.logger.log(f"[CẢNH BÁO] Không khởi tạo được Excel: {e}", after_callback=None)
        return self._excel_app

    def print_word(self, filepath: str):
        word = self._get_word()
        if not word:
            raise Exception("Word không khả dụng")
        
        if self.printer_name:
            try:
                word.ActivePrinter = self.printer_name
            except Exception as e:
                self.logger.log(f"[CẢNH BÁO] Không đặt được máy in cho Word: {e}", after_callback=None)
        
        doc = None
        try:
            # Prevent password prompt hanging
            doc = word.Documents.Open(filepath, ReadOnly=True, PasswordDocument="DUMMY_PASSWORD_123")
            # In đồng bộ (Background=False) để hệ thống đợi lệnh nạp vào Spooler thay vì sleep thủ công
            doc.PrintOut(Background=False)
        finally:
            if doc:
                try:
                    doc.Close(False)
                except Exception as e:
                    self.logger.log(f"[CẢNH BÁO] Lỗi khi đóng file Word: {e}", after_callback=None)

    def print_excel(self, filepath: str):
        excel = self._get_excel()
        if not excel:
            raise Exception("Excel không khả dụng")
            
        if self.printer_name:
            try:
                excel.ActivePrinter = self.printer_name
            except Exception as e:
                self.logger.log(f"[CẢNH BÁO] Không đặt được máy in cho Excel: {e}", after_callback=None)
                
        wb = None
        try:
            wb = excel.Workbooks.Open(filepath, ReadOnly=True, UpdateLinks=0, IgnoreReadOnlyRecommended=True)
            # Excel PrintOut mặc định là đồng bộ
            wb.PrintOut()
        finally:
            if wb:
                try:
                    wb.Close(False)
                except Exception as e:
                    self.logger.log(f"[CẢNH BÁO] Lỗi khi đóng file Excel: {e}", after_callback=None)

    def print_pdf(self, filepath: str):
        if self.sumatra_path:
            cmd = [self.sumatra_path, "-print-to", self.printer_name] if self.printer_name else [self.sumatra_path, "-print-to-default"]
            cmd.extend(["-silent", filepath])
            
            try:
                subprocess.run(cmd, timeout=SUMATRA_TIMEOUT)
            except subprocess.TimeoutExpired:
                raise Exception("Timeout in SumatraPDF")
        else:
            self._switch_default_printer_if_needed()
            os.startfile(filepath, "print")
            time.sleep(3)

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
            self.logger.log(f"[CẢNH BÁO] Không đổi được máy in mặc định tạm thời: {e}", after_callback=None)

    def restore_default_printer(self):
        if self._original_default_printer:
            try:
                import win32print
                win32print.SetDefaultPrinter(self._original_default_printer)
            except Exception:
                pass

    def _reset_word(self):
        try:
            if self._word_app is not None:
                self._word_app.Quit()
        except Exception:
            pass
        finally:
            self._word_app = None

    def _reset_excel(self):
        try:
            if self._excel_app is not None:
                self._excel_app.Quit()
        except Exception:
            pass
        finally:
            self._excel_app = None

    def print_file(self, filepath: str) -> tuple[bool, float]:
        """Trả về (Thành công?, Thời gian xử lý)"""
        start_time = time.time()
        
        if not os.path.exists(filepath):
            return False, 0.0

        ext = os.path.splitext(filepath)[1].lower()

        for attempt in range(PRINT_RETRY_COUNT):
            try:
                if ext == ".pdf":
                    self.print_pdf(filepath)
                elif ext in (".doc", ".docx"):
                    self.print_word(filepath)
                elif ext in (".xls", ".xlsx"):
                    self.print_excel(filepath)
                else:
                    return False, 0.0
                
                return True, time.time() - start_time
                
            except Exception as e:
                if attempt < PRINT_RETRY_COUNT - 1:
                    if ext in (".doc", ".docx"):
                        self._reset_word()
                    elif ext in (".xls", ".xlsx"):
                        self._reset_excel()
                    time.sleep(PRINT_RETRY_DELAY)
                else:
                    raise e
                    
        return False, time.time() - start_time

    def cleanup(self):
        self._reset_word()
        self._reset_excel()
        self.restore_default_printer()
