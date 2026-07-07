import os
import time
import logging
from utils import ALL_EXTENSIONS
from constants import IGNORED_DIRS

logger = logging.getLogger(__name__)

class FolderScanner:
    def __init__(self, folder_path: str, progress_callback, finished_callback, cancel_event):
        self.folder_path = folder_path
        self.progress_callback = progress_callback
        self.finished_callback = finished_callback
        self.cancel_event = cancel_event

    def scan(self):
        scanned_files = []
        file_count = 0
        folder_count = 0
        total_size = 0
        
        start_time = time.time()
        last_callback_time = start_time
        
        def walk_error(err):
            logger.warning(f"Lỗi truy cập thư mục khi quét: {err}")

        for root, dirs, files in os.walk(self.folder_path, onerror=walk_error):
            if self.cancel_event.is_set():
                break
                
            folder_count += 1
            # Lọc bỏ các thư mục không cần thiết để tối ưu
            dirs[:] = [d for d in dirs if d not in IGNORED_DIRS]
            
            for f in files:
                if self.cancel_event.is_set():
                    break
                    
                if f.startswith("~$"):
                    continue
                    
                name_lower = f.lower()
                ext = os.path.splitext(f)[1].lower()
                if ext in ALL_EXTENSIONS:
                    full = os.path.join(root, f)
                    try:
                        # Tối ưu: Dùng os.stat 1 lần để lấy cả mtime và size
                        stat = os.stat(full)
                        mtime = time.strftime(
                            "%d/%m/%Y %H:%M", time.localtime(stat.st_mtime)
                        )
                        size = stat.st_size
                    except (PermissionError, OSError) as e:
                        logger.warning(f"Lỗi truy cập file info: {full} - {e}")
                        continue
                        
                    scanned_files.append({
                        "name": f,
                        "name_lower": name_lower,
                        "path": full, 
                        "ext": ext, 
                        "mtime": mtime,
                        "size": size
                    })
                    
                    file_count += 1
                    total_size += size
                    
                    current_time = time.time()
                    if current_time - last_callback_time >= 0.1: # Chỉ cập nhật UI 10 lần/s để không làm chậm scan
                        elapsed = current_time - start_time
                        speed = file_count / elapsed if elapsed > 0 else 0
                        self.progress_callback(file_count, folder_count, elapsed, speed)
                        last_callback_time = current_time

        # Final update
        elapsed = time.time() - start_time
        speed = file_count / elapsed if elapsed > 0 else 0
        self.progress_callback(file_count, folder_count, elapsed, speed)

        self.finished_callback(scanned_files, total_size)
