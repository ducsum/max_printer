import os
import time
import logging
from utils import ALL_EXTENSIONS
from constants import IGNORED_DIRS

logger = logging.getLogger(__name__)

class FolderScanner:
    def __init__(self, folder_path: str, progress_callback, finished_callback):
        self.folder_path = folder_path
        self.progress_callback = progress_callback
        self.finished_callback = finished_callback

    def scan(self):
        scanned_files = []
        file_count = 0
        total_size = 0
        
        def walk_error(err):
            logger.warning(f"Lỗi truy cập thư mục khi quét: {err}")

        for root, dirs, files in os.walk(self.folder_path, onerror=walk_error):
            # Lọc bỏ các thư mục không cần thiết để tối ưu
            dirs[:] = [d for d in dirs if d not in IGNORED_DIRS]
            
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
                        size = os.path.getsize(full)
                    except PermissionError as e:
                        logger.warning(f"Không có quyền truy cập file: {full} - {e}")
                        continue
                    except OSError as e:
                        logger.warning(f"Lỗi đọc file info: {full} - {e}")
                        continue
                        
                    scanned_files.append({
                        "name": f, 
                        "path": full, 
                        "ext": ext, 
                        "mtime": mtime,
                        "size": size
                    })
                    
                    file_count += 1
                    total_size += size
                    
                    if file_count % 50 == 0:
                        self.progress_callback(file_count)

        self.finished_callback(scanned_files, total_size)
