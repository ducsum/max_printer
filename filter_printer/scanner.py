import os
import sys
import time
import logging
from utils import ALL_EXTENSIONS

class FileRecord:
    __slots__ = ('name', 'name_lower', 'path', 'ext', 'mtime', 'size')
    def __init__(self, name, name_lower, path, ext, mtime, size):
        self.name = name
        self.name_lower = name_lower
        self.path = path
        self.ext = ext
        self.mtime = mtime
        self.size = size
    def __getitem__(self, key):
        return getattr(self, key)
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
        
        # Chuyển đổi đệ quy (Recursion) sang Lặp (Iterative DFS) bằng Stack để tránh lỗi RecursionError
        stack = [self.folder_path]
        
        while stack:
            if self.cancel_event.is_set():
                break
            
            current_path = stack.pop()
            folder_count += 1
            
            try:
                with os.scandir(current_path) as it:
                    for entry in it:
                        if self.cancel_event.is_set():
                            break
                        if entry.is_dir(follow_symlinks=False):
                            if entry.name not in IGNORED_DIRS:
                                stack.append(entry.path)
                        elif entry.is_file(follow_symlinks=False):
                            f = entry.name
                            if f.startswith("~$"):
                                continue
                            
                            name_lower = f.lower()
                            ext = sys.intern(os.path.splitext(f)[1].lower())
                            if ext in ALL_EXTENSIONS:
                                try:
                                    stat = entry.stat()
                                    mtime = time.strftime("%d/%m/%Y %H:%M", time.localtime(stat.st_mtime))
                                    size = stat.st_size
                                except (PermissionError, OSError) as e:
                                    logger.warning(f"Lỗi truy cập file info: {entry.path} - {e}")
                                    continue
                                
                                # Sử dụng class FileRecord (với __slots__) thay vì dict {} để tiết kiệm 50% RAM
                                scanned_files.append(FileRecord(
                                    name=f,
                                    name_lower=name_lower,
                                    path=entry.path, 
                                    ext=ext, 
                                    mtime=mtime,
                                    size=size
                                ))
                                file_count += 1
                                total_size += size
                                
                                current_time = time.time()
                                if current_time - last_callback_time >= 0.1:
                                    elapsed = current_time - start_time
                                    speed = file_count / elapsed if elapsed > 0 else 0
                                    self.progress_callback(file_count, folder_count, elapsed, speed)
                                    last_callback_time = current_time
            except PermissionError as e:
                logger.warning(f"Lỗi truy cập thư mục khi quét: {e}")
            except OSError as e:
                logger.warning(f"OS Error quét: {e}")

        # Final update
        elapsed = time.time() - start_time
        speed = file_count / elapsed if elapsed > 0 else 0
        self.progress_callback(file_count, folder_count, elapsed, speed)

        self.finished_callback(scanned_files, total_size)
