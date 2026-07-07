import os
import re
import functools

FILE_TYPE_GROUPS = {
    "PDF": [".pdf"],
    "DOC": [".doc", ".docx"],
    "XLS": [".xls", ".xlsx"],
}
ALL_EXTENSIONS = [ext for exts in FILE_TYPE_GROUPS.values() for ext in exts]

_NATURAL_SPLIT = re.compile(r'(\d+)')

@functools.lru_cache(maxsize=10000)
def natural_sort_key(text: str) -> tuple:
    """Tạo key sắp xếp để chuỗi có số được sắp xếp theo đúng thứ tự logic. Caching tiết kiệm CPU."""
    return tuple(int(c) if c.isdigit() else c.lower() for c in _NATURAL_SPLIT.split(text))

def format_size(size_in_bytes: int) -> str:
    """Định dạng kích thước file sang đơn vị dễ đọc."""
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size_in_bytes < 1024.0:
            return f"{size_in_bytes:.2f} {unit}"
        size_in_bytes /= 1024.0
    return f"{size_in_bytes:.2f} PB"

def format_time(seconds: float) -> str:
    """Định dạng số giây thành chuỗi thời gian HH:MM:SS."""
    m, s = divmod(int(seconds), 60)
    h, m = divmod(m, 60)
    if h > 0:
        return f"{h:02d}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"

def is_file_locked(filepath: str) -> bool:
    """Kiểm tra xem file có đang bị khóa bởi chương trình khác không."""
    if not os.path.exists(filepath):
        return False
    try:
        # Mở file để append. Nếu bị khóa sẽ văng lỗi PermissionError (Windows).
        with open(filepath, 'a'):
            pass
        return False
    except IOError:
        return True

_COMPANY_CODE_RE = re.compile(r'([a-zA-Z]+\d+)')

def get_company_code(filename: str) -> str:
    """Trích xuất mã công ty từ tên file. Trả về 'Khác' nếu không tìm thấy."""
    match = _COMPANY_CODE_RE.search(filename)
    if match:
        return match.group(1).upper()
    return "Khác"

def get_long_path(path: str) -> str:
    """Thêm prefix \\\\?\\ cho Windows để hỗ trợ đường dẫn siêu dài (quá 260 ký tự)."""
    if os.name == 'nt' and not path.startswith(r"\\?\\"):
        return r"\\?\\" + os.path.abspath(path)
    return os.path.abspath(path)
