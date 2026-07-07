import os
import shutil
import time
import hashlib
from utils import get_company_code, get_long_path, is_file_locked
from constants import COPY_RETRY_COUNT, COPY_RETRY_DELAY, HASH_LIMIT_BYTES

class CopyManager:
    def __init__(self, app):
        self.app = app

    def _get_auto_renamed_path(self, target_dir: str, filename: str) -> str:
        base, ext = os.path.splitext(filename)
        counter = 1
        while True:
            new_name = f"{base} ({counter}){ext}"
            new_path = os.path.join(target_dir, new_name)
            if not os.path.exists(new_path):
                return new_path
            counter += 1

    def _calculate_sha256(self, filepath: str) -> str:
        sha256_hash = hashlib.sha256()
        with open(filepath, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()

    def run(self, files: list, dest_dir: str, is_move: bool):
        total = len(files)
        action_name = "Di chuyển" if is_move else "Sao chép"
        
        self.app.logger.log(f"--- Bắt đầu {action_name.lower()} {total} file ---", after_callback=self.app.after)

        success_copy = 0
        success_move = 0
        skipped = 0
        failed = 0
        processed = 0
        cancelled = False
        
        global_conflict_action = None

        for path in files:
            if self.app.copy_cancel_requested.is_set():
                cancelled = True
                self.app.logger.log(f"--- Đã dừng theo yêu cầu. Đã {action_name.lower()} {processed}/{total} file trước khi dừng. ---", after_callback=self.app.after)
                break

            start_time = time.time()
            filename = os.path.basename(path)
            company_code = get_company_code(filename)
            
            target_dir = os.path.join(dest_dir, company_code)
            target_dir = get_long_path(target_dir)
            dest_path = os.path.join(target_dir, filename)
            
            long_path = get_long_path(path)

            if long_path == dest_path:
                skipped += 1
                self.app.logger.log_action("SKIP", filename, long_path, dest_path, "SKIPPED_SAME_PATH", after_callback=self.app.after)
                processed += 1
                self.app.after(0, self.app._update_progress, processed, total)
                continue

            try:
                os.makedirs(target_dir, exist_ok=True)
                
                if os.path.exists(dest_path):
                    if global_conflict_action:
                        action = global_conflict_action
                    else:
                        action, apply_all = self.app.ask_conflict(filename)
                        if apply_all:
                            global_conflict_action = action
                    
                    if action == "skip":
                        skipped += 1
                        self.app.logger.log_action("SKIP", filename, long_path, dest_path, "SKIPPED", after_callback=self.app.after)
                        processed += 1
                        self.app.after(0, self.app._update_progress, processed, total)
                        continue
                    elif action == "rename":
                        dest_path = self._get_auto_renamed_path(target_dir, filename)
                    elif action == "overwrite":
                        pass

                # Retry logic for lock/permission errors
                success = False
                last_error = None
                for attempt in range(COPY_RETRY_COUNT):
                    try:
                        if is_file_locked(long_path):
                            raise PermissionError(f"File nguồn đang bị khóa: {long_path}")
                        if os.path.exists(dest_path) and is_file_locked(dest_path):
                            raise PermissionError(f"File đích đang bị khóa: {dest_path}")

                        if is_move:
                            if os.path.exists(dest_path):
                                os.replace(long_path, dest_path)
                            else:
                                shutil.move(long_path, dest_path)
                        else:
                            if os.path.exists(dest_path):
                                os.remove(dest_path)
                            shutil.copy2(long_path, dest_path)
                            
                            # Verify
                            src_size = os.path.getsize(long_path)
                            dst_size = os.path.getsize(dest_path)
                            if src_size != dst_size:
                                raise Exception("Kích thước file không khớp sau khi copy")
                                
                            if src_size < HASH_LIMIT_BYTES:
                                if self._calculate_sha256(long_path) != self._calculate_sha256(dest_path):
                                    raise Exception("Mã SHA256 không khớp sau khi copy")
                        success = True
                        break
                    except PermissionError as e:
                        last_error = e
                        time.sleep(COPY_RETRY_DELAY)
                    except Exception as e:
                        last_error = e
                        break # Other errors, don't retry

                if not success:
                    raise last_error

                if is_move:
                    success_move += 1
                    duration = time.time() - start_time
                    self.app.logger.log_action("MOVE", filename, long_path, dest_path, "SUCCESS", duration, after_callback=self.app.after)
                else:
                    success_copy += 1
                    duration = time.time() - start_time
                    self.app.logger.log_action("COPY", filename, long_path, dest_path, "SUCCESS", duration, after_callback=self.app.after)

            except Exception as e:
                failed += 1
                duration = time.time() - start_time
                self.app.logger.log_action("ERROR", filename, long_path, dest_path, "FAILED", duration, error=str(e), after_callback=self.app.after)

            processed += 1
            self.app.after(0, self.app._update_progress, processed, total)

        if not cancelled:
            self.app.logger.log(f"--- Hoàn tất {action_name.lower()} ---", after_callback=self.app.after)

        self.app.after(0, self.app._on_copy_finished, success_copy, success_move, skipped, failed, is_move, dest_dir)
