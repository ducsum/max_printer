import os
import shutil
import time
from utils import get_company_code, get_long_path

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

                if is_move:
                    if os.path.exists(dest_path):
                        os.replace(long_path, dest_path)
                    else:
                        shutil.move(long_path, dest_path)
                    success_move += 1
                    duration = time.time() - start_time
                    self.app.logger.log_action("MOVE", filename, long_path, dest_path, "SUCCESS", duration, after_callback=self.app.after)
                else:
                    if os.path.exists(dest_path):
                        os.remove(dest_path)
                    shutil.copy2(long_path, dest_path)
                    
                    # Verify size
                    if os.path.getsize(long_path) != os.path.getsize(dest_path):
                        raise Exception("Kích thước file không khớp sau khi copy")
                        
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
