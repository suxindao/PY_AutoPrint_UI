import os
import sys
import time
import shutil
import win32api
import win32print
import pythoncom
import win32com.client
import logging
from datetime import datetime
import ctypes
import ctypes.wintypes
from utils.path_utils import get_app_path, ensure_directory_exists
from typing import Callable, Dict, Any


class PrinterCore:
    def __init__(self, config: dict, log_callback: Callable[[str], None] = print):
        self._is_running = True  # 新增
        self.config = config
        self.log_callback = log_callback
        self._setup_logging()

        self.source_root = config.get("source_dir")
        today_str = datetime.now().strftime("%Y-%m-%d")
        self.target_root = f"{self.source_root}_打印备份_{today_str}"

        self.DEFAULT_PRINTER = win32print.GetDefaultPrinter()
        self.MONTHLY_PRINTER_NAME = config.get("monthly_printer_name", "")
        self.DEFAULT_PAPER_SIZE = int(config.get("default_paper_size", 132))
        self.DEFAULT_PAPER_ZOOM = int(config.get("default_paper_zoom", 75))
        self.DELAY_SECONDS = float(config.get("delay_seconds", 5))
        self.ENABLE_WAIT_PROMPT = bool(config.get("enable_wait_prompt", True))
        self.WAIT_PROMPT_SLEEP = float(config.get("wait_prompt_sleep", 30))

        self._log_config()

    def _setup_logging(self):
        self.logger = logging.getLogger("PrinterCore")
        self.logger.setLevel(logging.INFO)

        # 清除所有已有的handlers
        for handler in self.logger.handlers[:]:
            self.logger.removeHandler(handler)

        # 添加控制台handler（输出到UI）
        class CallbackHandler(logging.Handler):
            def __init__(self, callback):
                super().__init__()
                self.callback = callback

            def emit(self, record):
                msg = self.format(record)
                self.callback(msg)

        self.logger.addHandler(CallbackHandler(self.log_callback))

        # 添加文件handler（保存到EXE同级目录的logs文件夹）
        log_dir = ensure_directory_exists(get_app_path("logs"))
        log_filename = datetime.now().strftime("print_log_%Y-%m-%d_%H-%M-%S.log")
        log_path = os.path.join(log_dir, log_filename)

        file_handler = logging.FileHandler(log_path, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        self.logger.addHandler(file_handler)

    def _log_config(self):
        self.logger.info("-------------------------")
        self.logger.info("⚙️ 配置信息:")
        self.logger.info(f"📂 源目录: {self.source_root}")
        self.logger.info(f"📂 保存目录: {self.target_root}")
        self.logger.info(f"🖨️ 月结单使用的打印机名称️: {self.MONTHLY_PRINTER_NAME}")
        self.logger.info(f"📄 针式打印机纸张编号: {self.DEFAULT_PAPER_SIZE}")
        self.logger.info(f"📄 针式打印机打印缩放比例: {self.DEFAULT_PAPER_ZOOM}")
        self.logger.info(f"📄 打印间隔: {self.DELAY_SECONDS}")
        self.logger.info(f"🔔 打印完目录是否弹窗并等待: {self.ENABLE_WAIT_PROMPT}")
        self.logger.info("-------------------------")

    def is_monthly_file(self, filename):
        return "月结单" in filename

    def print_pdf(self, path, use_alt=False):
        printer = self.DEFAULT_PRINTER

        self.logger.info(f"📄 打印 PDF: {path}")
        self.logger.info(f"🖨️ 打印机: {printer}")

        try:
            win32api.ShellExecute(0, "print", path, f'/d:"{printer}"', ".", 0)
            self.logger.info(f"✅ 打印成功 (PDF)")
            return True
        except Exception as e:
            self.logger.error(f"❌ 打印失败 (PDF): {e}")
            return False

    def print_excel(self, path, use_alt=False):
        printer = self.DEFAULT_PRINTER

        self.logger.info(f"")
        self.logger.info(f"📊 打印 Excel: {path}")
        self.logger.info(f"🖨️ 打印机: {printer}")

        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            wb = excel.Workbooks.Open(path, ReadOnly=True)
            for sheet in wb.Sheets:
                if use_alt:
                    sheet.PageSetup.PaperSize = 9  # A4
                    sheet.PageSetup.Zoom = False
                    sheet.PageSetup.FitToPagesWide = 1
                    sheet.PageSetup.FitToPagesTall = 1
                    sheet.PageSetup.Orientation = 1
                else:
                    try:
                        sheet.PageSetup.PaperSize = self.DEFAULT_PAPER_SIZE
                    except:
                        sheet.PageSetup.PaperSize = 9  # A4
                    sheet.PageSetup.Zoom = self.DEFAULT_PAPER_ZOOM
                    sheet.PageSetup.FitToPagesWide = False
                    sheet.PageSetup.FitToPagesTall = False
                    sheet.PageSetup.Orientation = 1

            wb.PrintOut(From=1, To=1, ActivePrinter=printer)
            self.logger.info(f"✅ 打印成功 (Excel)")
            return True
        except Exception as e:
            self.logger.error(f"❌ 打印失败 (Excel): {e}")
            return False
        finally:
            try:
                wb.Close(False)
            except:
                pass
            try:
                excel.Quit()
            except:
                pass
            pythoncom.CoUninitialize()

    def move_and_cleanup(self, src_file, src_root, target_root):
        rel_path = os.path.relpath(src_file, src_root)
        dest_file = os.path.join(target_root, rel_path)
        os.makedirs(os.path.dirname(dest_file), exist_ok=True)
        shutil.move(src_file, dest_file)
        self.logger.info(f"📁 已移动文件: {dest_file}")

        # 删除空目录
        src_dir = os.path.dirname(src_file)

        while src_dir != src_root:
            if not os.path.exists(src_dir):
                break  # 路径已不存在，不能继续
            if not any(f for f in os.listdir(src_dir) if not f.startswith("~$")):
                try:
                    os.rmdir(src_dir)
                    logging.info(f"🗑️ 删除空目录: {src_dir}")
                except Exception as e:
                    logging.warning(f"⚠️ 删除目录失败: {src_dir} - {e}")

    def show_message_box_with_timeout(self, text, caption, timeout_ms):
        MB_YESNO = 0x04
        MB_ICONQUESTION = 0x20
        IDYES = 6
        IDNO = 7

        MessageBoxTimeoutW = ctypes.windll.user32.MessageBoxTimeoutW
        MessageBoxTimeoutW.restype = ctypes.c_int
        MessageBoxTimeoutW.argtypes = [
            ctypes.wintypes.HWND,
            ctypes.wintypes.LPCWSTR,
            ctypes.wintypes.LPCWSTR,
            ctypes.wintypes.UINT,
            ctypes.wintypes.WORD,
            ctypes.wintypes.DWORD
        ]

        return MessageBoxTimeoutW(
            0,  # hWnd
            text,
            caption,
            MB_YESNO | MB_ICONQUESTION,
            0,  # Default button (0 = first button)
            timeout_ms  # Timeout in milliseconds
        )

    def run(self) -> bool:
        if not self._is_running:
            return False

        if not os.path.exists(self.source_root):
            self.logger.error(f"❌ 源目录不存在: {self.source_root}")
            return False

        for root, _, files in os.walk(self.source_root, topdown=False):

            if not self._is_running:  # 添加中断检查
                self.logger.info("🛑 打印被用户中断")
                return False

            any_printed = False

            for name in files:
                if name.startswith("~$"):
                    continue
                full_path = os.path.join(root, name)
                is_monthly = self.is_monthly_file(name)

                success = False

                if name.lower().endswith(".pdf"):
                    success = self.print_pdf(full_path, use_alt=is_monthly)
                elif name.lower().endswith((".xls", ".xlsx")):
                    success = self.print_excel(full_path, use_alt=is_monthly)

                time.sleep(self.DELAY_SECONDS)

                if success:
                    self.move_and_cleanup(full_path, self.source_root, self.target_root)
                    any_printed = True
                else:
                    return False

            if any_printed and self.ENABLE_WAIT_PROMPT:
                msg = (
                    f"📁 当前目录打印完成: \n{root}\n\n📢 将在 {self.WAIT_PROMPT_SLEEP} 秒后继续打印下一个目录..."
                    "请选择操作：\n"
                    f"【是】 = 是的，继续等待 {self.WAIT_PROMPT_SLEEP} 秒\n"
                    "【否】 = 继续打印"
                )
                self.logger.info(f"📁 当前目录打印完成: {root}")
                self.logger.info(f"📢 将在 {self.WAIT_PROMPT_SLEEP} 秒后继续打印下一个目录...")

                response = self.show_message_box_with_timeout(
                    msg,
                    "📢 打印完成提示",
                    int(self.WAIT_PROMPT_SLEEP * 1000)
                )

                if response == 6:  # IDYES
                    self.logger.info(f"✅ 用户选择等待，等待 {self.WAIT_PROMPT_SLEEP} 秒...")
                    time.sleep(self.WAIT_PROMPT_SLEEP)
                else:
                    self.logger.info("⏩ 用户选择跳过等待")

        # self.logger.info("✅ 所有文件打印完成")
        return True
