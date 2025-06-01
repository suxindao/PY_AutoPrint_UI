from typing import Dict, Any, Optional
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
                             QLabel, QLineEdit, QSpinBox, QDoubleSpinBox, QCheckBox,
                             QPushButton, QTextEdit, QFileDialog, QMessageBox, QSizePolicy)
from PyQt5.QtCore import Qt, pyqtSignal, QThread

from utils.config_manager import ConfigManager
from printer_core import PrinterCore
from utils.path_utils import get_app_path, ensure_directory_exists
from PyQt5.QtGui import QFont


class PrinterThread(QThread):
    finished = pyqtSignal(bool)
    log_message = pyqtSignal(str)

    def __init__(self, config: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.config = config
        self._is_running = True  # 控制线程运行的标志

    def run(self):
        try:
            printer = PrinterCore(self.config, self.log_message.emit)
            while self._is_running:  # 添加循环检查
                if not printer.run():  # 修改run方法使其可中断
                    break
            self.finished.emit(True)
        except Exception as e:
            self.log_message.emit(f"❌ 打印线程异常: {str(e)}")
            self.finished.emit(False)

    def stop(self):
        self._is_running = False  # 设置标志位停止线程
        self.quit()  # 确保线程退出


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # 确保日志目录存在
        self.log_dir = ensure_directory_exists(get_app_path("logs"))

        self.config_manager = ConfigManager()
        self.printer_thread: Optional[PrinterThread] = None

        # 设置大字体
        self.setFont(QFont("Microsoft YaHei", 12))  # 设置默认字体
        self.init_ui()

        # 默认全屏显示
        # self.showMaximized()

        self.load_config()

    def init_ui(self):
        self.setWindowTitle("诊所出货单月结单自动打印工具")
        self.setGeometry(100, 100, 800, 600)

        # 主布局
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        # 配置区域
        config_group = QGroupBox("打印配置")
        config_layout = QVBoxLayout()

        # 源目录
        source_layout = QHBoxLayout()
        source_layout.addWidget(QLabel("源目录:"))
        self.source_edit = QLineEdit()
        self.source_edit.setPlaceholderText("请选择包含PDF/Excel文件的目录")
        source_layout.addWidget(self.source_edit)
        self.source_btn = QPushButton("浏览...")
        self.source_btn.clicked.connect(self.select_source_dir)
        source_layout.addWidget(self.source_btn)
        config_layout.addLayout(source_layout)

        # 打印机设置
        printer_layout = QHBoxLayout()
        printer_layout.addWidget(QLabel("月结单打印机:"))
        self.printer_edit = QLineEdit()
        self.printer_edit.setPlaceholderText("留空则使用默认打印机")
        printer_layout.addWidget(self.printer_edit)
        config_layout.addLayout(printer_layout)

        # 纸张设置
        paper_layout = QHBoxLayout()
        paper_layout.addWidget(QLabel("针式打印机纸张编号:"))
        self.paper_size_spin = QSpinBox()
        self.paper_size_spin.setRange(1, 500)
        paper_layout.addWidget(self.paper_size_spin)

        paper_layout.addWidget(QLabel("缩放比例(%):"))
        self.paper_zoom_spin = QSpinBox()
        self.paper_zoom_spin.setRange(10, 200)
        paper_layout.addWidget(self.paper_zoom_spin)
        config_layout.addLayout(paper_layout)

        # 其他设置
        other_layout = QHBoxLayout()
        other_layout.addWidget(QLabel("打印间隔(秒):"))
        self.delay_spin = QDoubleSpinBox()
        self.delay_spin.setRange(0.1, 60)
        self.delay_spin.setSingleStep(0.5)
        other_layout.addWidget(self.delay_spin)

        self.wait_prompt_check = QCheckBox("打印完成弹窗提示")
        other_layout.addWidget(self.wait_prompt_check)

        other_layout.addWidget(QLabel("等待时间(秒):"))
        self.wait_sleep_spin = QDoubleSpinBox()
        self.wait_sleep_spin.setRange(1, 300)
        other_layout.addWidget(self.wait_sleep_spin)
        config_layout.addLayout(other_layout)

        config_group.setLayout(config_layout)
        main_layout.addWidget(config_group)

        # 日志区域
        log_group = QGroupBox("打印日志")
        log_layout = QVBoxLayout()
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        log_layout.addWidget(self.log_edit)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)

        # 控制按钮
        button_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始打印")
        self.start_btn.clicked.connect(self.start_printing)
        button_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("停止打印")
        self.stop_btn.clicked.connect(self.stop_printing)
        self.stop_btn.setEnabled(False)
        button_layout.addWidget(self.stop_btn)

        self.save_btn = QPushButton("保存配置")
        self.save_btn.clicked.connect(self.save_config)
        button_layout.addWidget(self.save_btn)
        main_layout.addLayout(button_layout)

        # 设置布局间距和对齐
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # 设置控件大小策略
        for widget in [self.source_edit, self.printer_edit,
                       self.paper_size_spin, self.paper_zoom_spin,
                       self.delay_spin, self.wait_sleep_spin]:
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

    def load_config(self):
        config = self.config_manager.get_all()
        self.source_edit.setText(config.get("source_dir", ""))
        self.printer_edit.setText(config.get("monthly_printer_name", ""))
        self.paper_size_spin.setValue(config.get("default_paper_size", 132))
        self.paper_zoom_spin.setValue(config.get("default_paper_zoom", 75))
        self.delay_spin.setValue(config.get("delay_seconds", 5))
        self.wait_prompt_check.setChecked(config.get("enable_wait_prompt", True))
        self.wait_sleep_spin.setValue(config.get("wait_prompt_sleep", 30))

    def save_config(self):
        if not self.source_edit.text():
            QMessageBox.warning(self, "警告", "请先选择源目录!")
            return

        config = {
            "source_dir": self.source_edit.text(),
            "monthly_printer_name": self.printer_edit.text(),
            "default_paper_size": self.paper_size_spin.value(),
            "default_paper_zoom": self.paper_zoom_spin.value(),
            "delay_seconds": self.delay_spin.value(),
            "enable_wait_prompt": self.wait_prompt_check.isChecked(),
            "wait_prompt_sleep": self.wait_sleep_spin.value()
        }
        self.config_manager.update_config(config)
        if self.config_manager.save_config():
            QMessageBox.information(self, "成功", "配置已保存!")
        else:
            QMessageBox.warning(self, "错误", "保存配置失败!")

    def select_source_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择源目录")
        if dir_path:
            self.source_edit.setText(dir_path)

    def log_message(self, message: str):
        self.log_edit.append(message)

    def start_printing(self):
        if not self.source_edit.text():
            QMessageBox.warning(self, "警告", "请先选择源目录!")
            return

        if self.printer_thread and self.printer_thread.isRunning():
            return

        self.save_config()  # 开始前自动保存配置

        config = self.config_manager.get_all()
        self.printer_thread = PrinterThread(config, self)
        self.printer_thread.log_message.connect(self.log_message)
        self.printer_thread.finished.connect(self.printing_finished)

        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.log_edit.clear()

        self.printer_thread.start()

    def stop_printing(self):
        if self.printer_thread and self.printer_thread.isRunning():
            self.printer_thread.stop()  # 调用停止方法
            self.log_message("🛑 正在停止打印...")
            if not self.printer_thread.wait(2000):  # 等待2秒线程结束
                self.printer_thread.terminate()  # 强制终止
                self.log_message("⚠️ 打印线程已强制终止")

        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

    def printing_finished(self, success: bool):
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

        if success:
            self.log_message("✅ 所有文件打印完成")
        else:
            self.log_message("❌ 打印过程中出现错误")

    def closeEvent(self, event):
        if self.printer_thread and self.printer_thread.isRunning():
            reply = QMessageBox.question(
                self, '确认退出',
                "打印任务仍在运行，确定要退出吗?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )

            if reply == QMessageBox.Yes:
                self.stop_printing()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()
