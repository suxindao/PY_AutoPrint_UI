from typing import Dict, Any, Optional
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
                             QLabel, QLineEdit, QSpinBox, QDoubleSpinBox, QCheckBox,
                             QPushButton, QTextEdit, QFileDialog, QMessageBox, QSizePolicy, QComboBox)
from PyQt5.QtCore import Qt, pyqtSignal, QThread

from utils.config_manager import ConfigManager
from printer_core import PrinterCore
from utils.path_utils import get_app_path, ensure_directory_exists
from PyQt5.QtGui import QFont
import win32print


class PrinterThread(QThread):
    finished = pyqtSignal(bool)
    log_message = pyqtSignal(str)

    def __init__(self, config: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.config = config
        self._is_running = True  # 控制线程运行的标志
        self._parent = parent

    def run(self):
        try:
            printer = PrinterCore(self.config, self._parent, self.log_message.emit)
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

        # 存储打印机列表
        self.available_printers = []

        # 存储纸张编号和名称的映射
        self.paper_sizes = {}

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

        # 纸张设置
        paper_layout = QHBoxLayout()
        paper_layout.addWidget(QLabel("针式打印机纸张编号:"))
        self.paper_size_spin = QSpinBox()
        self.paper_size_spin.setRange(1, 500)
        self.paper_size_spin.setDisabled(True)
        paper_layout.addWidget(self.paper_size_spin)

        # paper_layout.addWidget(QLabel("缩放比例(%):"))
        # self.paper_zoom_spin = QSpinBox()
        # self.paper_zoom_spin.setRange(10, 200)
        # paper_layout.addWidget(self.paper_zoom_spin)
        # config_layout.addLayout(paper_layout)

        # 其他设置
        other_layout = QHBoxLayout()
        other_layout.addWidget(QLabel("打印间隔(秒):"))
        self.delay_spin = QDoubleSpinBox()
        self.delay_spin.setRange(0.01, 100)
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
        self.save_btn.clicked.connect(lambda: self.save_config(showAlert=True))
        button_layout.addWidget(self.save_btn)

        # 添加纸张信息按钮
        self.paper_info_btn = QPushButton("纸张信息")
        self.paper_info_btn.clicked.connect(self.show_printer_info)
        button_layout.addWidget(self.paper_info_btn)

        # 添加"打开打印机设置"按钮
        self.printer_setup_btn = QPushButton("系统打印机设置")
        self.printer_setup_btn.clicked.connect(self.open_printer_settings)
        self.printer_setup_btn.setToolTip("打开系统打印机设置窗口")  # 鼠标悬停提示
        button_layout.addWidget(self.printer_setup_btn)

        main_layout.addLayout(button_layout)

        # 在打印机设置组中添加打印机选择下拉框
        # printer_group = QGroupBox("打印机设置")
        printer_group = QGroupBox("")
        printer_layout = QVBoxLayout()  # 改为垂直布局

        # 打印机选择标签
        printer_label = QLabel("设置默认打印机:")

        # 打印机下拉选择框
        self.printer_combo = QComboBox()
        self.printer_combo.setMinimumWidth(300)

        # 第二行：操作按钮
        printer_btn_layout = QHBoxLayout()

        # 刷新打印机列表按钮
        refresh_btn = QPushButton("刷新列表")
        refresh_btn.clicked.connect(self.refresh_printer_list)

        # 新增"设为默认"按钮
        set_default_btn = QPushButton("设为默认")
        set_default_btn.clicked.connect(self.set_default_printer)
        set_default_btn.setToolTip("将选中的打印机设置为系统默认打印机")

        # 将控件添加到布局
        printer_btn_layout.addWidget(refresh_btn)
        printer_btn_layout.addWidget(set_default_btn)
        # printer_btn_layout.addStretch()

        printer_layout.addWidget(printer_label)
        printer_layout.addWidget(self.printer_combo)
        printer_layout.addLayout(printer_btn_layout)
        printer_group.setLayout(printer_layout)

        # 替换原有的打印机设置组
        config_layout.insertWidget(1, printer_group)  # 放在源目录设置下面

        # 修改纸张设置部分
        # paper_group = QGroupBox("纸张设置")
        paper_group = QGroupBox("")
        paper_layout = QHBoxLayout()

        # 纸张选择下拉框
        paper_label = QLabel("纸张类型:")

        self.paper_combo = QComboBox()
        self.paper_combo.setMinimumWidth(250)
        self.paper_combo.currentIndexChanged.connect(self.on_paper_selected)

        # 缩放比例设置保持不变
        zoom_label = QLabel("缩放比例(%):")
        self.paper_zoom_spin = QSpinBox()
        self.paper_zoom_spin.setRange(10, 200)

        paper_layout.addWidget(paper_label)
        paper_layout.addWidget(self.paper_combo)
        # paper_layout.addStretch(1)
        paper_layout.addWidget(zoom_label)
        paper_layout.addWidget(self.paper_zoom_spin)
        paper_group.setLayout(paper_layout)

        # 月结单打印机选择
        monthly_printer_layout = QHBoxLayout()
        monthly_printer_label = QLabel("月结单打印机:")
        self.monthly_printer_combo = QComboBox()
        monthly_printer_layout.addWidget(monthly_printer_label)
        monthly_printer_layout.addWidget(self.monthly_printer_combo)
        config_layout.addLayout(monthly_printer_layout)

        # 在打印机列表刷新时同时加载纸张信息
        self.refresh_printer_list()

        # 替换原有的打印机设置组
        config_layout.insertWidget(2, paper_group)  # 放在打印机设置下面

        # 在纸张设置组中添加打印参数
        print_params_group = QGroupBox("打印参数设置")
        print_params_layout = QHBoxLayout()

        # 黑白打印选项
        self.bw_print_check = QCheckBox("黑白打印")
        self.bw_print_check.setFont(QFont("Microsoft YaHei", 12))
        self.bw_print_check.setChecked(False)

        # 双面打印选项（可选）
        self.duplex_check = QCheckBox("双面打印")
        self.duplex_check.setFont(QFont("Microsoft YaHei", 12))
        self.duplex_check.setChecked(False)

        print_params_layout.addWidget(self.bw_print_check)
        print_params_layout.addWidget(self.duplex_check)
        print_params_layout.addStretch()
        print_params_group.setLayout(print_params_layout)

        # 添加到主布局（放在纸张设置下方）
        main_layout.insertWidget(1, print_params_group)

        # 设置布局间距和对齐
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # 设置控件大小策略
        for widget in [self.source_edit,
                       self.paper_size_spin, self.paper_zoom_spin,
                       self.delay_spin, self.wait_sleep_spin,
                       self.printer_combo, self.monthly_printer_combo, ]:
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

    def load_config(self):
        config = self.config_manager.get_all()
        self.source_edit.setText(config.get("source_dir", ""))
        self.paper_size_spin.setValue(config.get("default_paper_size", 132))
        self.paper_zoom_spin.setValue(config.get("default_paper_zoom", 75))
        self.delay_spin.setValue(config.get("delay_seconds", 5))
        self.wait_prompt_check.setChecked(config.get("enable_wait_prompt", True))
        self.wait_sleep_spin.setValue(config.get("wait_prompt_sleep", 30))
        self.bw_print_check.setChecked(config.get("bw_print", True))
        self.duplex_check.setChecked(config.get("duplex_print", False))

        # 设置纸张默认选择
        if self.paper_size_spin.value() < len(self.paper_sizes):
            default_paper_id = self.paper_size_spin.value() - 1
        else:
            default_paper_id = len(self.paper_sizes) - 1
        self.paper_combo.setCurrentIndex(default_paper_id)

    def get_default_printer_by_default_name(self):
        for index in range(self.printer_combo.count()):
            if "默认" in self.printer_combo.itemText(index):
                return self.printer_combo.itemText(index).replace(" (默认)", "")
        return self.get_selected_printer()

    def save_config(self, showAlert=True):
        if not self.source_edit.text():
            QMessageBox.warning(self, "警告", "请先选择源目录!")
            return

        config = {
            "source_dir": self.source_edit.text(),
            "default_paper_size": self.paper_size_spin.value(),
            "default_paper_zoom": self.paper_zoom_spin.value(),
            "delay_seconds": self.delay_spin.value(),
            "enable_wait_prompt": self.wait_prompt_check.isChecked(),
            "wait_prompt_sleep": self.wait_sleep_spin.value(),
            "bw_print": self.bw_print_check.isChecked(),
            "duplex_print": self.duplex_check.isChecked(),
        }

        # 保存默认打印机
        if hasattr(self, 'printer_combo'):
            # config['selected_printer'] = self.get_selected_printer()
            config['selected_printer'] = self.get_default_printer_by_default_name()

        # 保存月结单打印机
        if hasattr(self, 'monthly_printer_combo'):
            config['monthly_printer_name'] = self.monthly_printer_combo.currentText()

        # 保存打印参数
        if hasattr(self, 'bw_print'):
            config['bw_print'] = self.bw_print_check.isChecked()
        if hasattr(self, 'duplex_print'):
            config['duplex_print'] = self.duplex_check.isChecked()

        self.config_manager.update_config(config)
        if self.config_manager.save_config():
            if showAlert:
                QMessageBox.information(self, "成功", "配置已保存!")
        else:
            QMessageBox.warning(self, "错误", "保存配置失败!")

    def show_paper_info(self):
        """显示打印机支持的纸张信息"""
        try:
            import win32print

            if not win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL):
                self.log_message("⚠️ 未检测到可用的打印机")
                return

            printer_name = win32print.GetDefaultPrinter()
            self.log_message(f"\n🖨️ 正在获取打印机 '{printer_name}' 的纸张信息...")

            hprinter = win32print.OpenPrinter(printer_name)
            forms = win32print.EnumForms(hprinter)

            self.log_message(f"✅ 找到 {len(forms)} 种支持的纸张尺寸:")
            for i, form in enumerate(forms, 1):
                width_cm = form['Size']['cx'] / 1000
                height_cm = form['Size']['cy'] / 1000
                self.log_message(
                    f"{i}. {form['Name']} "
                    f"(宽度: {width_cm:.1f}cm × 高度: {height_cm:.1f}cm)"
                )

            win32print.ClosePrinter(hprinter)

        except win32print.error as e:
            self.log_message(f"❌ 打印机API错误: {str(e)}")
        except Exception as e:
            self.log_message(f"❌ 获取纸张信息失败: {str(e)}")

    # 设置默认打印机
    def set_default_printer(self):
        """将选中的打印机设置为系统默认打印机"""
        try:
            import win32print

            if self.printer_combo.count() == 0:
                self.log_message("⚠️ 没有可用的打印机")
                return

            selected_printer = self.get_selected_printer()

            # 尝试设置默认打印机
            win32print.SetDefaultPrinter(selected_printer)

            # 刷新列表显示新的默认打印机
            self.refresh_printer_list()

            self.log_message(f"✅ 已将 '{selected_printer}' 设置为默认打印机")

        except Exception as e:
            self.log_message(f"❌ 设置默认打印机失败: {str(e)}")
            # 尝试使用管理员权限
            if "Access is denied" in str(e):
                self.log_message("⚠️ 需要管理员权限，正在尝试获取权限...")
                self.run_as_admin(f'SetDefaultPrinter "{selected_printer}"')

    def run_as_admin(self, command):
        """尝试以管理员权限运行命令"""
        try:
            import ctypes
            import sys

            if ctypes.windll.shell32.IsUserAnAdmin():
                return True

            # 重新以管理员权限运行程序
            ctypes.windll.shell32.ShellExecuteW(
                None,
                "runas",
                sys.executable,
                f"{sys.argv[0]} {command}",
                None,
                1
            )
            return True
        except Exception as e:
            self.log_message(f"❌ 获取管理员权限失败: {str(e)}")
            return False

    def refresh_printer_list(self):
        """刷新打印机列表并设置默认选中"""
        try:
            import win32print

            self.log_message("\n🔄 正在刷新打印机列表...")
            self.printer_combo.clear()
            self.monthly_printer_combo.clear()

            # 获取所有打印机
            self.available_printers = win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            )

            # 获取默认打印机
            default_printer = win32print.GetDefaultPrinter()
            default_index = 0

            # 添加到下拉框
            for i, printer in enumerate(self.available_printers):
                printer_name = printer[2]
                self.printer_combo.addItem(printer_name)
                self.monthly_printer_combo.addItem(printer_name)

                # 标记默认打印机
                if printer_name == default_printer:
                    default_index = i
                    self.printer_combo.setItemText(i, f"{printer_name} (默认)")
                    # 加载该打印机的纸张类型
                    self.load_paper_sizes(printer_name)

            # 尝试恢复保存的月结单打印机设置
            saved_monthly_printer = self.config_manager.get("monthly_printer_name", "")
            if saved_monthly_printer:
                index = self.monthly_printer_combo.findText(saved_monthly_printer)
                if index >= 0:
                    self.monthly_printer_combo.setCurrentIndex(index)

            # 设置默认选中
            self.printer_combo.setCurrentIndex(default_index)
            self.log_message(f"✅ 已加载 {len(self.available_printers)} 台打印机")

        except Exception as e:
            self.log_message(f"❌ 刷新打印机列表失败: {str(e)}")

    def get_selected_printer(self):
        """获取当前选中的打印机"""
        if self.printer_combo.count() > 0:
            # 去除"(默认)"标记
            return self.printer_combo.currentText().replace(" (默认)", "")
        return win32print.GetDefaultPrinter()  # 回退到系统默认

    # 打印机设置
    def open_printer_settings(self):
        methods = [
            self._open_via_win32api,
            self._open_via_system_control,
            self._open_via_command
        ]

        for method in methods:
            if method():
                return

        self.log_message("❌ 所有方法尝试失败，请手动打开打印机设置")

    def _open_via_win32api(self):
        try:
            import win32api
            win32api.ShellExecute(0, "open", "control.exe", "printers", None, 1)
            return True
        except:
            return False

    def _open_via_system_control(self):
        try:
            import os
            os.system('control printers')
            return True
        except:
            return False

    def _open_via_command(self):
        try:
            import subprocess
            subprocess.run(['rundll32', 'printui.dll,PrintUIEntry'])
            return True
        except:
            return False

    # start: 打印纸列表设置
    def on_paper_selected(self, index):
        """当选择纸张类型时的处理"""
        if index >= 0 and self.paper_sizes:
            paper_id = self.paper_combo.itemData(index)
            self.log_message(f"📄 已选择纸张: {self.paper_combo.itemText(index)} (ID: {paper_id})")
            self.paper_size_spin.setValue(int(paper_id))

    def load_paper_sizes(self, printer_name):
        """加载指定打印机支持的纸张类型"""
        try:
            import win32print

            self.paper_combo.clear()
            self.paper_sizes.clear()

            hprinter = win32print.OpenPrinter(win32print.GetDefaultPrinter())
            forms = win32print.EnumForms(hprinter)
            win32print.ClosePrinter(hprinter)

            # 常见针式打印机纸张类型映射
            dot_matrix_papers = {
                # 1: "Letter 8.5x11英寸",
                # 2: "Legal 8.5x14英寸",
                # 5: "A4 210x297毫米",
                # 132: "连续纸(80列)",
                # 133: "连续纸(132列)",
                # 256: "自定义纸张"
            }

            self.log_message(f"✅ 找到 {len(forms)} 种支持的纸张尺寸:")
            for i, form in enumerate(forms, 1):
                width_cm = form['Size']['cx'] / 1000
                height_cm = form['Size']['cy'] / 1000
                # self.log_message(
                #     f"{i}. {form['Name']} "
                #     f"(宽度: {width_cm:.1f}cm × 高度: {height_cm:.1f}cm)"
                # )
                dot_matrix_papers[i] = f"{i}. {form['Name']}(宽度: {width_cm:.1f}cm × 高度: {height_cm:.1f}cm)"

            # 添加特殊纸张类型
            self.paper_sizes = {}
            for paper_id, paper_name in dot_matrix_papers.items():
                self.paper_combo.addItem(f"{paper_name}", paper_id)
                self.paper_sizes[paper_id] = paper_name

            # 添加打印机实际支持的纸张类型
            for form in forms:
                if form['Name'] not in dot_matrix_papers.values():
                    paper_id = self._get_paper_id_by_size(form['Size'])
                    if paper_id:
                        self.paper_combo.addItem(
                            f"{form['Name']} (ID:{paper_id})",
                            paper_id
                        )
                        self.paper_sizes[paper_id] = form['Name']

            # # 列表默认设置最后一个打印纸设置
            # default_paper_id = len(dot_matrix_papers)
            # index = self.paper_combo.findData(default_paper_id)
            # if index >= 0:
            #     self.paper_combo.setCurrentIndex(index)

            return True

        except Exception as e:
            self.log_message(f"❌ 加载纸张类型失败: {str(e)}")
            return False

    def _get_paper_id_by_size(self, size):
        """根据纸张尺寸获取标准ID"""
        width, height = size['cx'], size['cy']
        # A4: 210x297 mm (转换为0.1mm单位)
        if abs(width - 2100) < 50 and abs(height - 2970) < 50:
            return 9
        # Letter: 215.9x279.4 mm
        elif abs(width - 2159) < 50 and abs(height - 2794) < 50:
            return 1
        # 连续纸(80列)
        elif width == 2410 and height == 2794:
            return 132
        # 连续纸(132列)
        elif width == 3810 and height == 2794:
            return 133
        return None

    # end: 打印纸列表设置

    def show_printer_info(self):
        self.log_edit.clear()
        """显示完整的打印机信息"""
        try:
            import win32print

            printer_name = win32print.GetDefaultPrinter()
            hprinter = win32print.OpenPrinter(printer_name)

            # 获取打印机详细信息
            info_level = 2
            printer_info = win32print.GetPrinter(hprinter, info_level)

            self.log_message("\n📋 打印机详细信息:")
            self.log_message(f"名称: {printer_info['pPrinterName']}")
            self.log_message(f"驱动程序: {printer_info['pDriverName']}")
            self.log_message(f"端口: {printer_info['pPortName']}")
            self.log_message(f"状态: {self.get_printer_status(printer_info['Status'])}")

            # 显示纸张信息
            self.show_paper_info()

            win32print.ClosePrinter(hprinter)

        except Exception as e:
            self.log_message(f"❌ 获取打印机信息失败: {str(e)}")

    def get_printer_status(self, status_code):
        """将状态代码转换为可读文本"""
        status_map = {
            0: "准备就绪",
            1: "暂停",
            2: "错误",
            3: "待删除",
            4: "纸张卡住",
            5: "纸张用完",
            # ... 其他状态码 ...
        }
        return status_map.get(status_code, f"未知状态 ({status_code})")

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

        self.save_config(False)  # 开始前自动保存配置，不提示保存成功弹出框

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
