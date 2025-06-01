import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QFont
from ui.main_window import MainWindow


def main():
    app = QApplication(sys.argv)

    # 全局字体设置
    font = QFont()
    font.setFamily("Microsoft YaHei")
    font.setPointSize(14)  # 基础大小
    # font.setBold(True)  # 标题加粗
    app.setFont(font)

    # 设置应用程序样式
    app.setStyle('Fusion')

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
