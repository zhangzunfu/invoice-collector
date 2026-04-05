#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
发票归集软件启动文件
"""

import sys
import os
# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from PyQt6.QtWidgets import QApplication
from ui.mainwindow import MainWindow


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("发票归集软件")
    app.setApplicationVersion("1.0.0")
    
    # 设置应用样式
    app.setStyle('Fusion')  # 使用Fusion风格以获得更好的跨平台外观
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()