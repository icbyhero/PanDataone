"""
PanDataone 供应商数据匹配系统

主入口文件

功能：
    供应商数据匹配和分析工具，用于处理Excel格式的采购数据和供应商报价单。
    支持数据标准化、智能匹配、重复检测和统计分析。

Author:
    Your Name

Version:
    1.0.0
"""

import sys
from PySide6.QtWidgets import QApplication

from ui.main_window import MainWindow
from ui.styles import apply_app_style


def main():
    """应用程序主入口函数

    创建并启动GUI应用程序。

    Returns:
        int: 应用程序退出码（0表示正常退出）
    """
    # 创建Qt应用程序实例
    app = QApplication(sys.argv)

    # 设置应用程序信息
    app.setApplicationName("PanDataone")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("PanDataone")

    # 应用全局样式
    apply_app_style(app)

    # 创建并显示主窗口
    # MainWindow 内部会自动初始化日志系统
    main_window = MainWindow()
    main_window.show()

    # 进入事件循环
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
