"""
应用样式模块

该模块提供了应用程序的全局样式设置。
采用现代化的配色方案和macOS风格的滚动条设计。

主要功能:
- 设置应用程序全局样式
- macOS风格的滚动条
- 统一的配色方案
- 响应式控件样式

使用示例:
    from PySide6.QtWidgets import QApplication
    app = QApplication(sys.argv)
    apply_app_style(app)
"""

from PySide6.QtWidgets import QApplication


def apply_app_style(app: QApplication):
    """
    应用应用程序全局样式

    设置应用程序的全局外观样式，包括:
    - 字体和字号
    - 配色方案
    - 控件样式
    - macOS风格滚动条
    - 按钮状态样式

    设计理念:
        - 使用柔和的蓝灰色系(#546E7A)作为主色调
        - 绿色(#4CAF50)作为强调色和成功状态
        - 减少黑色使用，提升视觉舒适度
        - 圆角设计增强现代感
        - 苹果风格的滚动条更加精致

    Args:
        app: QApplication实例

    样式包含:
        1. 全局字体: Microsoft YaHei / PingFang SC, 13px
        2. 滚动条: macOS风格，圆角半透明
        3. 按钮: 绿色主题，悬停/按下效果
        4. 输入框: 圆角边框，焦点高亮
        5. 进度条: 绿色填充块
        6. 分组框: 白色背景，圆角边框
    """
    # 设置Fusion风格（跨平台一致的外观）
    QApplication.setStyle("Fusion")

    # 应用全局样式表
    app.setStyleSheet("""
        /* ========== 主窗口样式 ========== */
        QMainWindow {
            background-color: #F5F5F5;
        }

        /* ========== 全局字体设置 ========== */
        QWidget {
            font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
            font-size: 13px;
            color: #546E7A;
        }

        /* ========== 苹果风格滚动条 - 垂直 ========== */
        QScrollBar:vertical {
            border: none;
            background: transparent;
            width: 10px;
            margin: 0px;
        }
        QScrollBar::handle:vertical {
            background: #C1C1C1;
            min-height: 30px;
            border-radius: 5px;
            margin: 2px;
        }
        QScrollBar::handle:vertical:hover {
            background: #A8A8A8;
        }
        QScrollBar::handle:vertical:pressed {
            background: #8F8F8F;
        }
        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical {
            border: none;
            background: none;
            height: 0px;
        }
        QScrollBar::add-page:vertical,
        QScrollBar::sub-page:vertical {
            background: none;
        }

        /* ========== 苹果风格滚动条 - 水平 ========== */
        QScrollBar:horizontal {
            border: none;
            background: transparent;
            height: 10px;
            margin: 0px;
        }
        QScrollBar::handle:horizontal {
            background: #C1C1C1;
            min-width: 30px;
            border-radius: 5px;
            margin: 2px;
        }
        QScrollBar::handle:horizontal:hover {
            background: #A8A8A8;
        }
        QScrollBar::handle:horizontal:pressed {
            background: #8F8F8F;
        }
        QScrollBar::add-line:horizontal,
        QScrollBar::sub-line:horizontal {
            border: none;
            background: none;
            width: 0px;
        }
        QScrollBar::add-page:horizontal,
        QScrollBar::sub-page:horizontal {
            background: none;
        }

        /* ========== 滚动区域样式 ========== */
        QScrollArea {
            border: none;
            background-color: transparent;
        }

        /* ========== 按钮样式 ========== */
        QPushButton {
            background-color: #4CAF50;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 6px;
            font-weight: bold;
            font-size: 14px;
        }
        QPushButton:hover {
            background-color: #45a049;
        }
        QPushButton:pressed {
            background-color: #3d8b40;
        }
        QPushButton:disabled {
            background-color: #E0E0E0;
            color: #9E9E9E;
        }

        /* ========== 输入框样式 ========== */
        QLineEdit {
            padding: 8px 12px;
            border: 1px solid #CFD8DC;
            border-radius: 4px;
            background-color: white;
            font-size: 13px;
            color: #546E7A;
        }
        QLineEdit:focus {
            border: 1px solid #4CAF50;
        }

        /* ========== 进度条样式 ========== */
        QProgressBar {
            border: none;
            border-radius: 4px;
            background-color: #E0E0E0;
            text-align: center;
            font-size: 12px;
            color: #546E7A;
        }
        QProgressBar::chunk {
            background-color: #4CAF50;
            border-radius: 4px;
        }

        /* ========== 分组框样式 ========== */
        QGroupBox {
            border: 1px solid #E0E0E0;
            border-radius: 8px;
            margin-top: 10px;
            padding-top: 10px;
            font-weight: bold;
            background-color: white;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px;
            color: #607D8B;
        }
    """)


def get_style_guide() -> str:
    """
    获取样式使用指南

    Returns:
        str: 样式系统的说明文档
    """
    return """
应用样式指南
============

配色方案:
  主色调: #546E7A (蓝灰色 - 文字和说明)
  强调色: #4CAF50 (绿色 - 主要操作和成功状态)
  背景色: #F5F5F5 (浅灰 - 主背景)
  卡片色: #FFFFFF (白色 - 内容卡片)
  边框色: #E0E0E0 (浅灰 - 边框线)

字体系统:
  默认字体: Microsoft YaHei / PingFang SC
  默认字号: 13px
  按钮字号: 14px, 加粗
  标题字号: 13px, 加粗

控件样式:
  - 按钮: 圆角6px, 绿色主题
  - 输入框: 圆角4px, 焦点时绿色边框
  - 进度条: 圆角4px, 绿色填充
  - 分组框: 圆角8px, 白色背景
  - 滚动条: macOS风格, 圆角5px, 半透明

滚动条特点:
  - 宽度/高度: 10px
  - 滑块颜色: #C1C1C1 (默认), #A8A8A8 (悬停), #8F8F8F (按下)
  - 圆角: 5px
  - 无上下箭头按钮
  - 透明背景和轨道

使用建议:
  1. 保持一致的圆角大小 (6px/8px)
  2. 使用绿色表示主要操作和成功状态
  3. 使用浅灰色表示次要信息和禁用状态
  4. 避免使用纯黑色文字
  5. 保持足够的内边距(padding)提升可读性
    """


if __name__ == '__main__':
    # 测试样式应用
    import sys
    from PySide6.QtWidgets import QMainWindow, QLabel, QVBoxLayout, QWidget, QPushButton

    app = QApplication(sys.argv)
    apply_app_style(app)

    window = QMainWindow()
    window.setWindowTitle("样式测试")
    window.setMinimumSize(600, 400)

    central = QWidget()
    layout = QVBoxLayout(central)

    label = QLabel("这是一个样式测试窗口\n检查字体、颜色、控件样式是否正确")
    label.setAlignment(Qt.AlignCenter)

    button = QPushButton("测试按钮")

    layout.addWidget(label)
    layout.addWidget(button)

    window.setCentralWidget(central)
    window.show()

    sys.exit(app.exec())
