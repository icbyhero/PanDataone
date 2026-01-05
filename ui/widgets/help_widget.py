"""
帮助组件模块

该模块提供了精简版和详细版的帮助提示组件。
用于显示应用使用说明和操作指引。

主要功能:
- 显示精简版帮助信息（紧凑布局）
- 点击按钮查看详细帮助（弹窗）
- 蓝色主题，信息提示风格

使用示例:
    help_widget = HelpWidget(parent)
    help_widget.detail_requested.connect(show_detailed_help)
    layout.addWidget(help_widget)
"""

from PySide6.QtWidgets import QFrame, QLabel, QPushButton, QHBoxLayout, QMessageBox
from PySide6.QtCore import Signal


class HelpWidget(QFrame):
    """
    帮助提示组件

    提供紧凑型帮助信息显示，带有"查看详情"按钮可展开完整说明。

    布局结构:
        - 左侧: 图标(💡)
        - 中间: 简短提示文字(弹性空间)
        - 右侧: 查看详情按钮

    样式特点:
        - 浅蓝色背景(#E3F2FD)
        - 蓝色边框(#BBDEFB)
        - 圆角设计
        - 图标使用大号emoji

    信号:
        detail_requested: 当点击"查看详情"按钮时发射
    """

    detail_requested = Signal()

    def __init__(self, parent=None):
        """
        初始化帮助组件

        Args:
            parent: 父窗口对象，默认为None
        """
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        """
        设置用户界面

        创建帮助提示的视觉元素:
        - 图标标签(💡)
        - 提示文字(可自定义)
        - 查看详情按钮
        """
        # 设置组件样式
        self.setStyleSheet("""
            QFrame {
                background-color: #E3F2FD;
                border: 1px solid #BBDEFB;
                border-radius: 6px;
                padding: 10px;
            }
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 10, 15, 10)

        # 图标 - 左侧
        icon_label = QLabel("💡")
        icon_label.setStyleSheet("font-size: 20px;")

        # 提示文字 - 中间
        text_label = QLabel(
            "将包含两个工作表的Excel文件拖放到上方区域，第一个为待匹配表，第二个为匹配原表"
        )
        text_label.setWordWrap(True)
        text_label.setStyleSheet("color: #1976D2; font-size: 12px;")

        # 查看详情按钮 - 右侧
        toggle_button = QPushButton("查看详情")
        toggle_button.setCheckable(True)
        toggle_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #1976D2;
                border: 1px solid #1976D2;
                padding: 5px 15px;
                border-radius: 4px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #BBDEFB;
            }
        """)
        toggle_button.clicked.connect(self._on_detail_clicked)

        layout.addWidget(icon_label)
        layout.addWidget(text_label, 1)  # stretch=1 占据剩余空间
        layout.addWidget(toggle_button)

    def _on_detail_clicked(self):
        """
        处理查看详情按钮点击事件

        发射detail_requested信号，由父窗口处理详细帮助的显示。
        """
        self.detail_requested.emit()

    @staticmethod
    def get_default_help_text() -> str:
        """
        获取默认的详细帮助文本

        Returns:
            str: 格式化的帮助文本内容，包含数据准备、格式要求、结果说明和使用技巧
        """
        return """使用说明：

1. 📁 数据准备：
   • 第一个工作表为"供应商待匹配表"，放入需要查询的数据
   • 第二个工作表为"供应商匹配原表"，放入用于匹配的数据
   • 两个工作表的前三列必须包含：日期、客户名称、产品名称

2. 📅 数据格式要求：
   • 日期格式支持：2024-03、24年3月、3月、202411-12
   • 客户名称：不区分全角半角，自动处理空格
   • 产品名称：不区分大小写，自动处理特殊符号

3. 🎨 处理结果说明：
   • 🟩绿色：在匹配原表中找到对应数据
   • 🟥红色：在匹配原表中未找到对应数据
   • 🟨黄色：该数据重复查询（最高优先级）
   • 🟫棕色：日期范围内的数据未能全部匹配成功
   • 🟪紫色：日期范围内的数据全部匹配成功

4. 💡 使用技巧：
   • 可以直接拖拽Excel文件到窗口
   • 支持批量处理大量数据
   • 分析结果会自动保存到原文件"""

    def show_detailed_help_dialog(self, parent=None):
        """
        显示详细帮助对话框

        使用默认帮助文本显示信息弹窗。

        Args:
            parent: 父窗口对象，用于对话框定位，默认为None
        """
        help_text = self.get_default_help_text()
        QMessageBox.information(
            parent if parent else self,
            "使用说明",
            help_text,
            QMessageBox.Ok
        )

    def set_help_text(self, text: str):
        """
        自定义简短帮助文字

        用于修改中间显示的提示文字内容。

        Args:
            text: 新的提示文字
        """
        # 查找文本标签并更新
        for child in self.findChildren(QLabel):
            if child.wordWrap() and "Excel文件" in child.text():
                child.setText(text)
                break
