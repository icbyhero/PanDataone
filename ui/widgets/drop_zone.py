"""
拖拽区域组件模块

该模块提供了支持拖拽文件的GroupBox组件，用于Excel文件选择。
整个卡片区域都支持拖放操作，提供直观的文件选择体验。

主要功能:
- 支持拖拽Excel文件到整个GroupBox区域
- 提供视觉反馈（拖拽进入、放下时的样式变化）
- 显示文件信息（文件名、路径、大小）
- 发射信号通知主窗口文件已选择

使用示例:
    drop_zone = DropZoneGroupBox("📁 文件选择")
    drop_zone.file_selected.connect(on_file_selected)
    layout.addWidget(drop_zone)
"""

import os
from typing import Optional
from PySide6.QtWidgets import QGroupBox, QWidget, QVBoxLayout, QHBoxLayout, QLabel
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent


class DropZoneGroupBox(QGroupBox):
    """
    支持拖拽的文件选择GroupBox - 整个卡片都支持拖放

    该组件提供了一个支持拖拽文件的GroupBox，整个卡片区域都可以接收拖放的Excel文件。
    拖拽时会提供视觉反馈，放下后会显示文件信息并发射信号。

    信号:
        file_selected: 当文件被成功选择时发射，参数为文件路径(str)

    属性:
        current_file_path: 当前选择的文件路径

    样式特点:
        - 默认状态: 灰色虚线边框，浅灰背景
        - 悬停状态: 绿色边框，浅绿背景
        - 拖拽进入: 加粗绿色边框，更明显的浅绿背景
        - 文件已选择: 显示✅图标和文件信息
    """

    # 定义信号
    file_selected = Signal(str)

    def __init__(self, title: str, parent=None):
        """
        初始化拖拽区域组件

        Args:
            title: GroupBox的标题文本
            parent: 父窗口对象，默认为None
        """
        super().__init__(title, parent)
        self.setAcceptDrops(True)  # 启用拖拽
        self.current_file_path = ""  # 存储当前文件路径
        self._setup_ui()

    def _setup_ui(self):
        """
        设置用户界面

        创建拖拽区域的视觉元素:
        - 图标标签(📁)
        - 标题文本("拖拽Excel文件到这里")
        - 副标题文本("整个卡片都支持拖放文件")
        - 按钮容器(用于放置浏览按钮等)
        """
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # 拖拽区域内容 - 添加到主布局
        drag_content = QWidget()
        drag_layout = QVBoxLayout(drag_content)
        drag_layout.setAlignment(Qt.AlignCenter)
        drag_layout.setSpacing(10)

        # 图标和提示文字
        icon_label = QLabel("📁")
        icon_label.setStyleSheet("font-size: 64px;")
        icon_label.setAlignment(Qt.AlignCenter)

        title_label = QLabel("拖拽Excel文件到这里")
        title_label.setObjectName("drop_title_label")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #546E7A;
                padding: 10px 0px;
            }
        """)
        title_label.setAlignment(Qt.AlignCenter)

        subtitle_label = QLabel("整个卡片都支持拖放文件")
        subtitle_label.setObjectName("drop_subtitle_label")
        subtitle_label.setStyleSheet("""
            QLabel {
                font-size: 13px;
                color: #607D8B;
                padding: 5px 0px;
            }
        """)
        subtitle_label.setAlignment(Qt.AlignCenter)

        drag_layout.addWidget(icon_label)
        drag_layout.addWidget(title_label)
        drag_layout.addWidget(subtitle_label)

        # 设置整个GroupBox的样式
        self.setStyleSheet("""
            QGroupBox {
                border: 2px dashed #CCCCCC;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 20px;
                font-weight: bold;
                background-color: #FAFAFA;
            }
            QGroupBox:hover {
                border-color: #4CAF50;
                background-color: #F0F8F0;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #607D8B;
            }
        """)

        self.drag_content = drag_content
        layout.addWidget(drag_content)

        # 按钮布局容器
        self.button_container = QWidget()
        layout.addWidget(self.button_container)

    def add_button_layout(self, button_layout: QHBoxLayout):
        """
        添加按钮布局到组件底部

        用于添加浏览按钮、开始分析按钮等控件。

        Args:
            button_layout: 要添加的按钮布局(QHBoxLayout)
        """
        container_layout = QVBoxLayout(self.button_container)
        container_layout.addLayout(button_layout)

    def dragEnterEvent(self, event: QDragEnterEvent):
        """
        拖拽进入事件处理器

        当拖拽的对象进入该组件区域时触发。
        如果拖拽的是URL（文件），则接受拖拽并更新样式。

        Args:
            event: 拖拽进入事件对象
        """
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            # 更新为拖拽激活样式
            self.setStyleSheet("""
                QGroupBox {
                    border: 3px dashed #4CAF50;
                    border-radius: 8px;
                    margin-top: 10px;
                    padding-top: 20px;
                    font-weight: bold;
                    background-color: #E8F5E9;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 10px;
                    padding: 0 5px;
                    color: #4CAF50;
                }
            """)

    def dragLeaveEvent(self, event):
        """
        拖拽离开事件处理器

        当拖拽的对象离开该组件区域时触发。
        重置为默认样式。

        Args:
            event: 拖拽离开事件对象
        """
        # 重置为默认样式
        self.setStyleSheet("""
            QGroupBox {
                border: 2px dashed #CCCCCC;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 20px;
                font-weight: bold;
                background-color: #FAFAFA;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #607D8B;
            }
        """)

    def dropEvent(self, event: QDropEvent):
        """
        拖拽放下事件处理器

        当拖拽的对象在该组件区域放下时触发。
        验证文件类型（仅接受.xlsx和.xls），更新显示并发射信号。

        Args:
            event: 拖拽放下事件对象
        """
        # 从拖拽事件中获取文件路径
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            file_path = files[0]
            # 验证文件类型
            if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                self.file_selected.emit(file_path)  # 发射信号
                self._update_display(file_path)  # 更新显示
            else:
                from PySide6.QtWidgets import QMessageBox
                QMessageBox.warning(self, "警告", "请选择Excel文件(.xlsx或.xls)")
        self._reset_style()  # 重置样式

    def _reset_style(self):
        """
        重置组件样式到默认状态

        恢复组件的默认视觉外观。
        """
        self.setStyleSheet("""
            QGroupBox {
                border: 2px dashed #CCCCCC;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 20px;
                font-weight: bold;
                background-color: #FAFAFA;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #607D8B;
            }
        """)

    def _update_display(self, file_path: str):
        """
        更新显示文件信息

        当文件成功选择后，更新界面显示:
        - 标题显示文件名（带✅图标）
        - 副标题显示完整路径和文件大小
        - 颜色变为绿色表示成功

        Args:
            file_path: 已选择文件的完整路径
        """
        self.current_file_path = file_path
        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path)
        file_size_mb = file_size / (1024 * 1024)  # 转换为MB

        # 更新标题 - 显示文件名
        title = self.findChild(QLabel, "drop_title_label")
        if title:
            title.setText(f"✅ {file_name}")
            title.setStyleSheet("""
                QLabel {
                    font-size: 18px;
                    font-weight: bold;
                    color: #4CAF50;
                    padding: 10px 0px;
                }
            """)

        # 更新副标题 - 显示文件路径和大小
        subtitle = self.findChild(QLabel, "drop_subtitle_label")
        if subtitle:
            subtitle.setText(f"📂 {file_path}\n📊 文件大小: {file_size_mb:.2f} MB")
            subtitle.setStyleSheet("""
                QLabel {
                    font-size: 12px;
                    color: #4CAF50;
                    padding: 5px 0px;
                }
            """)
