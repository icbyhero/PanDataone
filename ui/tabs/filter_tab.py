"""
数据筛选标签页组件

该模块提供了一个用于Excel数据筛选和分析的用户界面标签页。
包含文件选择、统计信息展示和进度显示等功能。
"""

import os
import logging
from typing import Optional, Dict

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QScrollArea, QGroupBox, QProgressBar,
    QPushButton, QLabel, QSizePolicy, QFrame, QMessageBox
)
from PySide6.QtCore import Qt, Signal

from ui.widgets.drop_zone import DropZoneGroupBox
from ui.widgets.stat_card import StatCard


class FilterTab(QWidget):
    """
    数据筛选标签页组件

    该组件提供了一个完整的数据筛选和分析界面，包括：
    - 文件选择区域（支持拖拽）
    - 统计信息卡片展示
    - 进度条显示
    - 浏览文件和开始分析按钮

    Signals:
        file_browsed: 当用户点击浏览文件按钮时发出
        analyze_clicked: 当用户点击开始分析按钮时发出
        file_dropped: 当文件被拖放到区域时发出，参数为文件路径(str)
    """

    # 信号定义
    file_browsed = Signal()
    analyze_clicked = Signal()
    file_dropped = Signal(str)

    def __init__(self, parent: Optional[QWidget] = None):
        """
        初始化数据筛选标签页

        Args:
            parent: 父窗口组件
        """
        super().__init__(parent)
        self.current_file_path: str = ""

        # UI组件
        self.file_group: DropZoneGroupBox
        self.analyze_button: QPushButton
        self.progress_bar: QProgressBar
        self.stat_total: StatCard
        self.stat_matched: StatCard
        self.stat_unmatched: StatCard
        self.stat_rate: StatCard

        self._setup_ui()

    def _setup_ui(self):
        """设置用户界面"""
        # 创建主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # 创建滚动区域
        scroll_area = self._create_scroll_area()

        # 创建容器
        container = self._create_container()
        scroll_area.setWidget(container)

        main_layout.addWidget(scroll_area)

    def _create_scroll_area(self) -> QScrollArea:
        """
        创建滚动区域

        Returns:
            配置好的滚动区域组件
        """
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        return scroll_area

    def _create_container(self) -> QWidget:
        """
        创建主容器组件

        Returns:
            包含所有UI元素的容器组件
        """
        container = QWidget()
        container.setStyleSheet("background-color: #FAFAFA;")
        layout = QVBoxLayout(container)
        layout.setSpacing(20)

        # 添加帮助区域
        layout.addWidget(self._create_help_section())

        # 添加文件选择区域
        self.file_group = self._create_file_selection_section()
        layout.addWidget(self.file_group)

        # 添加统计信息区域
        stats_group = self._create_stats_section()
        layout.addWidget(stats_group, 3)  # stretch=3 让统计区域占据更多空间

        # 添加进度条
        self.progress_bar = self._create_progress_bar()
        layout.addWidget(self.progress_bar)

        return container

    def _create_help_section(self) -> QFrame:
        """
        创建帮助提示区域

        Returns:
            包含帮助信息的框架组件
        """
        widget = QFrame()
        widget.setStyleSheet("""
            QFrame {
                background-color: #E3F2FD;
                border: 1px solid #BBDEFB;
                border-radius: 6px;
                padding: 10px;
            }
        """)

        layout = QHBoxLayout(widget)
        layout.setContentsMargins(15, 10, 15, 10)

        icon_label = QLabel("💡")
        icon_label.setStyleSheet("font-size: 20px;")

        text_label = QLabel(
            "将包含两个工作表的Excel文件拖放到上方区域，"
            "第一个为待匹配表，第二个为匹配原表"
        )
        text_label.setWordWrap(True)
        text_label.setStyleSheet("color: #1976D2; font-size: 12px;")

        layout.addWidget(icon_label)
        layout.addWidget(text_label, 1)

        return widget

    def _create_file_selection_section(self) -> DropZoneGroupBox:
        """
        创建文件选择区域

        Returns:
            文件选择分组框组件
        """
        file_group = DropZoneGroupBox("📁 文件选择")
        file_group.setMinimumHeight(300)
        file_group.file_selected.connect(self._on_file_selected)

        # 创建按钮布局
        button_layout = self._create_button_layout()
        file_group.add_button_layout(button_layout)

        return file_group

    def _create_button_layout(self) -> QHBoxLayout:
        """
        创建按钮布局

        Returns:
            包含浏览文件和开始分析按钮的布局
        """
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        # 浏览文件按钮
        browse_button = self._create_browse_button()
        button_layout.addWidget(browse_button)

        # 开始分析按钮
        self.analyze_button = self._create_analyze_button()
        button_layout.addWidget(self.analyze_button)

        button_layout.addStretch()

        return button_layout

    def _create_browse_button(self) -> QPushButton:
        """
        创建浏览文件按钮

        Returns:
            配置好的按钮组件
        """
        browse_button = QPushButton("📂 浏览文件")
        browse_button.clicked.connect(self._on_browse_clicked)
        browse_button.setMinimumHeight(45)
        browse_button.setMinimumWidth(150)
        browse_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 12px 24px;
                border-radius: 8px;
                font-weight: bold;
                font-size: 15px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        return browse_button

    def _create_analyze_button(self) -> QPushButton:
        """
        创建开始分析按钮

        Returns:
            配置好的按钮组件
        """
        analyze_button = QPushButton("🚀 开始分析")
        analyze_button.clicked.connect(self._on_analyze_clicked)
        analyze_button.setEnabled(False)  # 初始状态不可点击
        analyze_button.setMinimumHeight(45)
        analyze_button.setMinimumWidth(150)
        analyze_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 12px 24px;
                border-radius: 8px;
                font-weight: bold;
                font-size: 15px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #E0E0E0;
                color: #9E9E9E;
                border: 1px solid #D0D0D0;
            }
        """)
        return analyze_button

    def _create_stats_section(self) -> QGroupBox:
        """
        创建统计信息区域

        Returns:
            包含统计卡片的分组框
        """
        stats_group = QGroupBox("📊 分析统计")
        stats_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        stats_group.setMinimumHeight(500)

        stats_layout = QVBoxLayout(stats_group)
        stats_layout.setSpacing(15)
        stats_layout.setContentsMargins(10, 20, 10, 10)

        # 创建第一行卡片
        first_row = self._create_stat_row([
            ("总数据", "0", "📋", "待处理的数据总数"),
            ("已匹配", "0", "✅", "成功匹配到的数据条数")
        ])
        stats_layout.addWidget(first_row)

        # 创建第二行卡片
        second_row = self._create_stat_row([
            ("未匹配", "0", "❌", "未找到对应的数据条数"),
            ("匹配率", "0%", "📈", "成功匹配的百分比")
        ])
        stats_layout.addWidget(second_row)

        stats_layout.addStretch()

        return stats_group

    def _create_stat_row(self, card_configs: list) -> QWidget:
        """
        创建统计卡片行

        Args:
            card_configs: 卡片配置列表，每个元素为(title, value, icon, description)元组

        Returns:
            包含统计卡片的容器组件
        """
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setSpacing(15)

        for i, config in enumerate(card_configs):
            title, value, icon, description = config
            stat_card = StatCard(title, value, icon, description)
            stat_card.setMinimumHeight(220)
            stat_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            row_layout.addWidget(stat_card)

            # 保存引用以便后续更新
            if i == 0 and title == "总数据":
                self.stat_total = stat_card
            elif i == 1 and title == "已匹配":
                self.stat_matched = stat_card
            elif i == 0 and title == "未匹配":
                self.stat_unmatched = stat_card
            elif i == 1 and title == "匹配率":
                self.stat_rate = stat_card

        return row

    def _create_progress_bar(self) -> QProgressBar:
        """
        创建进度条组件

        Returns:
            配置好的进度条组件
        """
        progress_bar = QProgressBar()
        progress_bar.setVisible(False)
        progress_bar.setFixedHeight(25)
        return progress_bar

    def _on_browse_clicked(self):
        """处理浏览文件按钮点击事件"""
        logging.info("用户点击浏览文件按钮")
        self.file_browsed.emit()

    def _on_analyze_clicked(self):
        """处理开始分析按钮点击事件"""
        logging.info("用户点击开始分析按钮")
        self.analyze_clicked.emit()

    def _on_file_selected(self, file_path: str):
        """
        处理文件选择事件

        Args:
            file_path: 选中的文件路径
        """
        if os.path.exists(file_path):
            self.current_file_path = file_path
            self.analyze_button.setEnabled(True)
            self.file_dropped.emit(file_path)
            logging.info(f"已选择文件: {file_path}")
        else:
            QMessageBox.warning(self, "警告", "文件不存在")
            logging.error(f"文件不存在: {file_path}")

    def set_file_path(self, file_path: str):
        """
        设置当前文件路径

        Args:
            file_path: 文件路径
        """
        self.current_file_path = file_path
        if file_path and os.path.exists(file_path):
            self.analyze_button.setEnabled(True)
            # 更新文件选择区域的显示
            self.file_group._update_display(file_path)
        else:
            self.analyze_button.setEnabled(False)

    def get_file_path(self) -> str:
        """
        获取当前文件路径

        Returns:
            当前文件路径
        """
        return self.current_file_path

    def update_stats(self, stats: Dict[str, any]):
        """
        更新统计信息

        Args:
            stats: 统计信息字典，包含total, matched, unmatched, rate等字段
        """
        self.stat_total.update_value(str(stats.get('total', 0)))
        self.stat_matched.update_value(str(stats.get('matched', 0)))
        self.stat_unmatched.update_value(str(stats.get('unmatched', 0)))
        self.stat_rate.update_value(f"{stats.get('rate', 0)}%")
        logging.info(f"统计信息已更新: {stats}")

    def set_progress_visible(self, visible: bool):
        """
        设置进度条可见性

        Args:
            visible: 是否显示进度条
        """
        self.progress_bar.setVisible(visible)

    def set_progress_value(self, value: int):
        """
        设置进度条当前值

        Args:
            value: 进度值（0-最大值）
        """
        self.progress_bar.setValue(value)

    def set_progress_maximum(self, maximum: int):
        """
        设置进度条最大值

        Args:
            maximum: 最大进度值
        """
        self.progress_bar.setMaximum(maximum)

    def enable_analyze_button(self, enabled: bool = True):
        """
        启用或禁用分析按钮

        Args:
            enabled: True为启用，False为禁用
        """
        self.analyze_button.setEnabled(enabled)
