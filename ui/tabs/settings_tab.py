"""
设置标签页组件

提供应用程序设置和关于信息的用户界面组件。
包括日志记录开关和应用说明信息。
"""

import os
import logging
from typing import Optional
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QCheckBox, QLabel
)
from PySide6.QtCore import QSettings, Signal


class SettingsTab(QWidget):
    """
    设置标签页组件

    提供应用程序的设置选项和关于信息，包括：
    - 日志记录开关
    - 关于应用说明
    - 版本信息展示

    Attributes:
        logging_toggled: 信号，当日志设置切换时发出，参数为新的启用状态(bool)
    """

    # 定义信号
    logging_toggled = Signal(bool)

    def __init__(self, parent=None):
        """
        初始化设置标签页

        Args:
            parent: 父窗口部件
        """
        super().__init__(parent)
        self.settings = QSettings('供应商数据智能匹配系统', 'DataAnalysis')
        self.log_file: Optional[str] = None
        self.log_path_label: Optional[QLabel] = None  # 保存标签引用以便更新
        self._setup_ui()

    def _setup_ui(self):
        """设置用户界面"""
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(20, 20, 20, 20)

        # 添加日志设置组
        layout.addWidget(self._create_logging_group())

        # 添加关于信息组
        layout.addWidget(self._create_about_group())

        # 添加弹性空间
        layout.addStretch()

    def _create_logging_group(self) -> QGroupBox:
        """
        创建日志设置组

        Returns:
            QGroupBox: 包含日志设置选项的组框
        """
        log_group = QGroupBox("📝 日志设置")
        log_layout = QVBoxLayout(log_group)

        # 日志启用复选框
        log_checkbox = QCheckBox("启用日志记录")
        log_checkbox.setChecked(self.settings.value('enable_logging', False, bool))
        log_checkbox.stateChanged.connect(self._on_logging_changed)
        log_layout.addWidget(log_checkbox)

        # 日志文件路径显示
        self.log_path_label = QLabel(f"日志文件位置：{self._get_log_file_path()}")
        self.log_path_label.setWordWrap(True)
        self.log_path_label.setStyleSheet("color: #666666; font-size: 11px;")
        log_layout.addWidget(self.log_path_label)

        return log_group

    def _create_about_group(self) -> QGroupBox:
        """
        创建关于信息组

        Returns:
            QGroupBox: 包含应用关于信息的组框
        """
        about_group = QGroupBox("ℹ️ 关于")
        about_layout = QVBoxLayout(about_group)

        # 应用说明文本
        about_text = QLabel(
            "供应商数据智能匹配系统 v1.0\n\n"
            "一个现代化的Excel数据匹配工具，\n"
            "帮助您快速处理和分析供应商数据。\n\n"
            "特性：\n"
            "• 支持拖拽上传\n"
            "• 智能日期范围处理\n"
            "• 实时统计显示\n"
            "• 简洁现代界面"
        )
        about_text.setWordWrap(True)
        about_layout.addWidget(about_text)

        return about_group

    def _on_logging_changed(self, state: int):
        """
        日志设置改变的处理函数

        Args:
            state: 复选框状态（Qt.Checked或Qt.Unchecked）
        """
        enabled = bool(state)
        self.settings.setValue('enable_logging', enabled)
        logging.getLogger().disabled = not enabled
        status = "启用" if enabled else "禁用"
        logging.info(f"日志记录已{status}")

        # 发出信号通知外部
        self.logging_toggled.emit(enabled)

    def _get_log_file_path(self) -> str:
        """
        获取日志文件的绝对路径

        Returns:
            str: 日志文件的完整路径，如果未设置则返回提示信息
        """
        if self.log_file:
            return os.path.abspath(self.log_file)
        return "未设置"

    def set_log_file(self, log_file: str):
        """
        设置日志文件路径

        Args:
            log_file: 日志文件路径
        """
        self.log_file = log_file
        # 更新显示
        if self.log_path_label:
            self.log_path_label.setText(f"日志文件位置：{self._get_log_file_path()}")

    def is_logging_enabled(self) -> bool:
        """
        获取日志记录的启用状态

        Returns:
            bool: 如果日志记录已启用返回True，否则返回False
        """
        return self.settings.value('enable_logging', False, bool)

    def get_settings(self) -> QSettings:
        """
        获取设置对象

        Returns:
            QSettings: 应用程序设置对象
        """
        return self.settings
