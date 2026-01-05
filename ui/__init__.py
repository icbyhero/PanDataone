"""
UI组件模块

包含主窗口、样式定义、自定义组件和标签页。
"""

from .main_window import MainWindow
from .styles import apply_app_style

__all__ = [
    'MainWindow',
    'apply_app_style',
]
