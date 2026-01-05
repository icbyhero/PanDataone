"""
统计卡片组件模块

该模块提供了用于显示统计信息的卡片组件。
支持自定义标题、数值、图标和描述说明，具有响应式布局特性。

主要功能:
- 显示统计数值和图标
- 支持可选的描述说明（会自动扩展）
- 响应式布局，自适应容器大小
- 悬停效果增强交互体验

使用示例:
    card = StatCard("总数据", "100", "📋", "待处理的数据总数")
    card.update_value("200")  # 更新数值
    layout.addWidget(card)
"""

from PySide6.QtWidgets import QFrame, QLabel, QVBoxLayout, QHBoxLayout, QSizePolicy
from PySide6.QtCore import Qt


class StatCard(QFrame):
    """
    统计卡片组件

    用于显示统计信息的卡片组件，采用现代化的卡片式设计。
    支持显示标题、数值、图标和可选的描述说明。

    布局特点:
        - 顶部行: 图标(左) + 数值(右对齐)
        - 中部: 标题文本
        - 底部: 描述说明(可选，占据剩余空间)

    样式特点:
        - 白色背景，圆角边框
        - 悬停时背景变为浅灰色
        - 数值使用大号绿色字体
        - 描述使用蓝色背景高亮显示

    响应式设计:
        - 水平方向: 自动扩展填充容器
        - 垂直方向: 自动扩展填充容器
        - 描述文字自动换行
    """

    def __init__(self, title: str, value: str = "0", icon: str = "📊",
                 description: str = "", parent=None):
        """
        初始化统计卡片组件

        Args:
            title: 卡片标题(如"总数据"、"已匹配"等)
            value: 初始数值，默认为"0"
            icon: 图标(emoji或字符)，默认为"📊"
            description: 描述说明文字，默认为空字符串
            parent: 父窗口对象，默认为None
        """
        super().__init__(parent)
        self.title = title
        self.value = value
        self.icon = icon
        self.description = description
        self._setup_ui()

    def _setup_ui(self):
        """
        设置用户界面

        创建卡片的视觉元素:
        1. 顶部行: 图标和数值（水平布局）
        2. 标题标签
        3. 描述标签（如果提供了描述文字）

        样式设置:
        - 卡片边框和圆角
        - 悬停效果
        - 数值大号字体
        - 描述高亮背景
        """
        # 设置卡片样式
        self.setFrameStyle(QFrame.Box)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setStyleSheet("""
            QFrame {
                border: 1px solid #E0E0E0;
                border-radius: 8px;
                background-color: white;
                padding: 15px;
            }
            QFrame:hover {
                background-color: #FAFAFA;
            }
        """)

        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # 图标和数值 - 顶部行
        header_layout = QHBoxLayout()

        icon_label = QLabel(self.icon)
        icon_label.setStyleSheet("font-size: 24px;")

        value_label = QLabel(self.value)
        value_label.setObjectName("value_label")
        value_label.setStyleSheet("""
            QLabel#value_label {
                font-size: 28px;
                font-weight: bold;
                color: #4CAF50;
            }
        """)
        value_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        header_layout.addWidget(icon_label)
        header_layout.addStretch()
        header_layout.addWidget(value_label)

        layout.addLayout(header_layout)

        # 标题 - 中部
        title_label = QLabel(self.title)
        title_label.setStyleSheet("""
            QLabel {
                font-size: 13px;
                font-weight: bold;
                color: #546E7A;
            }
        """)
        layout.addWidget(title_label)

        # 描述说明 - 底部（如果提供）
        if self.description:
            desc_label = QLabel(self.description)
            desc_label.setWordWrap(True)  # 自动换行
            desc_label.setAlignment(Qt.AlignLeft)
            desc_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            desc_label.setStyleSheet("""
                QLabel {
                    font-size: 16px;
                    font-weight: bold;
                    color: #0D47A1;
                    background-color: #E3F2FD;
                    padding: 15px;
                    border-radius: 8px;
                    border: 2px solid #BBDEFB;
                }
            """)
            layout.addWidget(desc_label, 1)  # stretch=1 让描述占据剩余空间

            # 同时添加工具提示
            self.setToolTip(self.description)

        self.setLayout(layout)

    def update_value(self, value: str):
        """
        更新卡片显示的数值

        用于动态更新统计数值，例如数据分析完成后更新匹配结果。

        Args:
            value: 新的数值字符串

        使用示例:
            card.update_value("150")
        """
        value_label = self.findChild(QLabel, "value_label")
        if value_label:
            value_label.setText(value)
