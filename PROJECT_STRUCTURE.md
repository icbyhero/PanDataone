# PanDataone 项目结构

## 重构后的模块化架构

```
PanDataone/
│
├── main.py (61行)
│   └── 应用程序入口点
│
├── core/                          # 核心业务逻辑模块
│   ├── __init__.py (22行)
│   ├── data_models.py (174行)     # MatchResult, CellStyle, CellStyles
│   ├── data_standardizer.py (366行) # 数据标准化函数
│   ├── excel_processor.py (218行)  # Excel文件处理
│   └── logging_config.py (217行)   # 日志系统配置
│
├── ui/                            # UI组件模块
│   ├── __init__.py (13行)
│   ├── main_window.py (927行)     # 主窗口类
│   ├── styles.py (273行)          # 全局样式定义
│   │
│   ├── widgets/                   # 可复用UI组件
│   │   ├── __init__.py (15行)
│   │   ├── drop_zone.py (298行)  # 拖放区域组件
│   │   ├── stat_card.py (172行)  # 统计卡片组件
│   │   └── help_widget.py (180行) # 帮助组件
│   │
│   └── tabs/                      # 标签页组件
│       ├── __init__.py (13行)
│       ├── filter_tab.py (435行)  # 数据筛选标签页
│       └── settings_tab.py (165行) # 设置标签页
│
└── utils/                         # 工具函数模块
    └── __init__.py (7行)          # 预留扩展
```

## 文件统计

### 核心模块 (core/)
- **总行数**: 997行
- **文件数**: 5个
- **平均**: 199行/文件

### UI主模块 (ui/)
- **总行数**: 1213行
- **文件数**: 3个
- **平均**: 404行/文件

### UI组件 (ui/widgets/)
- **总行数**: 665行
- **文件数**: 4个
- **平均**: 166行/文件

### UI标签页 (ui/tabs/)
- **总行数**: 613行
- **文件数**: 3个
- **平均**: 204行/文件

### 主入口 (main.py)
- **行数**: 61行
- **文件数**: 1个

## 总计

- **总文件数**: 16个Python文件
- **总代码行数**: 3549行
- **平均文件大小**: 222行/文件

## 备份文件

- `main_backup.py` (798行) - 原始main.py备份
- `main_ui_enhanced_backup.py` (1481行) - 原始main_ui_enhanced.py备份

## 设计原则

1. **单一职责**: 每个模块专注于特定功能
2. **低耦合**: 模块间依赖最小化
3. **高内聚**: 相关功能聚合在同一模块
4. **可测试性**: 独立模块便于单元测试
5. **可维护性**: 文件大小合理，易于理解
6. **可重用性**: UI组件可在其他项目中使用

## 导入示例

```python
# 导入核心模块
from core import MatchResult, CellStyle, standardize_data
from core import get_sheet_data, clear_sheet, setup_logging

# 导入UI组件
from ui import MainWindow, apply_app_style
from ui.widgets import DropZoneGroupBox, StatCard, HelpWidget
from ui.tabs import FilterTab, SettingsTab
```
