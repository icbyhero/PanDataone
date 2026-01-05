# Design Document: Code Modularization

## Architecture Overview

### Current State
```
main.py (798 lines) - 基础版本
├── Data Classes
├── Basic UI
└── Simple data processing

main_ui_enhanced.py (1481 lines) - 增强版本
├── Data Classes (MatchResult, CellStyle, CellStyles)
├── Utility Functions (data standardization, Excel processing)
├── UI Components (DropZoneGroupBox, StatCard)
└── MainWindow (application logic, UI composition)
```

**注意**：旧的 `DropZoneWidget` 已被 `DropZoneGroupBox` 替代，不再使用。

**问题**：两个main文件造成版本混乱，用户不知道应该运行哪个。

### Target State
```
pan_dataone/
├── main.py                       # 新主入口，替换旧main.py (~50 lines)
├── main_backup.py               # 旧main.py备份
├── main_ui_enhanced_backup.py   # 旧增强版备份
├── core/                         # Business logic layer
│   ├── data_models.py           # Data structures (~100 lines)
│   ├── data_standardizer.py     # Data normalization (~150 lines)
│   ├── excel_processor.py       # Excel operations (~150 lines)
│   └── logging_config.py        # Logging setup (~80 lines)
├── ui/                           # Presentation layer
│   ├── main_window.py           # Main window coordinator (~300 lines)
│   ├── styles.py                # UI styling (~200 lines)
│   ├── widgets/                 # Reusable UI components
│   │   ├── drop_zone.py         # File drag-drop (~300 lines)
│   │   ├── stat_card.py         # Statistics display (~150 lines)
│   │   └── help_widget.py       # Help messages (~100 lines)
│   └── tabs/                    # Page-level components
│       ├── filter_tab.py        # Data analysis page (~400 lines)
│       └── settings_tab.py      # Settings page (~100 lines)
└── utils/                        # Shared utilities
    └── constants.py             # Constants (~100 lines)
```

**改进**：
- 统一为单一入口 `main.py`
- 增强版功能模块化
- 清晰的文件职责划分

## Layer Architecture

### 1. Core Layer (业务逻辑层)

**Responsibilities**:
- 数据模型定义
- 业务逻辑实现
- Excel文件处理
- 数据标准化和转换

**Dependencies**:
- OpenPyXL (Excel操作)
- Python standard library (re, datetime, logging)

**No dependencies on**:
- UI framework (PySide6)

### 2. UI Layer (表现层)

**Responsibilities**:
- 用户界面渲染
- 用户交互处理
- 组件组合和布局

**Dependencies**:
- PySide6 (Qt框架)
- Core layer (业务逻辑调用)

**Design Principles**:
- 单一职责：每个组件只负责一个UI功能
- 可复用性：组件可以在不同上下文中使用
- 松耦合：通过信号-槽机制通信

### 3. Utils Layer (工具层)

**Responsibilities**:
- 常量定义
- 通用工具函数
- 配置管理

**Dependencies**:
- 无依赖或仅依赖标准库

## Module Design

### core/data_models.py

**Purpose**: 定义数据结构

**Classes**:
```python
@dataclass
class MatchResult:
    """匹配结果数据类"""
    # 存储匹配的数据行及其样式信息

@dataclass
class CellStyle:
    """单元格样式数据类"""
    # 定义Excel单元格的格式样式

class CellStyles:
    """样式集合类"""
    # 管理不同类型的样式定义
```

**Design Rationale**:
- 使用 `@dataclass` 减少样板代码
- 集中定义数据结构，便于维护
- 类型注解提高IDE支持

### core/data_standardizer.py

**Purpose**: 数据标准化和转换

**Functions**:
```python
def standardize_data(value: str, column_index: int) -> str:
    """主入口：根据列索引选择合适的标准化方法"""

def _standardize_date(value: str) -> str:
    """日期数据标准化"""

def _parse_date_range(value: str) -> Optional[str]:
    """解析日期范围字符串"""

def _standardize_customer_name(value: str) -> str:
    """客户名称标准化"""

def _standardize_product_name(value: str) -> str:
    """产品名称标准化"""
```

**Design Rationale**:
- 函数式设计，无状态，易于测试
- 私有函数（前缀 `_`）封装实现细节
- 清晰的单一职责

### core/excel_processor.py

**Purpose**: Excel文件操作

**Functions**:
```python
def get_sheet_data(sheet, row: int) -> Tuple[str, str, str]:
    """获取工作表指定行的数据"""

def clear_sheet(sheet) -> None:
    """清空工作表内容"""

def copy_title_row(source_sheet, target_sheet) -> None:
    """复制标题行"""

def init_result_sheet(workbook, sheet_name: str):
    """初始化结果工作表"""
```

**Design Rationale**:
- 与OpenPyXL的接口封装
- 错误处理集中管理
- 返回类型明确

### ui/widgets/drop_zone.py

**Purpose**: 文件拖放区域组件

**Class**:
```python
class DropZoneGroupBox(QGroupBox):
    """增强的拖放区域组件（整个卡片可拖放）"""
```

**Signals**:
- `file_selected = Signal(str)` - 文件选择完成

**Design Rationale**:
- 继承Qt容器类以支持拖放
- 信号机制实现松耦合通信
- 整个卡片区域都可拖放，用户体验更好
- **已弃用**：旧的 `DropZoneWidget` 不再迁移

### ui/widgets/stat_card.py

**Purpose**: 统计数据展示卡片

**Classes**:
```python
class StatCard(QFrame):
    """统计卡片组件

    显示标题、数值、图标和描述信息
    """
```

**Features**:
- 响应式布局
- 描述文字自动换行
- 可配置的图标和颜色

### ui/tabs/filter_tab.py

**Purpose**: 数据筛选和分析页面

**Class**:
```python
class FilterTab(QWidget):
    """数据筛选标签页

    包含：
    - 文件选择区域
    - 统计卡片展示
    - 进度条
    - 开始分析按钮
    """
```

**Signals**:
- `analysis_started = Signal()` - 分析开始
- `analysis_finished = Signal(results)` - 分析完成

**Design Rationale**:
- 将整个页面封装为类
- 通过信号与主窗口通信
- 内部管理子组件生命周期

### ui/main_window.py

**Purpose**: 主窗口协调器

**Class**:
```python
class MainWindow(QMainWindow):
    """主窗口

    职责：
    - 创建和管理标签页
    - 协调各组件间的交互
    - 管理应用程序状态
    - 处理全局事件
    """
```

**Composition**:
```
MainWindow
├── QTabWidget
│   ├── FilterTab
│   └── SettingsTab
├── Status Bar (optional)
└── Menu Bar (optional)
```

## Communication Patterns

### Signal-Slot Pattern (Qt)

**Example**:
```python
# In DropZoneGroupBox
file_selected = Signal(str)

# In FilterTab
drop_zone.file_selected.connect(self._on_file_selected)

# In MainWindow
filter_tab.analysis_started.connect(self._on_analysis_started)
```

**Benefits**:
- 松耦合
- 易于扩展
- 符合Qt框架规范

### Dependency Injection

**Example**:
```python
class FilterTab(QWidget):
    def __init__(self, data_processor: DataProcessor, parent=None):
        super().__init__(parent)
        self.processor = data_processor
        # ...
```

**Benefits**:
- 可测试性
- 灵活性
- 依赖关系清晰

## Import Strategy

### Relative Imports (within package)
```python
# In ui/tabs/filter_tab.py
from ..widgets.drop_zone import DropZoneGroupBox
from ..widgets.stat_card import StatCard
from ...core.data_standardizer import standardize_data
```

### Absolute Imports (from entry point)
```python
# In main_ui_enhanced.py
from ui.main_window import MainWindow
from core.logging_config import setup_logging
```

## Error Handling Strategy

### Core Layer
```python
def process_excel(file_path: str) -> Result:
    try:
        # 处理逻辑
        return Result(success=True, data=...)
    except FileNotFoundError as e:
        logger.error(f"文件不存在: {file_path}")
        return Result(success=False, error=str(e))
```

### UI Layer
```python
def browse_file(self):
    try:
        file_path, _ = QFileDialog.getOpenFileName(...)
        if file_path:
            self.file_selected.emit(file_path)
    except Exception as e:
        QMessageBox.critical(self, "错误", f"选择文件失败: {str(e)}")
```

## Testing Strategy

### Unit Tests (Future)
```python
# tests/test_data_standardizer.py
def test_standardize_date():
    assert _standardize_date("2024年1月1日") == "2024-01-01"

def test_parse_date_range():
    assert _parse_date_range("2024.01-2024.03") == "2024年01月-03月"
```

### Integration Tests (Future)
```python
# tests/test_excel_processor.py
def test_full_workflow():
    workbook = load_test_workbook()
    result = process_data_analysis(workbook)
    assert result.matched_count > 0
```

## Performance Considerations

### File Size Targets
- 每个文件 < 500 行
- 每个类 < 300 行
- 每个函数 < 50 行

### Import Optimization
- 延迟导入非必需模块
- 使用 `__all__` 明确导出
- 避免循环导入

## Migration Risks

### Risk 1: Import Errors
**Mitigation**:
- 分阶段迁移
- 每阶段验证运行
- 使用 `try-except` 捕获导入错误

### Risk 2: Circular Dependencies
**Mitigation**:
- 设计阶段绘制依赖图
- 核心层不依赖UI层
- 使用信号机制解耦

### Risk 3: Runtime Errors
**Mitigation**:
- 完整的回归测试清单
- 功能对比验证
- 保持备份可用

## Success Metrics

### Code Quality
- [ ] 每个文件有明确的职责
- [ ] 所有公共API有文档字符串
- [ ] 关键逻辑有行内注释
- [ ] 无循环依赖
- [ ] 所有模块可独立导入

### Maintainability
- [ ] 新功能易于添加
- [ ] Bug易于定位和修复
- [ ] 代码易于理解
- [ ] 模块易于测试

### Functionality
- [ ] 所有现有功能正常工作
- [ ] UI外观无变化
- [ ] 性能无明显下降
- [ ] 无新的错误或警告
