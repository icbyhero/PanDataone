# code-organization Specification

## Purpose
TBD - created by archiving change refactor-modular-ui. Update Purpose after archive.
## Requirements
### Requirement: Modular File Structure
代码MUST按照功能模块组织到独立的文件中，每个文件SHALL NOT超过500行代码。

#### Scenario: Core business logic separation
**Given** 当前的 `main_ui_enhanced.py` 包含所有功能
**When** 开发者需要修改数据标准化逻辑
**Then** 应该能够在 `core/data_standardizer.py` 中找到相关代码，而不需要在1481行的大文件中搜索

#### Scenario: UI component isolation
**Given** 应用程序包含多个UI组件（拖放区域、统计卡片等）
**When** 开发者需要重用拖放区域组件
**Then** 应该能够独立导入 `ui.widgets.drop_zone` 模块，而不依赖主窗口

#### Scenario: File size constraints
**Given** 代码已经模块化
**When** 检查任何Python文件
**Then** 文件行数不应超过500行

---

### Requirement: Chinese Documentation
所有类、方法和公共函数MUST提供详细的中文文档字符串。

#### Scenario: Class documentation
**Given** 一个Python类（如 `StatCard`）
**When** 开发者查看类的定义
**Then** 应该看到中文文档字符串，说明类的用途、职责和使用方法

#### Scenario: Function documentation
**Given** 一个公共函数（如 `standardize_data`）
**When** 开发者需要理解函数的行为
**Then** 文档字符串应该包含：
- 功能描述
- 参数说明
- 返回值说明
- 使用示例（如果复杂）

#### Scenario: Inline comments
**Given** 代码包含复杂的业务逻辑或算法
**When** 代码阅读者查看实现细节
**Then** 关键步骤应该有中文行内注释解释其目的

---

### Requirement: Clear Module Responsibilities
每个模块MUST有单一、明确的职责，并SHALL在模块级文档中说明。

#### Scenario: Core layer modules
**Given** `core/` 目录下的模块
**When** 查看模块的职责
**Then** 模块应该：
- 不包含UI相关代码
- 不依赖PySide6或UI框架
- 专注于业务逻辑或数据处理

#### Scenario: UI layer modules
**Given** `ui/` 目录下的模块
**When** 查看模块的职责
**Then** 模块应该：
- 专注于用户界面渲染和交互
- 通过信号-槽与业务逻辑层通信
- 不直接实现数据处理算法

#### Scenario: Utility modules
**Given** `utils/` 目录下的模块
**When** 查看模块的职责
**Then** 模块应该：
- 提供通用的、可复用的工具函数
- 不依赖特定的业务逻辑
- 可以在任何层中使用

---

### Requirement: No Circular Dependencies
模块之间的依赖关系MUST是有向无环图（DAG）。

#### Scenario: Layer dependency rule
**Given** UI层需要使用核心层的功能
**When** 导入模块
**Then** UI层可以导入核心层，但核心层不能导入UI层

#### Scenario: Import validation
**Given** 所有模块已经创建
**When** 运行导入检查工具（如 `pylint` 或 `circular_imports`）
**Then** 不应该检测到任何循环依赖

#### Scenario: Module independence
**Given** 核心层的任何模块
**When** 在没有UI层的情况下导入
**Then** 模块应该能够成功导入和使用

---

### Requirement: Backward Compatibility
重构后的应用程序MUST保持所有现有功能和用户界面不变。

#### Scenario: Feature preservation
**Given** 当前应用程序的所有功能（文件选择、拖放、数据分析、统计显示等）
**When** 重构完成后运行应用程序
**Then** 所有功能应该与重构前完全一致

#### Scenario: UI consistency
**Given** 当前应用程序的外观和行为
**When** 用户使用重构后的应用程序
**Then** 应该无法察觉任何差异（除了代码组织）

#### Scenario: Configuration compatibility
**Given** 用户使用QSettings保存的配置
**When** 重构后的应用程序读取配置
**Then** 应该能够正确读取和使用所有现有配置

---

### Requirement: Import Path Organization
模块导入SHALL遵循Python最佳实践，MUST使用清晰的相对和绝对导入。

#### Scenario: Package-relative imports
**Given** 同一包内的模块需要互相导入
**When** 编写导入语句
**Then** 应该使用相对导入（如 `from ..widgets import drop_zone`）

#### Scenario: Entry point imports
**Given** `main_ui_enhanced.py` 作为应用程序入口
**When** 导入其他模块
**Then** 应该使用绝对导入（如 `from ui.main_window import MainWindow`）

#### Scenario: Explicit exports
**Given** 一个模块定义了公共API
**When** 其他模块需要导入
**Then** 模块应该定义 `__all__` 列表明确导出的符号

---

