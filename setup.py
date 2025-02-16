import sys
from cx_Freeze import setup, Executable

# 基础配置
build_exe_options = {
    "packages": ["PySide6", "openpyxl"],
    "include_files": [
        ("resources/icon.ico", "resources/icon.ico")
    ],
    "excludes": []
}

# 创建执行文件
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="供应商数据智能匹配系统",
    version="1.0",
    description="供应商数据智能匹配分析工具",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "main.py",
            base=base,
            target_name="供应商数据智能匹配系统.exe",
            icon="resources/icon.ico"
        )
    ]
)