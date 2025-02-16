import sys
from cx_Freeze import setup, Executable

# 构建选项
build_exe_options = {
    "packages": ["PySide6", "openpyxl"],
    "excludes": [],
    "include_files": [
        ("resources/icon.ico", "resources/icon.ico")
    ]
}

# 可执行文件选项
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="供应商数据智能匹配系统",
    version="1.0",
    description="供应商数据智能匹配系统",
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