"""
日志配置模块

本模块提供日志系统的初始化和配置功能。
包括日志目录管理、历史日志清理、日志格式配置等。

主要功能:
- 自动创建日志目录
- 清理历史日志文件（保留当天）
- 配置日志格式和级别
- 返回日志文件路径供其他模块使用

日志配置:
- 日志级别: DEBUG（记录所有级别的日志）
- 日志格式: 时间 - 级别 - 消息
- 文件命名: 供应商匹配_YYYYMMDD.log
- 保存策略: 每天一个日志文件，自动清理旧日志

作者: 供应商数据智能匹配系统开发团队
版本: 1.0
"""

import os
import logging
from datetime import datetime


def setup_logging(log_dir: str) -> str:
    """
    设置日志系统

    初始化应用程序的日志系统，包括创建日志目录、清理历史日志、配置日志格式等。
    确保每次运行时使用新的日志文件，便于追踪和调试。

    参数:
        log_dir (str): 日志文件存储目录的路径

    返回:
        str: 当前日志文件的完整路径

    处理流程:
        1. 检查并创建日志目录（如果不存在）
        2. 清理历史日志文件（保留当天之前的所有日志）
        3. 删除当天的旧日志文件（如果存在）
        4. 创建新的日志文件并配置日志格式
        5. 返回日志文件路径

    日志配置:
        - 级别: DEBUG（记录所有详细信息）
        - 格式: "%(asctime)s - %(levelname)s - %(message)s"
          示例: "2024-01-06 10:30:45,123 - DEBUG - 处理数据"
        - 文件名: 供应商匹配_YYYYMMDD.log
          示例: 供应商匹配_20240106.log

    错误处理:
        - 目录创建失败: 会抛出OSError异常
        - 文件删除失败: 捕获异常并打印警告，不中断程序
        - 文件权限问题: 捕获异常并打印警告信息

    示例:
        >>> # 设置日志系统
        >>> log_file = setup_logging("./logs")
        >>> print(f"日志文件: {log_file}")
        日志文件: ./logs/供应商匹配_20240106.log

        >>> # 使用日志系统
        >>> logging.info("系统启动")
        >>> logging.debug("处理数据: ...")
        >>> logging.error("发生错误: ...")

    使用建议:
        - 在程序启动时调用一次
        - 保存返回的日志文件路径供后续使用
        - 可通过QSettings控制日志的启用/禁用状态
        - 日志文件可用于问题诊断和性能分析

    注意:
        - 每次运行都会创建新的日志文件
        - 当天的旧日志会被删除，避免日志文件过大
        - 历史日志文件（非当天）会被自动清理
        - 如果需要保留历史日志，应该备份到其他位置

    日志清理策略:
        - 清理所有以'供应商匹配_'开头且以'.log'结尾的文件
        - 保留日期小于等于今天的文件
        - 这意味着只有今天之前的日志会被清理，今天的日志会先删除后重新创建

    文件命名说明:
        - 日期格式: YYYYMMDD（年月日，各两位）
        - 示例: 20240106 表示2024年1月6日
        - 使用日期后缀便于按天组织日志文件
    """
    # 步骤1: 创建日志目录（如果不存在）
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # 步骤2: 获取当前日期字符串
    today = datetime.now().strftime("%Y%m%d")

    # 步骤3: 清理历史日志文件
    try:
        # 遍历日志目录中的所有文件
        for log_file_name in os.listdir(log_dir):
            # 检查是否是系统的日志文件
            if log_file_name.startswith('供应商匹配_') and log_file_name.endswith('.log'):
                # 从文件名中提取日期
                file_date = log_file_name.replace('供应商匹配_', '').replace('.log', '')

                # 如果文件日期小于等于今天，删除该文件
                if file_date <= today:
                    old_log_path = os.path.join(log_dir, log_file_name)
                    try:
                        os.remove(old_log_path)
                        print(f"已清理历史日志: {log_file_name}")
                    except Exception as e:
                        # 删除失败时打印警告，不中断程序
                        print(f"清理日志文件失败 {log_file_name}: {str(e)}")
    except Exception as e:
        # 遍历目录失败时打印警告，不中断程序
        print(f"清理历史日志时出错: {str(e)}")

    # 步骤4: 构建当天日志文件路径
    log_file = os.path.join(log_dir, f'供应商匹配_{today}.log')

    # 步骤5: 如果当天的日志文件已存在，删除它（确保从头开始记录）
    if os.path.exists(log_file):
        try:
            os.remove(log_file)
        except Exception as e:
            # 删除失败时打印警告，继续执行
            print(f"清理旧日志文件失败: {str(e)}")

    # 步骤6: 配置日志系统
    logging.basicConfig(
        filename=log_file,          # 日志文件路径
        level=logging.DEBUG,        # 记录所有级别的日志
        format='%(asctime)s - %(levelname)s - %(message)s'  # 日志格式
    )

    # 步骤7: 返回日志文件路径供其他模块使用
    return log_file


# ==================== 使用示例 ====================

"""
使用示例:

1. 基本使用:
   ```python
   from core.logging_config import setup_logging

   # 初始化日志系统
   log_file = setup_logging("./logs")

   # 记录日志
   logging.debug("这是调试信息")
   logging.info("这是普通信息")
   logging.warning("这是警告信息")
   logging.error("这是错误信息")
   ```

2. 与QSettings配合使用:
   ```python
   from core.logging_config import setup_logging
   from PySide6.QtCore import QSettings

   # 初始化日志系统
   log_file = setup_logging("./logs")

   # 从配置中读取日志开关
   settings = QSettings('供应商数据智能匹配系统', 'DataAnalysis')
   enable_logging = settings.value('enable_logging', False, bool)

   # 根据配置启用或禁用日志
   logging.getLogger().disabled = not enable_logging

   if enable_logging:
       logging.info("日志系统已启用")
   ```

3. 记录异常信息:
   ```python
   try:
       # 一些可能出错的操作
       result = process_data()
   except Exception as e:
       # 记录异常信息和堆栈跟踪
       logging.error(f"处理数据时出错: {str(e)}", exc_info=True)
   ```

4. 记录数据处理过程:
   ```python
   logging.info("开始处理数据")
   for i, item in enumerate(data_list):
       logging.debug(f"处理第{i}项: {item}")
       # 处理逻辑...
   logging.info(f"数据处理完成，共处理{len(data_list)}项")
   ```

5. 在UI中显示日志路径:
   ```python
   from core.logging_config import setup_logging

   log_file = setup_logging("./logs")

   # 在界面上显示日志文件位置
   label = QLabel(f"日志文件: {os.path.abspath(log_file)}")
   ```

注意事项:
- 日志文件会记录所有DEBUG级别及以上的信息
- exc_info=True参数可以记录完整的异常堆栈信息
- 日志文件是文本文件，可以用任何文本编辑器打开
- 建议定期备份重要的日志文件
- 日志文件可能包含敏感信息，注意保密
"""
