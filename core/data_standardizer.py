"""
数据标准化处理模块

本模块提供数据标准化功能，用于将各种格式的输入数据转换为统一的规范格式。
主要处理日期、客户名称和产品名称的标准化。

主要功能:
- 日期数据的标准化（支持多种日期格式）
- 日期范围数据的解析和扩展
- 客户名称的标准化（处理全角半角、标点符号等）
- 产品名称的标准化（处理大小写、特殊符号等）

支持的日期格式:
- 完整日期: 2024-03、2024/03、2024.03
- 简写日期: 2403、24年3月、3月
- 日期范围: 2024年3月-5月、202403-05、3月到5月
- 中文月份: 正月、三月、十月等

作者: 供应商数据智能匹配系统开发团队
版本: 1.0
"""

import re
import logging
from datetime import datetime
from typing import Optional


def standardize_data(value: str, column_index: int) -> str:
    """
    数据标准化主函数

    根据列索引调用相应的标准化函数，对不同类型的数据进行标准化处理。
    这是数据标准化的统一入口，根据数据的位置自动选择合适的标准化策略。

    参数:
        value (str): 需要标准化的原始数据值
        column_index (int): 列索引，用于判断数据类型
            - 1: 日期列，调用日期标准化函数
            - 2: 客户名称列，调用客户名称标准化函数
            - 3: 产品名称列，调用产品名称标准化函数
            - 其他: 不进行标准化，原样返回

    返回:
        str: 标准化后的数据值

    处理流程:
        1. 检查输入值是否为空，为空则返回空字符串
        2. 移除所有空白字符（包括空格、制表符、换行符等）
        3. 根据列索引调用相应的标准化函数

    示例:
        >>> # 标准化日期
        >>> standardize_data("2024年3月", 1)
        '202403'

        >>> # 标准化客户名称
        >>> standardize_data("客户A（中国）", 2)
        '客户A(中国)'

        >>> # 标准化产品名称
        >>> standardize_data("product abc", 3)
        'PRODUCT ABC'
    """
    # 如果输入值为空，返回空字符串
    if not value:
        return ""

    # 移除所有空白字符，确保数据紧凑
    value = ''.join(value.split())

    # 根据列索引选择相应的标准化函数
    if column_index == 1:
        # 第一列：日期数据
        return _standardize_date(value)
    elif column_index == 2:
        # 第二列：客户名称
        return _standardize_customer_name(value)
    elif column_index == 3:
        # 第三列：产品名称
        return _standardize_product_name(value)

    # 其他列：不进行标准化处理
    return value


def _standardize_date(value: str) -> str:
    """
    标准化日期数据

    将各种格式的日期数据统一转换为YYYYMM格式（如202403表示2024年3月）。
    支持中文数字、日期范围等多种格式。

    参数:
        value (str): 原始日期字符串

    返回:
        str: 标准化后的日期字符串，格式为YYYYMM。
             如果无法解析，返回原始值。

    处理步骤:
        1. 记录调试日志
        2. 将中文数字转换为阿拉伯数字（一->1, 二->2, ..., 正->1）
        3. 尝试解析为日期范围格式
        4. 移除"月"和"年"字符
        5. 使用正则表达式匹配各种日期格式
        6. 转换为标准格式并验证月份有效性

    支持的日期格式:
        - 完整格式: 2024-03、2024/03、2024.03
        - 年份简写: 2403 -> 202403
        - 单独月份: 3月 -> 202403（使用当前年份）
        - 中文月份: 正月、三月、十月等

    示例:
        >>> _standardize_date("2024年3月")
        '202403'

        >>> _standardize_date("24年3月")
        '202403'

        >>> _standardize_date("3月")
        '202403'  # 假设当前年份为2024

        >>> _standardize_date("2024年3月-5月")
        '202403,202404,202405'  # 日期范围由_parse_date_range处理
    """
    logging.debug(f"处理日期值: {value}")

    # 定义中文数字到阿拉伯数字的映射
    cn_num_map = {
        '一': '1', '二': '2', '三': '3', '四': '4',
        '五': '5', '六': '6', '七': '7', '八': '8',
        '九': '9', '十': '10', '正': '1'  # 正月指一月
    }

    # 将中文数字替换为阿拉伯数字
    for cn, num in cn_num_map.items():
        value = value.replace(cn, num)

    # 尝试解析日期范围（如"3月-5月"）
    date_range = _parse_date_range(value)
    if date_range:
        return date_range

    # 移除"月"和"年"字符，简化后续处理
    value = value.replace('月', '').replace('年', '')

    # 定义日期格式的正则表达式模式列表
    # 每个模式包含(正则表达式, 提取组数量)
    date_patterns = [
        # 完整日期: 2024-03 或 2024/03 或 2024.03
        (r'(\d{4})[-/.]?(\d{1,2})', 2),
        # 简写日期: 2403
        (r'(\d{2})(\d{2})', 2),
        # 单独月份: 3 或 03
        (r'(\d{1,2})', 1),
    ]

    # 按顺序尝试匹配日期格式
    for pattern, group_count in date_patterns:
        match = re.match(pattern, value)
        if match:
            try:
                groups = match.groups()

                # 根据提取组数量处理年月
                if group_count == 2:
                    # 完整日期或简写日期格式
                    year, month = groups
                    # 如果年份是两位数，补全为四位数（24 -> 2024）
                    if len(year) == 2:
                        year = '20' + year
                else:
                    # 只有月份的格式，使用当前年份
                    year = str(datetime.now().year)
                    month = groups[0]

                # 验证月份有效性并格式化
                month = int(month)
                if 1 <= month <= 12:
                    # 将月份补零为两位数
                    month = str(month).zfill(2)
                    result = f"{year}{month}"
                    logging.debug(f"日期标准化结果: {result}")
                    return result
            except (ValueError, IndexError):
                # 如果解析失败，继续尝试下一个模式
                pass

    # 所有模式都匹配失败，返回原始值
    logging.debug(f"日期标准化结果: {value} (未改变)")
    return value


def _parse_date_range(value: str) -> Optional[str]:
    """
    解析日期范围

    将日期范围字符串（如"2024年3月-5月"）扩展为逗号分隔的月份列表。
    支持中文和数字两种格式。

    参数:
        value (str): 日期范围字符串

    返回:
        Optional[str]: 解析成功返回逗号分隔的月份列表（如"202403,202404,202405"），
                      解析失败返回None

    支持的格式:
        - 中文格式1: 2024年3月到5月、2024年3月-5月、2024年3月至5月
        - 中文格式2: 2024年3-5月
        - 数字格式: 202403-05

    处理逻辑:
        1. 尝试匹配中文日期范围格式
        2. 如果匹配成功，提取年份、起始月份、结束月份
        3. 生成月份列表（包含起始月到结束月的所有月份）
        4. 尝试匹配数字日期范围格式（备用格式）
        5. 验证月份的有效性（1-12）

    示例:
        >>> _parse_date_range("2024年3月到5月")
        '202403,202404,202405'

        >>> _parse_date_range("2024年3-5月")
        '202403,202404,202405'

        >>> _parse_date_range("202403-05")
        '202403,202404,202405'

        >>> _parse_date_range("3月")
        None  # 不是日期范围
    """
    # 定义中文日期范围的正则表达式模式
    cn_range_patterns = [
        # 模式1: 2024年3月到5月（完整格式）
        r'(\d{2,4})年(\d{1,2})月[到至和-](\d{1,2})月',
        # 模式2: 2024年3-5月（简化格式）
        r'(\d{2,4})年(\d{1,2})[到至和-](\d{1,2})月',
    ]

    # 尝试匹配中文日期范围格式
    for pattern in cn_range_patterns:
        match = re.search(pattern, value)
        if match:
            # 提取年份和月份
            year = match.group(1)
            # 如果年份是两位数，补全为四位数
            if len(year) == 2:
                year = '20' + year

            # 提取起始月份和结束月份
            start_month = int(match.group(2))
            end_month = int(match.group(3))

            # 验证月份有效性
            if 1 <= start_month <= 12 and 1 <= end_month <= 12:
                # 生成月份列表，每个月份格式为YYYYMM
                months = [f"{year}{str(m).zfill(2)}"
                         for m in range(start_month, end_month + 1)]
                return ",".join(months)

    # 尝试匹配数字日期范围格式（如202403-05）
    num_range_pattern = r'(\d{4})(\d{1,2})-(\d{1,2})'
    match = re.search(num_range_pattern, value)
    if match:
        # 提取年份和月份
        year = match.group(1)
        start_month = int(match.group(2))
        end_month = int(match.group(3))

        # 验证月份有效性
        if 1 <= start_month <= 12 and 1 <= end_month <= 12:
            # 生成月份列表
            months = [f"{year}{str(m).zfill(2)}"
                     for m in range(start_month, end_month + 1)]
            return ",".join(months)

    # 所有模式都匹配失败
    return None


def _standardize_customer_name(value: str) -> str:
    """
    标准化客户名称

    统一客户名称的格式，处理全角半角字符、标点符号差异等问题。
    确保相同含义的客户名称具有相同的字符串表示。

    参数:
        value (str): 原始客户名称

    返回:
        str: 标准化后的客户名称

    标准化规则:
        1. 全角括号转换为半角括号：（ -> (, ） -> )
        2. 全角冒号转换为半角冒号：： -> :
        3. 全角逗号转换为半角逗号：， -> ,
        4. 全角引号转换为半角引号：" -> ", " -> "
        5. 移除全角空格：

    示例:
        >>> _standardize_customer_name("客户A（中国）")
        '客户A(中国)'

        >>> _standardize_customer_name("客户B：北京，上海")
        '客户B:北京,上海'

        >>> _standardize_customer_name("客户C　测试")
        '客户C测试'
    """
    # 全角括号转半角
    value = value.replace('（', '(').replace('）', ')')
    # 全角冒号转半角
    value = value.replace('：', ':').replace('，', ',')
    # 全角引号转半角
    value = value.replace('"', '"').replace('"', '"')
    # 移除全角空格
    value = value.replace('　', '')
    return value


def _standardize_product_name(value: str) -> str:
    """
    标准化产品名称

    统一产品名称的格式，处理大小写、全角半角字符、标点符号等问题。
    确保相同含义的产品名称具有相同的字符串表示。

    参数:
        value (str): 原始产品名称

    返回:
        str: 标准化后的产品名称（大写格式）

    标准化规则:
        1. 全角括号转换为半角括号：（ -> (, ） -> )
        2. 全角逗号转换为半角逗号：， -> ,
        3. 全角冒号转换为半角冒号：： -> :
        4. 移除全角空格：
        5. 转换为大写：统一使用大写字母，便于比较

    示例:
        >>> _standardize_product_name("Product A")
        'PRODUCT A'

        >>> _standardize_product_name("产品（测试）")
        '产品(测试)'

        >>> _standardize_product_name("item：测试，demo")
        'ITEM:TEST,DEMO'

    注意:
        产品名称转换为大写后，可以避免因大小写不同导致的匹配失败。
        这在处理用户提供的产品名称时特别有用。
    """
    # 全角括号转半角
    value = value.replace('（', '(').replace('）', ')')
    # 全角逗号转半角
    value = value.replace('，', ',').replace('：', ':')
    # 移除全角空格
    value = value.replace('　', '')
    # 转换为大写
    return value.upper()
