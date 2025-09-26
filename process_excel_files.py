import pandas as pd
import xlwings as xw
from pathlib import Path
import time
import re


def process_excel_files_with_pivot(folder_path):
    """
    使用xlwings处理Excel文件，并为每个文件创建数据透视表
    1. 处理原始数据
       -处理文件夹中所有xlsx文件（排除00汇总表.xlsx）
       -删除第一行
       -在W列标题写"计数用"，N列不为空时W列填充1
       -按P列筛选出值为空的行
       -将筛选出来的行的Q列填充"一直离线"
    2. 新建工作表并创建数据透视表
       - 行字段：小区
       - 列字段：业务编码
       - 值字段：计数用（求和）
    """
    folder = Path(folder_path)

    if not folder.exists():
        print(f"错误: 文件夹 '{folder_path}' 不存在")
        return

    # 查找所有xlsx文件，排除00汇总表.xlsx
    xlsx_files = [f for f in folder.glob('*.xlsx') if f.name != '00汇总表.xlsx']

    if not xlsx_files:
        print("未找到需要处理的xlsx文件")
        return

    print(f"找到 {len(xlsx_files)} 个需要处理的xlsx文件")

    # 启动Excel应用程序
    app = xw.App(visible=False)

    for file_path in xlsx_files:
        try:
            print(f"正在处理: {file_path.name}")

            # 打开工作簿
            wb = app.books.open(str(file_path))
            sht = wb.sheets[0]  # 选择第一个工作表

            # 获取已使用的范围
            used_range = sht.used_range
            max_row = used_range.last_cell.row
            max_col = used_range.last_cell.column

            # 1. 删除第一行（如果不止一行数据）
            if max_row > 1:
                sht.range('1:1').api.Delete()
                max_row -= 1  # 行数减少1
                print("  已删除第一行")

            # 更新使用范围（因为删除了行）
            used_range = sht.used_range
            max_row = used_range.last_cell.row
            max_col = used_range.last_cell.column

            # 2. 处理W列和N列
            # 确保有W列（第23列）
            if max_col < 23:
                # 添加列直到W列
                for col in range(max_col + 1, 24):
                    col_letter = get_column_letter(col)
                    sht.range(f"{col_letter}1").value = f"Unnamed_{col - 1}"
                max_col = 23

            # 设置W列标题
            sht.range('W1').value = "计数用"

            # 处理N列和W列的数据
            if max_col >= 14 and max_row > 1:  # 确保有N列和数据
                for row in range(2, max_row + 1):  # 从第2行开始（第1行是标题）
                    n_value = sht.range(f'N{row}').value
                    if n_value is not None and str(n_value).strip() != '':
                        sht.range(f'W{row}').value = 1
                print("  已设置W列")

            # 3. 处理P列和Q列
            if max_col >= 16 and max_row > 1:  # 确保有P列和数据
                p_empty_count = 0
                for row in range(2, max_row + 1):
                    p_value = sht.range(f'P{row}').value
                    if p_value is None or str(p_value).strip() == '':
                        sht.range(f'Q{row}').value = "一直离线"
                        p_empty_count += 1
                print(f"  已处理 {p_empty_count} 行P列为空的数据")

            # 4. 创建数据透视表
            create_pivot_table(wb, sht, max_row, max_col)

            # 保存并关闭工作簿
            wb.save()
            wb.close()
            print(f"  已完成处理: {file_path.name}\n")

        except Exception as e:
            print(f"处理文件 {file_path.name} 时出错: {str(e)}\n")
            try:
                wb.close()
            except:
                pass

    # 关闭Excel应用程序
    app.quit()
    print("批量处理完成! 所有文件已添加数据透视表。")


def create_pivot_table(wb, source_sheet, max_row, max_col):
    """
    创建数据透视表（优化版本）
    行字段：小区
    列字段：业务编码
    值字段：计数用 (求和)
    """
    try:
        print("  开始创建数据透视表...")

        # 获取数据源的所有标题行，用于精确匹配字段
        headers = source_sheet.range(f"A1:{get_column_letter(max_col)}1").value
        print(f"  数据源标题行: {headers}")

        # 查找列索引（使用模糊匹配）
        xiaoqu_col_index = find_column_index(headers, ["小区", "区域", "片区"])
        yewu_col_index = find_column_index(headers, ["业务编码", "业务代码", "业务编号"])
        jishu_col_index = find_column_index(headers, ["计数用", "计数", "数量"])

        # 检查必要列是否存在
        if xiaoqu_col_index is None:
            print("  错误：未找到行字段('小区'或类似列)，跳过创建数据透视表。")
            return

        if yewu_col_index is None:
            print("  错误：未找到列字段('业务编码'或类似列)，跳过创建数据透视表。")
            return

        if jishu_col_index is None:
            print("  警告：未找到值字段'计数用'，尝试使用W列。")
            jishu_col_index = 23  # 默认W列

        # 添加新工作表
        try:
            # 检查是否已存在"数据透视表"工作表
            try:
                existing_sheet = wb.sheets["数据透视表"]
                existing_sheet.delete()
            except:
                pass  # 工作表不存在，继续创建

            pivot_sheet = wb.sheets.add(name="数据透视表", after=source_sheet)
        except Exception as e:
            print(f"  创建新工作表时出错: {e}")
            return

        # 定义数据源范围
        data_range = source_sheet.range(f"A1:{get_column_letter(max_col)}{max_row}")

        # 添加延迟，确保Excel完全准备就绪
        time.sleep(1)

        try:
            # 创建PivotCache
            pivot_cache = wb.api.PivotCaches().Create(
                SourceType=1,  # xlDatabase
                SourceData=data_range.api
            )

            # 创建PivotTable
            pivot_table = pivot_cache.CreatePivotTable(
                TableDestination=pivot_sheet.range("A3").api,
                TableName="DataPivotTable"
            )
        except Exception as e:
            print(f"  创建数据透视表对象时出错: {e}")
            # 尝试备用方法
            try:
                pivot_table = pivot_sheet.api.PivotTableWizard(
                    SourceType=1,
                    SourceData=data_range.api,
                    TableDestination=pivot_sheet.range("A3").api,
                    TableName="DataPivotTable"
                )
            except Exception as e2:
                print(f"  备用创建方法也失败: {e2}")
                return

        # 添加延迟，确保数据透视表对象完全初始化
        time.sleep(0.5)

        # 安全地配置数据透视表字段
        # 获取列名（使用找到的索引对应的实际标题）
        xiaoqu_field_name = headers[xiaoqu_col_index - 1] if xiaoqu_col_index <= len(headers) else "小区"
        yewu_field_name = headers[yewu_col_index - 1] if yewu_col_index <= len(headers) else "业务编码"
        jishu_field_name = headers[jishu_col_index - 1] if jishu_col_index <= len(headers) else "计数用"

        print(f"  使用字段 - 行: {xiaoqu_field_name}, 列: {yewu_field_name}, 值: {jishu_field_name}")

        # 设置行字段（小区）
        try:
            pivot_table.PivotFields(xiaoqu_field_name).Orientation = 1  # xlRowField
            print(f"  已添加行字段: {xiaoqu_field_name}")
        except Exception as e:
            print(f"  添加行字段时出错: {e}")
            # 尝试使用列字母
            try:
                pivot_table.PivotFields(get_column_letter(xiaoqu_col_index)).Orientation = 1
                print(f"  已添加行字段(使用列字母): {get_column_letter(xiaoqu_col_index)}")
            except Exception as e2:
                print(f"  使用列字母添加行字段也失败: {e2}")

        # 设置列字段（业务编码）
        try:
            pivot_table.PivotFields(yewu_field_name).Orientation = 2  # xlColumnField
            print(f"  已添加列字段: {yewu_field_name}")
        except Exception as e:
            print(f"  添加列字段时出错: {e}")
            # 尝试使用列字母
            try:
                pivot_table.PivotFields(get_column_letter(yewu_col_index)).Orientation = 2
                print(f"  已添加列字段(使用列字母): {get_column_letter(yewu_col_index)}")
            except Exception as e2:
                print(f"  使用列字母添加列字段也失败: {e2}")

        # 设置值字段（计数用）- 求和
        try:
            # 先尝试标准方法
            pivot_table.AddDataField(
                pivot_table.PivotFields(jishu_field_name),
                "计数总和",
                -4157  # xlSum
            )
            print(f"  已添加值字段: {jishu_field_name}")
        except Exception as e:
            print(f"  添加值字段时出错: {e}")
            # 尝试备用方法
            try:
                pivot_table.PivotFields(jishu_field_name).Orientation = 4  # xlDataField
                pivot_table.DataPivotField.Function = -4157  # xlSum
                print(f"  已添加值字段(备用方法): {jishu_field_name}")
            except Exception as e2:
                print(f"  备用方法添加值字段也失败: {e2}")

        # 设置数据透视表样式
        try:
            pivot_table.TableStyle2 = "PivotStyleMedium9"
        except:
            pass  # 样式设置失败不影响功能

        # 添加标题
        pivot_sheet.range("A1").value = "数据透视表"
        pivot_sheet.range("A1").font.bold = True
        pivot_sheet.range("A1").font.size = 14

        # 调整列宽以适应内容
        pivot_sheet.autofit()

        print("  数据透视表创建成功！")

    except Exception as e:
        print(f"  创建数据透视表过程中出错: {str(e)}")


def find_column_index(headers, possible_names):
    """
    在标题行中查找目标列名的索引（1-based）
    支持多个可能的列名和模糊匹配
    """
    if not headers:
        return None

    for i, header in enumerate(headers):
        if header is None:
            continue

        header_str = str(header).strip()
        for name in possible_names:
            # 精确匹配或部分匹配
            if name in header_str or header_str in name:
                return i + 1  # 返回1-based索引

    return None


def get_column_letter(col_num):
    """将列号转换为字母表示（如1->A, 27->AA）"""
    if col_num <= 0:
        return "A"

    letter = ''
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter


# 使用示例
if __name__ == "__main__":
    folder_path = input("请输入包含Excel文件的文件夹路径: ").strip()
    process_excel_files_with_pivot(folder_path)