import os
from pathlib import Path
import win32com.client as win32


def convert_xls_to_xlsx_win32com(folder_path):
    """
    使用win32com将xls文件转换为xlsx文件（仅Windows）
    """
    folder = Path(folder_path)
    xls_files = list(folder.glob('*.xls'))

    if not xls_files:
        print("未找到xls文件")
        return

    # 启动Excel应用程序
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False  # 不显示Excel界面

    try:
        for xls_file in xls_files:
            try:
                # 打开xls文件
                workbook = excel.Workbooks.Open(str(xls_file))

                # 生成新的文件名
                xlsx_file = xls_file.with_suffix('.xlsx')

                # 保存为xlsx格式
                workbook.SaveAs(str(xlsx_file), FileFormat=51)  # 51代表xlsx格式
                workbook.Close()
                print(f"转换成功: {xls_file.name} -> {xlsx_file.name}")

            except Exception as e:
                print(f"转换失败 {xls_file.name}: {str(e)}")

    finally:
        excel.Quit()


# 使用示例
folder_path ="D:\Desktop\新建文件夹"
convert_xls_to_xlsx_win32com(folder_path)
