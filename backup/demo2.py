from openpyxl import load_workbook
import openpyxl
wb = load_workbook("example.xlsx")
ws = wb["Sheet"]
ws.insert_rows(1)  # 在第一行前插入一行
ws.insert_rows(1, 2)  # 在第一行前插入两个
ws.delete_rows(2)  # 删除第二行
ws.delete_rows(2, 2)  # 删除第二行及其后边一行（共两行）
ws.insert_cols(3)  # 在第三列前插入一列
ws.insert_cols(3, 2)  # 在第三列前插入两列
ws.delete_cols(4)  # 删除第四列
ws.delete_cols(4, 2)  # 删除第四列及其后边一列（共两列）
ws.move_range("D4:F10", rows=-1, cols=2)
wb.save("example.xlsx")
