import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def append_df_to_excel(df, filename, sheet_name='Sheet1', startrow=None, truncate_sheet=False, **to_excel_kwargs):
    """
    将DataFrame追加到Excel文件中的指定工作表。
    
    :param df: 要追加的DataFrame
    :param filename: 目标Excel文件的路径
    :param sheet_name: 要追加数据的工作表名称
    :param startrow: 开始写入数据的行号（如果为None，则从最后一行开始）
    :param truncate_sheet: 是否在写入前截断工作表
    :param to_excel_kwargs: 传递给to_excel方法的其他参数
    """
    
    # 如果文件不存在，直接写入新文件
    try:
        wb = load_workbook(filename)
    except FileNotFoundError:
        df.to_excel(filename, sheet_name=sheet_name, **to_excel_kwargs)
        return
    
    # 打开目标工作表
    if sheet_name not in wb.sheetnames:
        wb.create_sheet(sheet_name)
    ws = wb[sheet_name]
    
    if truncate_sheet and ws.max_row > 0:
        ws.delete_rows(1, ws.max_row)
    
    if startrow is None:
        startrow = ws.max_row
    
    # 写入数据
    rows = dataframe_to_rows(df, index=False, header=False)
    for r_idx, row in enumerate(rows, startrow + 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    
    wb.save(filename)