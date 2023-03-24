import time
import pythoncom
import pandas as pd
import win32com.client
from typing import TypeVar

Excel = TypeVar("Excel")
WorkBook = TypeVar("WorkBook")
WorkSheet = TypeVar("WorkSheet")

# 엑셀 제어 객체 생성 / 삭제
def init_win32com_excel()->Excel:
    pythoncom.CoInitialize()
    time.sleep(2)
    excel = win32com.client.Dispatch("Excel.Application")
    # 작업 중, 엑셀 출력 여부
    excel.Visible = False
    # 이벤트 처리 출력 여부
    excel.EnableEvents = False
    # 경고 출력 여부
    excel.DisplayAlerts = False
    return excel

def quit_win32com_excel(excel: Excel)->None:
    excel.Quit()
    
# Excel data -> pd.DataFrame
# Task에 맞게 커스텀 필요
def XLSX_to_DF(workbook: WorkBook, sheet_name: str=None, isdrop=False)->pd.DataFrame:
    # 시트 설정
    if sheet_name==None: ws = workbook.ActiveSheet
    else: ws = workbook.Sheets[sheet_name]
    # xlsx -> dataframe
    df_temp = pd.DataFrame(ws.UsedRange.Rows.Value)
    # 일반적으로, 첫 행에 컬럼명이 들어오지만, 확인 후 변경이 필요
    df_temp = df_temp.rename(columns=df_temp.iloc[0])
    if isdrop: df_temp = df_temp.drop(df_temp.index[0])
    return df_temp.reset_index(drop=True)

# 병합된 셀에 대해, 같은 내용을 갖고 데이터프레임 화
def Merged_XLSX_to_DF(workbook: WorkBook)->pd.DataFrame:
    ws = workbook.ActiveSheet
    data = []
    for row in range(1, ws.UsedRange.Rows.Count+1):
        row_data = []
        for col in range(1, ws.UsedRange.Columns.Count+1):
            cell = ws.Cells(row, col)
            if cell.MergeCells:
                merge_area = cell.MergeArea
                row_data.append(merge_area.Value) # 병합된 셀의 갯수만큼 튜퓰 형식으로 나옴
            else:
                row_data.append(cell.Value)
        data.append(row_data)
    return pd.DataFrame(data)
