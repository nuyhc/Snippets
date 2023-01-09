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
