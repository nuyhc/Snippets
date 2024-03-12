from typing import TypeVar

Excel = TypeVar("Excel")
WorkBook = TypeVar("WorkBook")
WorkSheet = TypeVar("WorkSheet")
Session = TypeVar("Session")

# Logging
import logging
logging.basicConfig(
    handlers=[
        logging.FileHandler(
            filename="파일 경로.log",
            encoding="utf-8", mode="a+"
        )
    ],
    format="%(asctime)s:%(message)s",
    level=logging.INFO,
    datefmt="%I:%M:%S"
)

def logged(pos: str=" ")->None:
    def wrapup(func, *args, **kwargs):
        logger = logging.getLogger()
        def loggingFunc(*args, **kwargs):
            logger.info(f"[{pos}] Call {func.__name}\nargs: {[x for x in args if type(x) not in ["제외할 자료형"]]}\n")
            return func(*args, **kwargs)
        return loggingFunc
    return wrapup

# win32com
import os
import time
import pythoncom
import win32com.client
import pandas as pd
from datetime import datetime

class win32comExcel:
    @staticmethod
    def coi()->None:
        pythoncom.CoInitialize()
        time.sleep(3)
    @staticmethod
    def _assign()->Excel:
        win32comExcel.coi()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.EnableEvents = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        return excel
    @staticmethod
    def _release(excel: Excel)->None:
        excel.Quit()
    @staticmethod
    def _kill()->None:
        try: os.system("taskkill /f /im EXCEL.exe")
        except: pass
    @staticmethod
    class Encoder:
        @staticmethod
        def I2L(col_idx: int)->str:
            letter = ""
            while col_idx>0:
                col_idx, remain = divmod(col_idx-1, 26)
                letter = chr(remain+ord("A")) + letter
            return letter
        @staticmethod
        def L2I(col_letter: str)->int:
            idx = 0
            for letter in col_letter:
                idx = (idx*26) + 1 + ord(letter) - ord("A")
            return idx
    @staticmethod
    def RGB2INT(rgb: tuple)->int:
        return rgb[0] + (rgb[1]*256) + (rgb[2]*256*256)
    @staticmethod
    def break_links(workbook: WorkBook)->None:
        try: [workbook.BreakLink(x, Type=1) for x in workbook.LinkSources()]
        except: pass
    @staticmethod
    def extract_contain_merged_cells(worksheet: WorkSheet)->list:
        val, info = [], []
        for r in range(1, worksheet.UsedRange.Rows.Count+1):
            r_val, r_info = [], []
            for c in range(1, worksheet.UsedRange.Columns.Count+1):
                cell = worksheet.Cells(r, c)
                if cell.MergeCells: temp_val = (cell.MergeArea).Value[0][0]
                else: temp_val = cell.Value
                if isinstance(temp_val, datetime): temp_val = str(temp_val)
                r_val.append(temp_val)
                r_info.append(f"{win32comExcel.Encoder.I2L(c)}{r}")
            val.append(r_val)
            info.append(r_info)
        return pd.DataFrame(val), pd.DataFrame(info)
    
class win32comOutlook:
    @staticmethod
    def style()->str:
        return \
            """
            <style>
                p{
                    font-size:13px;
                    font-family:"Malgun Gothic";
                }
            </style>
            """
    @staticmethod
    class Send:
        def __init__(self, rTO: str, rCC: str, subjct: str, body: str, attach: str=None)->None:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            if rTO!=None: mail.To = rTO
            if rCC!=None: mail.CC = rCC
            if attach!=None:
                if type(attach)==str: mail.Attachments.Add(attach)
                else: [mail.Attachmetns.Add(x) for x in attach]
            mail.Subject = subjct
            mail.HTMLBody = body
            mail.Send()
    @staticmethod
    class Body:
        @staticmethod
        def Test()->str:
            return \
            f"""
            <p>
                메일 본문 작성
            </p>
            """
        
import zipfile
import shutil

# 한글 인코딩 -> 압축 해제
def zip_extract_remove(contain_str: str, src: str, dst: str)->None:
    _zips = [_zip for _zip in os.listdir(src) if contain_str in _zip][-1]
    shutil.move(
        os.path.join(src, _zips),
        os.path.join(dst, _zips)
    )
    # encoding
    with zipfile.ZipFile(os.path.join(dst, _zips), "r") as _czip:
        for member in _czip.infolist():
            member.filename = member.filename.encode("cp437").decode("euc-kr")
            _czip.extract(member, dst)
    # remove
    try: os.remove(os.path.join(dst, _zips))
    except: pass

# Hashing
class Hasher:
    @staticmethod
    def extract(worksheet: WorkSheet)->dict:
        val, cell = win32comExcel.extract_contain_merged_cells(worksheet)
        return {"Val": val, "Cell": cell}
    @staticmethod
    class Finder:
        @staticmethod
        def OneColRow(HT: dict, sheet_name: str, filtering: str, filter_col: str, target_col: str)->str:
            filter_col = win32comExcel.Encoder.I2L(filter_col)-1
            target_col = win32comExcel.Encoder.I2L(target_col)-1
            t_val, t_cell = HT[sheet_name]["Val"], HT[sheet_name]["Cell"]
            start_end_range = t_val[t_val.iloc[:, filter_col]==filtering].index.tolist()
            idx_start = t_cell.iloc[start_end_range[0], target_col]
            idx_end = t_cell.iloc[start_end_range[-1], target_col]
            return f"{idx_start}:{idx_end}"
        @staticmethod
        def NcolRow(HT: dict, sheet_name: str, filtering: str, filter_col: str, target_col1: str, target_col2: str)->str:
            idx_start = Hasher.Finder.OneColRow(HT, sheet_name, filtering, filter_col, target_col1)
            idx_end = Hasher.Finder.OneColRow(HT, sheet_name, filtering, filter_col, target_col2)
            return f"{idx_start.split(':')[0]}:{idx_end.split(':')[-1]}"
        @staticmethod
        def byVal(HT: dict, sheet_name: str, filtering: str, start_pos: int=-1)->str:
            t_val, t_cell = HT[sheet_name]["Val"], HT[sheet_name]["Cell"]
            _temp_df = t_val.isin([filtering])
            _temp_in = _temp_df.any()
            col = _temp_in[_temp_in==True].index[start_pos]
            row = _temp_df[col][_temp_df[col]==True].index[start_pos]
            coor = t_cell.iloc[int(row), int(col)]
            return coor
        
import subprocess

class SAP:
    @staticmethod
    def logon(id: str, pw: str, window_size: list=[90, 20])->Session:
        subprocess.check_call([
            "SAP 경로, sapshcut.exe",
            f"-system=MEP",
            f"-client=100",
            f"-user={id}",
            f"-pw={pw}"
        ])
        time.sleep(15)
        # 정상 실행 확인
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(SapGuiAuto)==win32com.client.CDispatch: return
        application = SapGuiAuto.GetScriptingEngine
        if not type(application)==win32com.client.CDispatch:
            SapGuiAuto = None
            return
        connection = application.Children(0)
        if not type(connection)==win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return
        session = connection.Children(0)
        if not type(session)==win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return
        session.findById("wnd[0]").resizeWorkingPane(window_size[0], window_size[1], False)
        time.sleep(15)
        return session
    @staticmethod
    def logout()->None:
        try: os.system("taskkill /f /im saplogon.exe")
        except: pass
    @staticmethod
    def tcode(session: Session, tcode: str)->Session:
        session.findById("wnd[0]/tbar[0]/okcd").text = tcode
        session.findById("wnd[0]").sendVKey(0)
        return session
