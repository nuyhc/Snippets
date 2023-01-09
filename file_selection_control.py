import time
from pywinauto.keyboard import send_keys

# 웹 제어(ex, 셀레니움) / 파일 업로드 창 발생시 파일 업로드
def file_select_control(Upload_File_Path: str)->None:
    time.sleep(1)
    send_keys(Upload_File_Path, with_spaces=True)
    send_keys("{ENTER}")
    time.sleep(3)
