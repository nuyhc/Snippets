import os
import sys

# 지정된 경로에서, 가장 최근 파일을 찾는 함수
def Get_Recent_File(PATH: str)->str:
    path_and_time = []
    for each_file_name in os.listdir(PATH):
        # 편집중인 문서 제거
        if each_file_name[:2]=="~$": each_file_name = each_file_name[2:]
        each_name = each_file_name
        each_path = f"{PATH}/{each_file_name}"
        each_time = os.path.getctime(each_path)
        path_and_time.append((each_name, each_time))
    recent = max(path_and_time, key=lambda x: x[1])[0]
    return recent

# 파일 이름 변경
def Change_File_Name(PATH: str, Prev: str, New: str)->None:
    file_path = f"{PATH}/{Prev}"
    os.rename(file_path, PATH+"/"+str(New))
    
# 해당 파일의 절대 경로를 찾아주는 함수
# pyinstaller를 이용해, .exe 형태로 만들때 주로 활용
def Resource_Path(PATH: str)->str:
    try: BASE_PATH = sys._MEIPASS
    except Exception: BASE_PATH = os.path.abspath(".")
    return os.path.join(BASE_PATH, PATH)
