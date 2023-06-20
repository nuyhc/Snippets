import sys
import torch
import pyautogui
from PIL import Image
import pandas as pd

class NotDetected(Exception):
    def __init__(self):
        super.__init__("더미 버튼 검출 실패")
        
virkp_temp = pd.DataFrame({
    "인덱스 별 실제 값"
})
special_char_dict = {"shift를 눌러야 나오는 특수 문자 변환 딕셔너리"}

# 가상 키패드 클릭
def click_on_virk(virkp_location: dict, x: float, y: float, mx: float, my: float)->None:
    """
    가상 키패드 클릭, Object Detection으로 추정된 좌표(xmin, ymin)을 입력하면 박스의 중앙을 클릭
    환경에 따라, virkp_location의 값이 일치하지 않을 수 있음. 고정된 해상도라면 휴리스틱하게 조절 가능
    Args:
        virkp_location (dict): selenium으로 추정한 키패드 위치
        x (float): 추정 xmin
        y (float): 추정 ymin
        mx (float): 추출된 박스들의 평균 w
        my (float): 추출된 박스들의 평균 h
    """
    pyautogui.click((
        virkp_location["x"] + x + mx/2,
        virkp_location["y"] + y + my/2
    ))
    
# 이미지 크롭
def crop_virkp(virkp_location: dict, virkp_size: dict)->Image:
    """
    셀레니움으로 제어하는 전체 창에서, 가상 키패드 부분만 크롭
    주소 입력창, 크롬 제어 문구가 표시되는 크기는 1920*1080 기준 115임 (url_y_margin)
    Args:
        virkp_location (dict): selenium으로 추정한 키패드 위치
        virkp_size (dict): selenium으로 추정한 키패드 크기

    Returns:
        Image: 크롭된 이미지
    """
    virkp_img = Image.open("selenium_captured_img_path")
    crop_x, crop_y = virkp_location["x"], virkp_location["y"]
    crop_w, crop_h = crop_x+virkp_size["width"], crop_y+virkp_size["height"]
    url_y_margin = 115
    virkp_img = virkp_img.crop((
        int(crop_x), int(crop_y)+url_y_margin, int(crop_w), int(crop_h)+url_y_margin
    ))
    virkp_img.save("selenium_captured_img_path")
    return virkp_img

# 더미 버튼 검출
def detect_dummpy_pos()->list:
    # yolov7x fine tune model architecture path (models/utils 폴더)
    sys.path.insert(0, "폴더 경로")
    # 가중치 로드
    dummy_detecter = torch.load("./hometax_for_cpu.pt", map_location="cpu")
    # (상위 2개 줄은, 구동 환경에 맞게 수정)
    detect_df = dummy_detecter("selenium_captured_img_path")
    detect_df = detect_df.pandas().xyxy[0]
    detect_df["xlen"] = detect_df["xmax"] - detect_df["xmin"]
    detect_df["ylen"] = detect_df["ymax"] - detect_df["ymin"]
    # 박스별 평균 w, h 사이즈
    mx, my = detect_df["xlen"].mean(), detect_df["ylen"].mean()
    detect_df = detect_df.sort_values(["ymin", "xmmin"]).reset_index(drop=True) # ymin으로 오름차순 정렬시, 1 행부터 순서대로 정렬됨
    return detect_df, mx, my

# 가상 키패드 좌표 추정
def assume_virkp(detect_df: pd.DataFrame, mx: float, my: float)->pd.DataFrame:
    if len(detect_df)!=0: raise NotDetected
    else:
        # 각 행별 더미 버튼 2개 찾기
        cri_x, cri_y = [], []
        for idx in range(0, 8, 2):
            dummy_1, dummy_2 = detect_df[["xmin", "ymin"]].iloc[idx:idx+2].values
            cri_y.append((dummy_1[1]+dummy_2[1])/2) # y값은 평균값 이용
            cri_x.append(sorted([dummy_1[0], dummy_2[0]])) # x값은 왼쪽 -> 오른쪽
        # 각 행별 버튼 위치 추정
        # 검출된 더미 버튼을 기준으로,
        # 박스의 평균 크기+박스간 거리(5)만큼 이동하면서 전체적이 좌표를 구하는 방식
        virkp_mapped = []
        for t_idx in range(0, len(cri_y)):
            x_cor = []
            # 첫번째 더미 버튼의 좌측 <-
            idx = 1
            while(cri_x[t_idx][0]-idx*(mx+5)>0):
                x_cor.append(cri_x[t_idx][0]-idx*(mx+5))
                idx += 1
            # 첫번째 더미 버튼 우측 ~ 두번째 더미 버튼 좌측
            idx = 1
            while(cri_x[t_idx[0]+idx*(mx+5)]<cri_x[t_idx][1]):
                x_cor.append(cri_x[t_idx[0]+idx*(mx+5)])
                idx += 1
            # 두번째 더미 버튼의 우측
            idx = 1
            while(cri_x[t_idx[1]+idx*(mx+5)]<14*(mx+5)):
                x_cor.append(cri_x[t_idx[1]+idx*(mx+5)])
                idx += 1
            #
            virkp_mapped.append(pd.DataFrame({
                "assume_x":sorted(x_cor),
                "assume_y":[cri_y[t_idx] for _ in range(len(x_cor))]
            }))
            # 마지막행: 4행에서 y축만 이동
            if t_idx==3:
                virkp_mapped.append(pd.DataFrame({
                    "assume_x":[sorted(x_cor)[0]+idx*(mx+5) for idx in range(14)],
                    "assume_y":[cri_y[t_idx]+(my+5) for _ in range(14)]
                }))
        virkp_mapped = pd.concat(virkp_mapped).reset_index(drop=True)
        virkp_mapped = virkp_mapped.reset_index(drop=False) # 맵핑을 위해 초기화
        virkp_mapped = pd.merge(
            left=virkp_temp, right=virkp_mapped,
            left_on=["index"], right_on=["index"], how="left"
        )[["value", "assume_x", "assume_y"]]
        return virkp_mapped
    
# 가상 키패드 비밀번호 입력
def login_virkp(virkp_location: dict, virkp_size: dict, pw: str)->None:
    virkp_img = crop_virkp(virkp_location, virkp_size) # 이미지 크롭
    try:
        detect_df, mx, my = detect_dummpy_pos() # 더미 버튼 위치 추정
        virkp_mapped = assume_virkp(detect_df, mx, my) # 전체 키패드 추정
    except NotDetected:
        # 8개의 더미 버튼 추정 실패 시,
        pass
    # 비밀 번호 입력
    _, shift_x, shift_y = virkp_mapped[virkp_mapped["value"]=="Shift"].values[0]
    for pw_c in pw:
        if pw_c.isalpha():
            if pw_c.islower(): # 소문자
                _, corx, cory = virkp_mapped[virkp_mapped["value"]==pw_c].values[0]
                click_on_virk(virkp_location, corx, cory, mx, my)
            else: # 대문자
                _, corx, cory = virkp_mapped[virkp_mapped["value"]==pw_c.lower()].values[0]
                click_on_virk(virkp_location, shift_x, shift_y, mx, my)
                click_on_virk(virkp_location, corx, cory, mx, my)
                click_on_virk(virkp_location, shift_x, shift_y, mx, my)
        elif pw_c.isnumeric(): # 숫자
            _, corx, cory = virkp_mapped[virkp_mapped["value"]==pw_c].values[0]
            click_on_virk(virkp_location, corx, cory, mx, my)
        else: # 특수 문자
            if pw_c in ["-", "=", "\\", "[" "]", ";", "'", ",", ".", "/"]: # 기본 화면
                _, corx, cory = virkp_mapped[virkp_mapped["value"]==pw_c].values[0]
                click_on_virk(virkp_location, corx, cory, mx, my)
            else: # shift 누른 화면
                _, corx, cory = virkp_mapped[virkp_mapped["value"]==pw_c].values[0]
                click_on_virk(virkp_location, shift_x, shift_y, mx, my)
                click_on_virk(virkp_location, corx, cory, mx, my)
                click_on_virk(virkp_location, shift_x, shift_y, mx, my)
    # Enter
    _, enter_x, enter_y = virkp_mapped[virkp_mapped["value"]=="Enter"].values[0]
    click_on_virk(virkp_location, enter_x, enter_y, mx, my)
