import os
import win32com.client

# .txt 파일에서 수신자/참조자 정보 가져오기
# 여러명일 경우, ";"를 이용해 구분되어 있음
def Get_TO_CC(FileName: str)->str:
    cwd = os.getcwd()
    with open(cwd+f"/{FileName}.txt", "r") as f:
        lines = f.readlines()
    email_list = [line.strip() + ";" for line in lines if "@" in line]
    result = " ".join(email_list)
    return result

# 메일 본문 작성 (html)
def make_mail()->str:
    # style과 body 커스텀 가능
    style = \
        """
        <style>
            p{
                font-size: 13px;
                font-family: "Malgun Gothic";
            }
        </style>
        """
    body = \
        f"""
        <p>
            메일 본문 (html)
        </p>
        """
    return style+body

TO = "수신.txt"
CC = "참조.txt"

# 메일 발송
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

mail.To = Get_TO_CC(TO)
mail.CC = Get_TO_CC(CC)

mail.Subject = f"메일 제목"
mail.HTMLBody = make_mail()

mail.Send()
