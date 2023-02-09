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

##########################################################################################

# 메일 첨부 항목 다운
# 아래 mail 객체를 이용해 수신 항목에 대한 다양한 정보를 받을 수 있음
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inboxfolder = outlook.GetDefaultFolder(6).Folders("폴더 이름") # 6은 받은 편지함을 의미 (전체)
messages = inboxfolder.Items

for mail in messages:
    # 보낸 사람 메일 주소
    if mail.SenderEmailType=="EX": mail.Sender.GetExchangeUser().PrimarySmtpAddress
    else: mail.SenderEmailAddress
    # 첨부 파일 저장
    attachments = mail.Attachments
    rnct = attachments.Count
    for cnt in range(1, rcnt+1):
        attachment = attachments.Item(cnt)
        attachment.SaveASFile("경로")
