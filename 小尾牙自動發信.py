# coding:utf-8
counterr = open('計數器.txt', 'r', encoding = 'UTF-8')
sequence = 0
counter = counterr.readlines()
counterr.close()

for i in counter:
    sequence = int(i)

import pandas as pd
data = pd.read_excel("來賓資料.xlsx")
NameList = data['姓名'].tolist()
AddressList = data['聯絡 E-mail'].tolist()
GenderList = data['性別'].tolist()
GenderTitle = {'男':'先生', '女':'女士'}
from PIL import Image, ImageDraw, ImageFont

tfont = ImageFont.truetype('testfont.TTC', size = 150)
tfont2 = ImageFont.truetype('testfont.TTC', size = 50)
(x1,y1) = (307, 1110)
(x2,y2) = (800, 1260)
color = 'rgb(0, 0, 0)'

import os
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
import smtplib
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.login('2021fincancebanquet@gmail.com','xxxx')
for i in range(len(NameList)):
    sequence += 1
    msg = MIMEMultipart()
    msg["From"] = '2021fincancebanquet@gmail.com'
    msg["To"] = AddressList[i]
    msg["Subject"] = '《2021財金小尾牙》誠摯邀請您蒞臨參與'
    temp = NameList[i]
    if len(temp) == 2:
        temp = temp[0]+' '+temp[1]
    content = '致' + NameList[i] + GenderTitle[GenderList[i]] + ':\n  您已成功報名2021財金小尾牙，感謝您的參與，期待在12/28與您一同共襄盛舉!\n  附件圖檔爲電子票，於活動當天入場時段向門口工作人員出示即可入場。\n  在活動尾聲有抽獎活動，請留意電子票上的「抽獎序號」。另外，fb粉專（「2021財金小尾牙」）也有四次抽獎活動，截止日皆為2021/12/27 12:00，請不要吝嗇的動手指抽抽看吧～\n\n\n\n《活動資訊》\n活動時間: 2021/12/28 18:00 - 18:25入場 18:30表演正式開始\n地點: 台灣大學第一學生活動中心1樓怡仁堂'
    text = MIMEText(content)
    msg.attach(text)

    image = Image.open('invitation.jpg')
    draw = ImageDraw.Draw(image)
    draw.text((x1, y1), temp, fill=color, font=tfont)
    draw.text((x2, y2), str(sequence), fill=color, font=tfont2)
    image.save('invitation/' + str(sequense) + NameList[i] + ".jpg")

    with open('invitation/' + str(sequense) + NameList[i] + ".jpg", 'rb') as f:
        img_data = f.read()

    invitation = MIMEImage(img_data, name=os.path.basename('invitation/' + str(sequense) + NameList[i] + ".jpg"))
    msg.attach(invitation)
    server.send_message(msg)
    print('x', NameList[i])

server.close()
counterw = open('計數器.txt', 'w')
counterw.write(str(sequence))
counterw.close()






