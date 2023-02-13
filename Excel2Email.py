import json as json
import pandas as pd
import smtplib
import re
import numpy as np
# import time
# import os
from email.mime.text import MIMEText
from email.header import Header


def load_config():
    configPath = "config.json"

    with open(configPath, "r", encoding="utf8") as f:
        data = json.load(f)
        return data


def load_excel(excelFile):
    pd.set_option("display.notebook_repr_html", False)
    data = pd.read_excel(io=excelFile, sheet_name=0)
    data.fillna("", inplace=True)
    return data


def load_html(htmlFile):
    with open(htmlFile) as f:
        content = f.read().replace("\n", " ")
    return content


def html_regex(match, data) -> str:
    matchGroup1 = match.group(1)
    if matchGroup1 is None:
        return ""

    replaceData = data.get(matchGroup1)
    if replaceData is None:
        return ""

    if not isinstance(replaceData, str):
        return str("" if np.isnan(replaceData) else replaceData)
    return str(replaceData)


def html_advance_replace(htmlTemplate, data):
    return re.sub(r"\${([^}]*)}", lambda match: html_regex(match, data), htmlTemplate)


def email_connect_to_server(mailConfig):
    try:
        smtp = smtplib.SMTP()
        smtp.connect(mailConfig.get("smtp"))
        smtp.login(mailConfig.get("email"), mailConfig.get("password"))
    except:
        print("邮件服务器连接失败")
    else:
        print("邮件服务器连接成功")
        return smtp


def email_disconnect_from_server(smtp):
    smtp.quit()
    print("邮件服务器断开连接")


def email_send(smtp, subject, fromEmail, toEmail, mail_body):
    if toEmail == None or toEmail == "":
        print("邮箱为空，跳过")
        return
    # t = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    # 组装邮件内容和标题，中文需参数utf-8，单字节字符不需要

    # print( "发送邮件" + toEmail)
    # print(mail_body)
    msg = MIMEText(mail_body, _subtype="html", _charset="utf-8")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = fromEmail
    msg["To"] = toEmail
    try:
        smtp.sendmail(fromEmail, toEmail, msg.as_string())
    except:
        print("邮件发送失败！")
    else:
        print("邮件发送成功！")


def run():
    config = load_config()
    # Load Excel data
    excelPath = config.get("excel")
    excelData = load_excel(excelPath)

    # Load HTML
    htmlPath = config.get("html")
    htmlTemplate = load_html(htmlPath)

    # Loop by Excel Row
    mailConfig = config.get("mail")
    fromEmail = mailConfig.get("email")
    Subject = mailConfig.get("subject")
    toMailKey = mailConfig.get("toEmail")
    logKey = mailConfig.get("log")

    smtp = email_connect_to_server(mailConfig)
    for idx, row in excelData.iterrows():
        print("开始发送邮件:" + str(idx) + " " + str(row.get(logKey)))
        # Replace HTML
        htmlBody = html_advance_replace(htmlTemplate, row)

        # Send Email
        # email_send(smtp, Subject, fromEmail, "yingsong@encompass8.cn", htmlBody)
        # # break
        email_send(smtp, Subject, fromEmail, row.get(toMailKey), htmlBody)


    email_disconnect_from_server(smtp)


run()
