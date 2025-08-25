import imaplib
import email
from email.header import decode_header
import re
import pandas as pd

# Gmail IMAP server settings
IMAP_SERVER = "imap.gmail.com"
IMAP_PORT = 993

# 登录 Gmail 帐号，使用 App Password 代替 Gmail 密码
USERNAME = "lyy010714@gmail.com"
PASSWORD = "xadxbsjbuqghswvr"


def clean(text):
    """清理邮件中的非法字符"""
    return "".join(c if c.isalnum() else "_" for c in text)

def get_emails():
    # 连接到 Gmail 邮箱
    print("正在连接到 Gmail 邮箱...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    try:
        mail.login(USERNAME, PASSWORD)
        print(f"登录成功: {USERNAME}")
    except Exception as e:
        print(f"登录失败: {e}")
        return []

    # 选择邮箱中的“收件箱”
    mail.select("inbox")
    print("已选择收件箱...")

    # 搜索所有发件人为 no-reply@arxiv.org 的邮件
    print("正在搜索邮件...")
    status, messages = mail.search(None, 'FROM "no-reply@arxiv.org"')

    if status != "OK":
        print("没有找到相关邮件。")
        return []

    # 获取所有邮件的 ID
    email_ids = messages[0].split()
    print(f"找到 {len(email_ids)} 封相关邮件...")

    email_data = []

    for email_id in email_ids:
        # 获取邮件
        print(f"正在检索第 {email_id} 封相关邮件...")
        status, msg_data = mail.fetch(email_id, "(RFC822)")

        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])

                # 解码邮件的主题
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding or "utf-8")

                # 提取发件人信息
                from_ = msg.get("From")

                # 获取邮件的正文内容
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            body = part.get_payload(decode=True).decode()
                            break
                else:
                    body = msg.get_payload(decode=True).decode()

                # 正则提取所需信息
                articles = re.findall(r'arXiv:(\d+\.\d+)', body)
                titles = re.findall(r'Title:\s*(.*)', body)
                authors = re.findall(r'Authors:\s*(.*)', body)
                link = re.findall(r'https://arxiv.org/abs/(\d+\.\d+)', body)

                for i in range(len(articles)):
                    article_data = {
                        "title": titles[i] if i < len(titles) else "No Title",
                        "authors": authors[i] if i < len(authors) else "No Authors",
                        "link": f"https://arxiv.org/abs/{articles[i]}",
                    }
                    email_data.append(article_data)

    return email_data

def save_as_excel(email_data):
    # 创建 Excel 文件并保存数据
    print("正在保存 Excel 文件...")
    df = pd.DataFrame(email_data)
    df.to_excel("arxiv_optics_articles.xlsx", index=False, engine='openpyxl')
    print("Excel 文件已保存：arxiv_optics_articles.xlsx")

def main():
    print("开始处理...")
    email_data = get_emails()
    
    if email_data:
        save_as_excel(email_data)
    else:
        print("没有获取到相关邮件，无法生成 Excel 文件。")

if __name__ == "__main__":
    main()