import time
import email
from email.header import decode_header
import yaml
from exchangelib import Credentials, Account, DELEGATE, EWSTimeZone
from exchangelib.queryset import Q
from tqdm import tqdm
import sqlite3
import threading
import os
import imaplib
from bs4 import BeautifulSoup
import threading
from urllib.parse import urlparse, parse_qs
import urllib.parse
import requests
import platform
from flask import Flask, render_template, send_file, request

proxy_url = "http://127.0.0.1:8080"
proxy_is_running = False


def decode_header_text(header_value, header_charset):
    if isinstance(header_value, bytes):
        text = header_value.decode(header_charset)
    else:
        text = header_value
    return text


def check_proxy():
    global proxy_is_running
    while True:
        try:
            requests.get(proxy_url, timeout=1)
            os.environ["HTTP_PROXY"] = proxy_url
            os.environ["HTTPS_PROXY"] = proxy_url
            proxy_is_running = True
        except requests.exceptions.RequestException:
            os.environ["HTTP_PROXY"] = ""
            os.environ["HTTPS_PROXY"] = ""
            proxy_is_running = False

        time.sleep(60)


def get_windows_version():
    if platform.system() != "Windows":
        return False
    version = platform.win32_ver()[0]
    if version == "10":
        return True
    elif version == "11":
        return True
    else:
        return False


def load_config(config_file):
    if os.path.exists(config_file):
        with open(config_file, "r", encoding="utf-8") as f:
            return yaml.safe_load(f)
    else:
        print(f"未找到{config_file}文件，必须提供邮箱账户和密码！\n")
        username = input("请输入邮箱账户：")
        import getpass

        password = getpass.getpass("请输入邮箱密码：")
        return {"username": username, "password": password}


config = load_config("config.yaml")
username = config.get("username") or ""
password = config.get("password") or ""
ESkeys = config.get("ESkeys") or ""
interval = config.get("interval") or 120
host = config.get("host") or "127.0.0.1"
port = config.get("port") or 5000
imap_server = config.get("imap_server") or "imap.qq.com"
address = config.get("address") or "scholaralerts-noreply@google.com"
related_tag = config.get("related_tag") or " - 新的相关研究工作"
new_articles_tag = config.get("new_articles_tag") or " - 新文章"
new_citations_tag = config.get("new_citations_tag") or "的文章新增了"
new_results_tag = config.get("new_results_tag") or " - 新的结果"
new_citations_num = config.get("new_citations_num") or 0


if (password == "") | (username == ""):
    username = input("请输入邮箱账户：")
    password = input("请输入邮箱密码：")


class senders:
    email_address = ""
    name = ""


def create_database():

    if os.path.exists("APP.db"):
        return None
    else:
        conn = sqlite3.connect("APP.db")
        c = conn.cursor()

        c.execute(
            """CREATE TABLE inbox
                       (account text, email_num text, time text, email_type text)"""
        )
        conn.commit()
        conn.close()


def add_database(account, email_num, time, email_type):

    conn = sqlite3.connect("APP.db")
    c = conn.cursor()
    c.execute("SELECT * FROM inbox WHERE account=?", (account,))
    rows = c.fetchall()
    if len(rows) == 0:
        c.execute(
            "INSERT INTO inbox VALUES (?,?,?,?)", (account, email_num, time, email_type)
        )
        conn.commit()
        conn.close()
        return True
    else:
        conn.close()
        return False


def get_database():
    conn = sqlite3.connect("APP.db")
    c = conn.cursor()
    c.execute("SELECT * FROM inbox")
    rows = c.fetchall()
    conn.close()
    return rows


def change_email_num(account, email_num):
    conn = sqlite3.connect("APP.db")
    c = conn.cursor()
    c.execute("UPDATE inbox SET email_num=? WHERE account=?", (email_num, account))
    conn.commit()
    conn.close()
    return True


def get_email_num(account):
    conn = sqlite3.connect("APP.db")
    c = conn.cursor()
    c.execute("SELECT * FROM inbox WHERE account=?", (account,))
    rows = c.fetchall()
    conn.close()
    if len(rows) == 0:
        return 9999999999999999999
    else:
        return rows[0][1]


def find_database(account):
    conn = sqlite3.connect("APP.db")
    c = conn.cursor()
    c.execute("SELECT * FROM inbox WHERE account=?", (account,))
    rows = c.fetchall()
    conn.close()
    if len(rows) == 0:
        return False
    else:
        return True


def get_emailtype(account):
    conn = sqlite3.connect("APP.db")
    c = conn.cursor()
    c.execute("SELECT * FROM inbox WHERE account=?", (account,))
    rows = c.fetchall()
    conn.close()
    return rows[0][3]


def store_emailtype(account, email_type):
    conn = sqlite3.connect("APP.db")
    c = conn.cursor()
    c.execute("UPDATE inbox SET email_type=? WHERE account=?", (email_type, account))
    conn.commit()
    conn.close()
    return True


class EmailAccount:
    def __init__(self, username, password):
        print(f" * 当前邮箱为 {username}")
        self.username = username
        self.password = password
        self.databasename = username + ".db"
        self.login_success = False
        self.account = None
        self.accountexist = find_database(username)
        self.databaseexist = self.databasecheck()
        self.email_type = "outlook"
        self.init_success = self.databaseexist & self.accountexist
        if self.accountexist & (not self.databaseexist):
            print(" * 数据库遗失，将删除账户重新初始化")
            conn = sqlite3.connect("APP.db")
            c = conn.cursor()
            c.execute("DELETE FROM inbox WHERE account=?", (username,))
            conn.commit()
            conn.close()
            self.accountexist = False

        if self.init_success:
            self.email_type = get_emailtype(username)

        else:
            print(" * 开始初始化账户，请稍后...")
            state = self.email_type_check()
            if state:
                self.create_database_paper()
                self.init_account()

                self.init_success = True
            else:
                pass

    # 自动校正邮箱类型
    def email_type_check(self):
        try:
            credentials = Credentials(username=self.username, password=self.password)
            account = Account(
                primary_smtp_address=self.username,
                credentials=credentials,
                autodiscover=True,
                access_type=DELEGATE,
            )
            self.email_type = "outlook"
            self.login_success = True
            self.account = account

            return True
        except Exception as e1:
            # 尝试使用qq登录
            try:
                imap = imaplib.IMAP4_SSL(imap_server)
                imap.login(self.username, self.password)
                self.email_type = "qq"
                self.login_success = True
                self.account = imap
                return True
            except Exception as e2:
                print(f"\n  Exchange登录异常: {e1}")
                print(f"  Imap    登录异常: {e2}\n")
                self.login_success = False
                return False

    def databasecheck(self):
        # 判断数据库是否存在
        if os.path.exists(self.databasename):
            return True
        else:
            return False

    # 定义创建数据库函数
    def log_in(self):
        if self.email_type == "outlook":
            try:
                credentials = Credentials(
                    username=self.username, password=self.password
                )
                account = Account(
                    primary_smtp_address=self.username,
                    credentials=credentials,
                    autodiscover=True,
                    access_type=DELEGATE,
                )
                self.login_success = True
                self.account = account
                return account
            except Exception as e:

                print("登录异常:", e)
                self.login_success = False
                return None
        elif self.email_type == "qq":
            # 使用imaplib登录
            try:

                imap = imaplib.IMAP4_SSL(imap_server)
                imap.login(self.username, self.password)
                self.login_success = True
                self.account = imap
                return imap
            except Exception as e:

                print("登录异常:", e)
                self.login_success = False
                return None

    # 初始化，获取现有邮件录入数据库
    def init_account(self):

        if ESkeys == "":
            print(" * 未设置Easy Scholar key，不会获取影响因子")
        print("\n * 开始录入已有邮件...")
        inbox = self.get_email_content_by(email_address=address)
        if inbox != None:
            num = len(inbox)
            timenow = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            # 将inbox[0]中的body按照\r\n分割
            articles = []
            for i in range(len(inbox)):
                bodyhtml = inbox[i]["body"]
                tag = inbox[i]["subject"]
                rtime = inbox[i]["time"]
                tag, tag_type = clear_tag(tag)
                articlesnew = get_articles(bodyhtml, tag, tag_type, rtime)
                articles = articles + articlesnew
            print("\n * 开始录入已有文献...")
            print(" * 共计", len(articles), "篇")
            # 存入数据库
            for article in tqdm(articles, desc=" * 正在录入文献 "):
                self.add_database_paper(
                    article["title"],
                    article["author"],
                    article["journal"],
                    article["date"],
                    article["link"],
                    article["IF"],
                    article["IF5"],
                    article["sciUp"],
                    article["tags"],
                    article["tag_type"],
                )

            add_database(self.username, num, timenow, self.email_type)
        else:
            num = 0
            timenow = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            add_database(self.username, num, timenow, self.email_type)
        print("\n * 账户初始化完成！\n")
        self.init_success = True

    def create_database_inbox(self):
        # 判断数据库是否存在
        if os.path.exists(self.databasename):
            return None
        else:
            conn = sqlite3.connect(self.databasename)
            c = conn.cursor()
            # 创建一个表
            c.execute(
                """CREATE TABLE inbox
                       (subject text, email_address text, name text, body text)"""
            )
            conn.commit()
            conn.close()

    # 获取数据库中的数据
    def get_database_inbox(self):
        conn = sqlite3.connect(self.databasename)
        c = conn.cursor()
        c.execute("SELECT * FROM inbox")
        rows = c.fetchall()
        conn.close()
        return rows

    # 添加数据到数据库
    def add_database_inbox(self, subject, email_address, name, body):
        # 判断数据是否存在
        conn = sqlite3.connect(self.databasename)
        c = conn.cursor()
        c.execute(
            "SELECT * FROM inbox WHERE subject=? AND email_address=? AND name=? AND body=?",
            (subject, email_address, name, body),
        )
        rows = c.fetchall()
        if len(rows) == 0:
            c.execute(
                "INSERT INTO inbox VALUES (?,?,?,?)",
                (subject, email_address, name, body),
            )
            conn.commit()
            conn.close()
            return True
        else:
            conn.close()
            return False

    # 创建论文数据库，包含论文标题，作者，期刊，发表时间，下载链接，影响因子
    def create_database_paper(self):
        # 判断数据库是否存在
        if os.path.exists(self.databasename):
            return None
        else:
            conn = sqlite3.connect(self.databasename)
            c = conn.cursor()
            # 创建一个表
            c.execute(
                """CREATE TABLE paper
                       (title text, author text, journal text, date text, link text, IF text, IF5 text, sciUp text, new_article text, new_citation text, related text, subject text)"""
            )
            conn.commit()
            conn.close()

    # 获取数据库中的数据
    def get_database_paper(self):
        conn = sqlite3.connect(self.databasename)
        c = conn.cursor()
        c.execute("SELECT * FROM paper")
        rows = c.fetchall()
        conn.close()
        return rows

    # 添加数据到数据库
    def add_database_paper(
        self, title, author, journal, date, link, IF, IF5, sciUp, tag, tag_type
    ):
        # 判断数据是否存在
        conn = sqlite3.connect(self.databasename)
        c = conn.cursor()
        c.execute("SELECT * FROM paper WHERE link=?", (link,))
        rows = c.fetchall()
        new_article = ""
        new_citation = ""
        related = ""
        subject = ""
        if tag_type == "new-articles":
            new_article = tag
        elif tag_type == "new-citations":
            new_citation = tag

        elif tag_type == "related":
            related = tag
        elif tag_type == "subject":
            subject = tag

        if len(rows) == 0:
            c.execute(
                "INSERT INTO paper VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                (
                    title,
                    author,
                    journal,
                    date,
                    link,
                    IF,
                    IF5,
                    sciUp,
                    new_article,
                    new_citation,
                    related,
                    subject,
                ),
            )
            conn.commit()
            conn.close()
            # print("已添加:",title[0:100])
            return True
        else:

            if len(rows[0]) == 12:
                new_articleold = rows[0][8]
                new_citationold = rows[0][9]
                relatedold = rows[0][10]
                subjectold = rows[0][11]
            else:
                new_articleold = ""
                new_citationold = ""
                relatedold = ""
                subjectold = ""

            if tag_type == "new-articles":
                tagsold = new_articleold
                tagname = "new_article"
                tag = new_article
            elif tag_type == "new-citations":
                tagsold = new_citationold
                tagname = "new_citation"
                tag = new_citation
            elif tag_type == "related":
                tagsold = relatedold
                tagname = "related"
                tag = related
            elif tag_type == "subject":
                tagsold = subjectold
                tagname = "subject"
                tag = subject

            else:
                tagsold = ""
                tagname = ""
                tag = ""
            tagsold = tagsold.split(";")
            # 去除空值
            if tagsold[0] == "":
                tagsold = []
            # 判断tag是否存在
            if tag in tagsold:
                # print("已存在：",title[0:100])
                conn.close()
                return False
            else:
                # 将tag添加到tagsold中
                tagsold.append(tag)
                # 将tagsold转换为字符串
                tagsnew = ";".join(tagsold)
                # 更新tags
                if tagname != "":
                    c.execute(
                        f"UPDATE paper SET {tagname}=? WHERE link=?", (tagsnew, link)
                    )
                    # 更新时间
                    c.execute("UPDATE paper SET date=? WHERE link=?", (date, link))
                    # print("已更新",title[0:100])
                else:
                    pass
                conn.commit()
                conn.close()
                return False

    # 定义检查邮件函数

    def check_for_email_numold(self, name):
        try:
            account = self.log_in()
            if account is None:
                return None
            # 获取当前收件箱中的所有邮件
            all_items = account.inbox.all().order_by("-datetime_received")
            # 过滤出来自name的邮件
            items_from_name = [item for item in all_items if item.author.name == name]
            # 获取邮件总数
            current_item_count = len(items_from_name)
            return current_item_count
        except Exception as e:
            print("检查邮件数量出错：", e)
            return 0

    def check_for_email_num(self):
        try:
            account = self.log_in()
            if account is None:
                return 0

            if self.email_type == "outlook":
                q = Q(sender__icontains=address)
                # 使用Q对象来过滤邮件
                items_from_name = account.inbox.filter(q)
                # 获取邮件总数
                current_item_count = items_from_name.count()
            elif self.email_type == "qq":
                # 使用imaplib获取邮件
                account.select("inbox")
                # 获取emailaddress为address的邮件
                typ, data = account.search(None, f'(FROM "{address}")')
                current_item_count = len(data[0].split())

            return current_item_count
        except Exception as e:
            print("检查邮件数量出错：", e)
            return 0

    def check_for_new_email(self):
        # 获取数据库中的num
        num = int(get_email_num(self.username))
        # 获取当前收件箱中邮件总数
        current_item_count = self.check_for_email_num()
        # 判断是否有新邮件
        if current_item_count > num:
            return current_item_count - num, current_item_count
        else:
            return False, current_item_count

    def check_for_new_email_stream(self):
        # 设置Outlook账户的凭证信息
        timenow = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        print(f" * 开始检查新邮件 {timenow}")

        numnew, num = self.check_for_new_email()
        if numnew:
            print(f"您有{numnew}封新邮件")
            # 判断操作系统是否为win10
            if get_windows_version():
                from win10toast import ToastNotifier

                toaster = ToastNotifier()
                toaster.show_toast("新邮件提醒", f"您有{numnew}封新邮件", duration=10)

            inbox = self.get_email_content_by(
                email_address=address, start=0, end=-numnew
            )

            if inbox != None:

                articles = []
                for i in range(len(inbox)):
                    bodyhtml = inbox[i]["body"]

                    rtime = inbox[i]["time"]
                    tag = inbox[i]["subject"]
                    tag, tag_type = clear_tag(tag)
                    articlesnew = get_articles(bodyhtml, tag, tag_type, rtime)
                    articles = articles + articlesnew
                # 存入数据库
                print(" * 新文章共计", len(articles), "篇")
                for article in tqdm(articles, desc=" * 正在录入文章 "):

                    self.add_database_paper(
                        article["title"],
                        article["author"],
                        article["journal"],
                        article["date"],
                        article["link"],
                        article["IF"],
                        article["IF5"],
                        article["sciUp"],
                        article["tags"],
                        article["tag_type"],
                    )

                change_email_num(self.username, num)
            else:
                pass

    def get_email_content_by(self, email_address, name=None, start=0, end=0):

        if name == None and email_address == None:
            return None

        account = self.account
        if account is None:
            return []

        # 创建一个Q对象，用于过滤发件人的名称
        q = Q(sender__icontains=address)
        # 使用Q对象来过滤邮件
        inbox = []
        if end == 0:
            if self.email_type == "outlook":
                items = account.inbox.filter(q).order_by("-datetime_received")
                tn = items.count()
                tn = tn - start
                items = items[start:]

            elif self.email_type == "qq":
                # 使用imaplib获取邮件
                account.select("inbox")
                # 获取emailaddress为address的邮件
                typ, data = account.search(None, f'(FROM "{address}")')
                items = data[0].split()
                items.reverse()
                items = items[start:]
                tn = len(items)
            print(f" * 共计{tn}封邮件")

        else:
            if self.email_type == "outlook":
                items = account.inbox.filter(q).order_by("-datetime_received")[
                    start:-end
                ]

            elif self.email_type == "qq":
                # 使用imaplib获取邮件
                account.select("inbox")
                # 获取emailaddress为address的邮件

                typ, data = account.search(None, f'(FROM "{address}")')
                items = data[0].split()
                items.reverse()
                items = items[start:-end]

        for item in tqdm(items, desc=" * 正在录入邮件 "):
            msg = None
            # 获取邮件主题
            if self.email_type == "outlook":
                rtime = item.datetime_received.astimezone(
                    EWSTimeZone.localzone()
                ).strftime("%Y-%m-%d %H:%M:%S")
                subject = item.subject
                # 获取邮件发送者
                sender = item.sender
                # 获取邮件正文
                body = item.body
            elif self.email_type == "qq":
                # 创建sender类

                sender = senders()

                # 获取emailaddress为address的邮件
                status, data = account.fetch(item, "(RFC822)")

                # 解码
                msg = email.message_from_bytes(data[0][1])
                # 获取头文件的编码

                # 获取邮件主题
                sender.email_address = email.utils.parseaddr(msg.get("From"))[1]
                header_value, header_charset = decode_header(msg["Subject"])[0]
                subject = decode_header_text(header_value, header_charset)

                # 获取邮件发送者
                sender.email_address = email.utils.parseaddr(msg.get("From"))[1]
                # 获取发送者名称
                sender.name = email.utils.parseaddr(msg.get("From"))[0]
                # 解码
                header_value, header_charset = decode_header(sender.name)[0]
                sender.name = decode_header_text(header_value, header_charset)
                if header_charset == None:
                    header_charset = "utf-8"
                # 获取邮件正文
                if msg.is_multipart():
                    body = ""
                else:
                    body = msg.get_payload(decode=True).decode(header_charset)

                # 获取邮件收取时间
                rtime = msg.get("Date")

                rtime = rtime.split(", ")[1]

                rtime = rtime.split(" -")[0]
                rtime = rtime.split(" +")[0]
                # 转换为时间戳
                rtime = time.mktime(time.strptime(rtime, "%d %b %Y %H:%M:%S"))
                # 转换为当地时间
                rtime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(rtime))

            inbox.append(
                {
                    "subject": subject,
                    "email_address": sender.email_address,
                    "name": sender.name,
                    "body": body,
                    "time": rtime,
                }
            )

        return inbox

    def updatajournal(self):
        # 检查ESkeys是否为空
        if ESkeys == "":
            print(" * 未设置Easy Scholar key，不会获取影响因子")
            raise Exception("未设置Easy Scholar key，不会获取影响因子")
        conn = sqlite3.connect(self.databasename)

        data = self.get_database_paper()
        for i in range(len(data)):
            url = data[i][4]
            journal = data[i][2]
            IF = [data[i][5], data[i][6], data[i][7]]
            if journal.replace(" ", "") != "":
                IF, journal = journal_to_IF(journal)
                # 更新数据库
                c = conn.cursor()
                c.execute("UPDATE paper SET IF=? WHERE link=?", (IF[0], url))
                c.execute("UPDATE paper SET IF5=? WHERE link=?", (IF[1], url))
                c.execute("UPDATE paper SET sciUp=? WHERE link=?", (IF[2], url))
                c.execute("UPDATE paper SET journal=? WHERE link=?", (journal, url))
                conn.commit()
        conn.close()


def get_IF(title):
    IF = 0
    IF5 = 0
    sciUp = ""
    if ESkeys == "":
        return (IF, IF5, sciUp)

    url = f"https://www.easyscholar.cc/open/getPublicationRank?secretKey={ESkeys}&publicationName="

    # 拼接url，url编码title
    title = urllib.parse.quote(title)
    url = url + title
    # get请求
    try:
        r = requests.get(url)
    except Exception as e:
        print(f"\n期刊{urllib.parse.unquote(title)}获取影响因子失败，错误信息：{e}")
        return (IF, IF5, sciUp)

    # 检查状态码
    if r.status_code != 200:
        print(
            f"\n期刊{urllib.parse.unquote(title)}获取影响因子失败，状态码：{r.status_code}"
        )
        return (IF, IF5, sciUp)

    # 获取返回的json数据
    data = r.json()
    if data["code"] != 200:
        print(
            f"\n期刊{urllib.parse.unquote(title)}获取影响因子失败，错误信息：{data['msg']}"
        )
        return (IF, IF5, sciUp)
    time.sleep(0.5)

    # 判断data["data"]["officialRank"]["all"]是否为空
    if data["data"]["officialRank"]["all"] == None:
        # 解码
        print(f"\n期刊{urllib.parse.unquote(title)}未找到影响因子")
        return (IF, IF5, sciUp)

    # 查看data["data"]["officialRank"]["all"]下是否有sciif字段
    if "sciif" in data["data"]["officialRank"]["all"]:
        IF = float(data["data"]["officialRank"]["all"]["sciif"])
    # 查看data["data"]["officialRank"]["all"]下是否有sciif5字段
    if "sciif5" in data["data"]["officialRank"]["all"]:
        IF5 = float(data["data"]["officialRank"]["all"]["sciif5"])
    # 查看data["data"]["officialRank"]["all"]下是否有sciUp字段
    if "sciUp" in data["data"]["officialRank"]["all"]:
        sciUp = data["data"]["officialRank"]["all"]["sciUp"]
    return (IF, IF5, sciUp)


def clear_link(link):
    query = urlparse(link).query
    params = parse_qs(query)
    if "url" in params:
        return params["url"][0]
    else:
        return link


def clear_tag(tag):
    if tag != None:
        if (" - new related research" in tag) | (related_tag in tag):
            tag = tag.replace(" - new related research", "")
            tag = tag.replace(related_tag, "")
            tag_type = "related"
        elif ("new articles" in tag) | (new_articles_tag in tag):
            tag = tag.replace(" - new articles", "")
            tag = tag.replace(new_articles_tag, "")
            tag_type = "new-articles"
        elif ("to articles by " in tag) | (new_citations_tag in tag):
            # 取new citations to articles by 后面的内容
            tag = tag.split("to articles by ")
            if len(tag) == 1:
                tag = tag[0].split(new_citations_tag)
                tag = tag[new_citations_num]
            else:
                tag = tag[1]
            tag_type = "new-citations"
        elif ("new results" in tag) | (new_results_tag in tag):
            tag = tag.replace(" - new results", "")
            tag = tag.replace(new_results_tag, "")
            tag_type = "subject"
        else:
            tag_type = "unknown"
    else:
        tag_type = "empty"
    return (tag, tag_type)


def get_title(title):
    titles = []
    from crossref.restful import Journals

    if "…" in title:
        # 替换
        title = title.replace("…", "")
        # 替换掉标题中的\xa0
        title = title.replace("\xa0", "")
        # 如果最后一位是空格，去掉
        if title[-1] == " ":
            title = title[:-1]
        # 如果第一位是空格，去掉
        if title[0] == " ":
            title = title[1:]
        titleold = title
        if "&" in title:
            title = title.replace("&", "")
    try:
        journals = Journals().query(title)
        if ESkeys == "":
            time.sleep(0.05)
        for item in journals:
            if titleold in item["title"]:
                titles.append(item["title"])
        return titles
    except:
        print("\n获取期刊名称失败:请检查网络连接")
        return titles


def journal_to_IF(journal):
    if ESkeys == "":
        IF = [0, 0, ""]
        return (IF, journal)
    if "…" in journal:
        journalnames = get_title(journal)
        if len(journalnames) == 1:
            journal = journalnames[0]

            IF = get_IF(journal)
        elif len(journalnames) > 1:
            addstr = ""
            if len(journalnames) > 3:
                addstr = "..."
            print(f"\n期刊{journal}存在多个匹配项：{journalnames[0:3]} {addstr}")
            IF = [0, 0, ""]
        else:
            print(f"\n期刊{journal}未找到匹配项")
            IF = [0, 0, ""]
    else:
        IF = get_IF(journal)
    return IF, journal


def get_articles(htmltext, tagtext, tag_type, rtime):
    articles = []
    tag_typeold = tag_type
    if tag_type == "unknown":
        tag_typeold = "unknown"
        tag_type = "subject"
    soup = BeautifulSoup(htmltext, "html.parser")
    # 找到所有h3标签且标签下包含a标签
    h3_tags = soup.find_all("h3")
    for tag in h3_tags:
        a_tag = tag.find("a")
        if a_tag:
            # 找到该h3标签后的第一个div标签
            div_tag = tag.find_next("div")
            div_text = div_tag.text
            # 以 - 分割，前一项为作者信息
            div_text = div_text.split("- ")
            # 判断第二项中有无逗号
            if len(div_text) == 1:
                date = ""
                journal = ""
                IF = [0, 0, ""]
            else:
                if "," in div_text[1]:
                    div_texti = div_text[1].rsplit(", ", 1)
                    date = div_texti[1]
                    journal = div_texti[0]
                    IF, journal = journal_to_IF(journal)

                else:
                    # 匹配年份
                    import re

                    pattern = re.compile(r"\d{4}")
                    match = pattern.search(div_text[1])
                    if match:
                        date = match.group()
                        # 其余为期刊
                        journal = div_text[1].replace(date, "")
                        if journal == "":
                            IF = [0, 0, ""]
                        else:
                            IF, journal = journal_to_IF(journal)
                    else:
                        date = ""
                        journal = div_text[1]
                        IF, journal = journal_to_IF(journal)
                    # 替换掉标题中的\xa0
            title = tag.text.replace("\xa0", " ")
            article = {
                "title": title,
                "author": div_text[0][:-1],
                "journal": journal,
                "date": rtime + "-" + date,
                "link": clear_link(a_tag["href"]),
                "IF": IF[0],
                "IF5": IF[1],
                "sciUp": IF[2],
                "tags": tagtext,
                "tag_type": tag_type,
            }
            articles.append(article)
    if (len(articles) != 0) & (tag_typeold == "unknown"):
        print("未知标签：", tagtext)
    return articles


app = Flask(__name__)


@app.route("/")
def index():
    data = myaccount.get_database_paper()
    return render_template("index.html", data=data)


# 更新单个期刊影响因子线路 用POST方法，传入期刊名称
@app.route("/updateone", methods=["POST"])
def updateone():
    if request.method == "POST":
        if ESkeys == "":
            print(" * 未设置Easy Scholar key，不会获取影响因子")
            # 报错
            raise Exception("未设置Easy Scholar key，不会获取影响因子")
        url = request.form.get("url")
        # 从数据库中获取期刊名称
        conn = sqlite3.connect(myaccount.databasename)
        c = conn.cursor()
        c.execute("SELECT * FROM paper WHERE link=?", (url,))
        rows = c.fetchall()
        conn.close()
        if len(rows) == 0:
            # 抛出错误
            return "error"
        journal = rows[0][2]
        if journal == "":
            return {"IF": 0, "IF5": 0, "sciUp": "", "journal": ""}
        IF, journal = journal_to_IF(journal)
        # 修改数据库
        conn = sqlite3.connect(myaccount.databasename)
        c = conn.cursor()
        c.execute("UPDATE paper SET IF=? WHERE link=?", (IF[0], url))
        c.execute("UPDATE paper SET IF5=? WHERE link=?", (IF[1], url))
        c.execute("UPDATE paper SET sciUp=? WHERE link=?", (IF[2], url))
        c.execute("UPDATE paper SET journal=? WHERE link=?", (journal, url))
        conn.commit()
        conn.close()

        return {"IF": IF[0], "IF5": IF[1], "sciUp": IF[2], "journal": journal}
    else:
        return "error"


# 更新期刊影响因子线路
@app.route("/update")
def update():
    myaccount.updatajournal()
    return "更新完成"


# 立即检查新邮件线路
@app.route("/check")
def check():
    myaccount.check_for_new_email_stream()
    return "检查完成"


# 导出excel线路
@app.route("/export")
def export():
    import csv

    data = myaccount.get_database_paper()
    columns = [
        "title",
        "author",
        "journal",
        "date",
        "link",
        "IF",
        "IF5",
        "sciUp",
        "new_article",
        "new_citation",
        "related",
        "subject",
    ]
# 创建新的 columns 列表，去除 "IF", "IF5", 和 "sciUp"
    new_columns = [
        "title",
        "author",
        "journal",
        "date",
        "link",
        "new_article",
        "new_citation",
        "related",
        "subject",
    ]

# 获取需要保留的列的索引
    indices_to_keep = [columns.index(col) for col in new_columns]

    # 过滤数据，只保留需要的列
    filtered_data = [
        tuple(row[i] for i in indices_to_keep) for row in data
    ]
        
    # 保存到csv文件
    filename = f"Alerts_data.csv"

    with open(filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(new_columns)
        writer.writerows(filtered_data)



if __name__ == "__main__":
    print(" * 程序启动！\n")

    threading.Thread(target=check_proxy, daemon=True).start()
    create_database()
    myaccount = EmailAccount(username, password)

    def loop():
        while True:
            myaccount.check_for_new_email_stream()
            time.sleep(interval)

    if myaccount.init_success:
        t = threading.Thread(target=loop, daemon=True, name="LoopThread")
        t.start()
        export()
        app.run(debug=True, use_reloader=False, port=port, host='127.0.0.1')
    else:
        print("初始化失败：您的网络或账户存在问题，无法登录邮箱！")
        os._exit(0)
