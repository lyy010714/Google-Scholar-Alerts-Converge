
看到邮箱里堆积的一百多封谷歌学术提醒邮件非常头痛，于是用混乱的代码写了这个程序。
也许勉强能提升一点查看这些邮件的体验，主要是可以去除大量的重复论文，结合浏览器自带的翻译软件可以快速翻译。

# 启动方法
- git clone https://github.com/St-ZFeng/Google-Scholar-Alerts-Converge.git
- 或者下载releases中的压缩包并解压（压缩包仅支持Windows系统）
- 下载python及requirements.txt中的模块（下载releases不需要）
- 打开并修改config.yaml文件(注意引号之后必须加空格，除username和password外，其它参数可不提供或保持默认状态)

```yaml
ESkeys: #easy scholar api key; 可从https://www.easyscholar.cc获取
address: scholaralerts-noreply@google.com #谷歌学术提醒的邮件地址
interval: 120 #检查新邮件的间隔时间，秒为单位
username: #邮箱的用户名
password: #邮箱的密码或者登录码（QQ邮箱需要在设置中额外申请）
imap_server: imap.qq.com #如果以imap方式访问，需要提供服务器地址。默认为QQ邮箱的服务器地址，对于outlook等邮箱以exchange方式访问则不需要
host: 127.0.0.1 #网页显示的本地服务器ip
port: 5000 #网页显示的本地服务器端口

# 以下参数用于归类邮件主题，英文版本已经被内置，可再多添加一种语言支持
related_tag:  #相关论文提醒的主题特有字符串，与人名紧挨，例如中文为"- 新的相关研究工作"
new_articles_tag:  #新文章提醒的主题特有字符串，与人名紧挨，例如中文为"- 新的文章"
new_citations_tag: #新引用提醒的主题特有字符串，与人名紧挨，例如中文为"的文章新增了"
new_citations_num: #新引用提醒的主题按照特有字符串拆分后，哪一部分为人名：前取0，后取1
new_results_tag:  #关键词论文提醒的主题特有字符串，与关键字紧挨，例如中文为"- 新的结果"
```

- 设置完成后运行python代码或.exe
- 等待初始化完成，如果你的邮箱中有大量谷歌学术提醒邮件，会花费很长时间
- 命令窗口显示Running on http://ip:port (默认为http://127.0.0.1:5000)
- 在浏览器中打开http://ip:port (使用过程中请勿关闭命令窗口)

# 页面功能

前端页面99%的部分由chatgpt编写，所以功能并不是很全面，主打一个够用就行。

- Export: 导出所有条目的csv文件
- Update IF: 尝试更新所有条目的影响因子（需要提供ESkeys，由于邮件中的很多期刊名不完整，所以一部分期刊无法获取影响因子）
- Check Now: 立即检查是否有新邮件并导入
- Update: 尝试更新特定条目的影响因子（需要提供ESkeys）
- Search: 检索特定字符串，支持&语法，示例：Title==group & Related with==Eliot Smith，这会给出标题包含group且与Eliot Smith的工作相关的论文。