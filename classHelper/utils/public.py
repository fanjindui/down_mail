import os
import re
from datetime import datetime,timedelta
from email.header import decode_header
import pyzmail  # pip install pyzmail36
from imapclient import IMAPClient

# 初始化程序
def init_dev():
    now_path = os.getcwd()
    if os.path.exists(now_path+'/config'):
        pass
    else:
        os.mkdir(now_path+'/config')
    if os.path.exists(now_path+'/download'):
        pass
    else:
        os.mkdir(now_path+'/download')
    if os.path.exists(now_path+'/cache'):
        pass
    else:
        os.mkdir(now_path+'/cache')


def login_mail(configs):
    #print(1, datetime.now().strftime("%H%M%S"))
    email_server = IMAPClient(configs['server'])
    #print(2, datetime.now().strftime("%H%M%S"))
    login_status_msg = email_server.login(configs['mail'], configs['pwd'])  # 登录
    #print(3, datetime.now().strftime("%H%M%S"))
    login_id_set_msg = email_server.id_({"name": "IMAPClient", "version": "2.1.0"})  # 网易邮箱安全设置
    #print(4, datetime.now().strftime("%H%M%S"))
    sel_f_msg = email_server.select_folder("INBOX")  # 设置收件箱
    return email_server, login_status_msg, login_id_set_msg, sel_f_msg

# 将qt原生时间转为datetime
def process_qt_time(t):
    tt = t.toString('yyyy-MM-dd hh:mm:ss')
    return datetime.strptime(tt, '%Y-%m-%d %H:%M:%S')
    return datetime.strptime(tt, '%Y-%m-%d %H:%M:%S')


# 处理中文乱码问题
def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

# 处理邮件正文
def process_mail_content(m):
    pattern = re.compile(r'//[0-9A-Za-z./\\?\\=\\:]+')


# 筛选生成邮件列表
def get_mail_list(server, select, key_word, sig):
    selection = f"SINCE {select['start'].strftime('%d-%b-%Y')} BEFORE {(select['end']+timedelta(days=1)).strftime('%d-%b-%Y')}"
    #if select['subject']:   # 找到筛选主题的方法后
    #print(select)
    messages = server.search(selection)  # 按照条件获取邮件uid列表
    index = len(messages)
    mails = []
    nn = 0
    for uid in messages:  # 倒序遍历邮件，这样取到的第一封就是最新邮件
        nn += 1
        print(f'正在下载第{str(nn)}封...', end='')
        messageList = server.fetch(uid, ["RFC822"])
        mailBody = messageList[uid][b"RFC822"]
        messageObj = pyzmail.PyzMessage.factory(mailBody)  # 邮件的原始文本:# lines是邮件内容，列表形式使用join拼成一个byte变量
        # imapClient只能找出日子，精确不到分钟
        send_date = (server.fetch(uid, ['ENVELOPE']))[uid][b'ENVELOPE'].date  # 时间
        #print(type(send_date), type(select['start']), type(select['end']))
        if send_date<select['start'] or send_date>select['end']:
            sig.emit(f'第{str(nn)}封邮件时间不符合要求')
            continue
        # 挑选合适主题，不符合条件的话直接下一封
        subject = messageObj.get_subject()
        if subject == '':
            subject = '(无主题)'
        print(subject)
        if key_word:
            if key_word not in subject:
                sig.emit(f'第{str(nn)}封邮件主题不符合要求')
                continue
        # 解析邮件
        sender = messageObj.get_addresses('from')  # 发件人

            # 正文与附件
        if messageObj.html_part:
            htmlContent = messageObj.html_part.get_payload().decode(messageObj.html_part.charset)
        else:
            htmlContent = ''
        if messageObj.text_part:
            try:
                textContent = messageObj.text_part.get_payload().decode(messageObj.text_part.charset)
            except:
                textContent = messageObj.text_part.get_payload().decode('utf-8')
        else:
            textContent = ''
        ctnt = [htmlContent, textContent] #正文内容
            # 附件名称
        files_list = []
        for part in messageObj.walk():
            name = part.get_filename()
            if name != None:
                name = decode_str(name)
                files_list.append(name)
        mails.append({'uid': uid,
                      'send_time': send_date,
                      'subject': subject,
                      'sender': sender[0][1], 'send_name': sender[0][0],
                      'file_num': len(files_list),
                      'files': files_list,
                      'content': ctnt,
                      'mail_obj': messageObj
                      })
        sig.emit(f'已下载第{str(nn)}封邮件信息 [{subject}（{sender[0][1]}）]，正在下载第{str(nn+1)}封，共{str(index)}封')
    return mails
