import json
import time
import os
import sys

from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QTableWidgetItem, QAbstractItemView, QMessageBox
from datetime import datetime,timedelta

from gui.MainWindow import Ui_MainWindow
from utils import public

from imapclient import IMAPClient
#-------------------------------------------------------------------------
# 获取邮件列表的线程
class GetMailsThread(QThread):
    sig_mails = pyqtSignal(list)
    sig_process = pyqtSignal(str)
    def __int__(self):
        # 初始化函数
        super(GetMailsThread, self).__init__()
    def set_args(self, mm):
        self.mm = mm
    def run(self):
        #print(self.mm)
        mails = public.get_mail_list(self.mm[0], self.mm[1], self.mm[2], self.sig_process)
        self.sig_mails.emit(mails)
class InitIMAPConnect(QThread):
    sig_conn = pyqtSignal(list)
    def __init__(self):
        super(InitIMAPConnect, self).__init__()
    def set_args(self, server):
        self.server = server
        #print(self.conn)
    def run(self):
        email_server = IMAPClient(self.server)
        self.sig_conn.emit([email_server])
class LoginMailThread(QThread):
    sig_server = pyqtSignal(list)
    def __init__(self):
        super(LoginMailThread, self).__init__()
    def set_args(self, config):
        self.configs = config
    def run(self):
        self.email_server, self.login_status_msg, self.login_id_set_msg, self.sel_f_msg = public.login_mail(self.configs)
        self.sig_server.emit([self.email_server, self.login_status_msg, self.login_id_set_msg, self.sel_f_msg])


#--------------------------------------------------------------------------
class PyQtMainEntry(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(PyQtMainEntry, self).__init__(parent)
        #--初始化UI组件--
        self.setupUi(self)
        self.start_time_set.dateTimeChanged.connect(lambda: self.start_time_check.setChecked(True))
        self.end_time_set.dateTimeChanged.connect(lambda: self.end_time_check.setChecked(True))
        #--初始化表格--
            # 5列
        self.mail_list_table.setColumnCount(5)
            # 表格列宽
        self.mail_list_table.setColumnWidth(0, 50)
        #self.mail_list_table.setColumnWidth(1, 100)
        self.mail_list_table.setColumnWidth(1, 150)
        self.mail_list_table.setColumnWidth(2, 300)
        self.mail_list_table.setColumnWidth(4, 50)
            # 列标题
        self.mail_list_table.setHorizontalHeaderLabels(['uid', 'date', '主题', '发件邮箱', '附件数'])
        #--收件工具页面--
        self.import_config_btn.clicked.connect(self.conn_mail)
            # 附件保存路径
        self.save_file_path_info.setText(os.getcwd().replace('\\', '/')+'/download')
        self.save_file_path_btn.clicked.connect(self.set_file_path)
        self.open_dir_btn.clicked.connect(self.open_dir)
        self.rcv_mail_btn.clicked.connect(self.get_mail_list)  # 下载邮件列表
        self.down_file_btn.clicked.connect(self.down_files)  # 保存文件
        self.import_teamer_btn.clicked.connect(self.import_namelist)  # 导入需要发邮件的全体成员名单
        self.exp_file_info_btn.clicked.connect(self.export_names)
        self.push2wxMsgSender_btn.clicked.connect(self.push2wxMsgSender)
        #--微信群发助手页面--
        self.test_wx_btn.clicked.connect(self.test_wx)
        self.import_receiver_btn.clicked.connect(self.import_receiver)
        self.match_wxFriend_btn.clicked.connect(self.match_wxFriend)
        self.send_msg_btn.clicked.connect(self.send_wx_msg)
        #--设置页面--
        self.refresh_config_btn.clicked.connect(self.load_config)
        self.save_config_btn.clicked.connect(self.save_config)

    # -----------------------收件工具页面-----------------------
    ## 1.测试邮箱连接
    def conn_mail(self):
        config_file = './config/config.json'
        if not os.path.exists(config_file):
            self.conn_label.setText("无邮箱信息，请在设置页面填写后保存")
            self.statusbar.showMessage('无邮箱信息，请在设置页面填写后保存')  # 显示状态栏信息
            return None
        else:
            with open(config_file) as f:
                self.configs = json.load(f)
            self.statusbar.showMessage('(1/4)邮箱信息加载成功，正在连接邮件服务器')  # 显示状态栏信息
            self.curr_mail_add_info.setText(self.configs['mail'])
            #return None
        try:
            # 连接imap服务器
            #self.conn_server_thread = InitIMAPConnect()
            #self.conn_server_thread.set_args(self.configs['server'])
            #self.conn_server_thread.sig_conn.connect(self.conn_server)
            #self.conn_server_thread.start()
            #if self.conn_server_thread.isFinished():
            #    self.conn_server_thread.exit()
            self.email_server = IMAPClient(self.configs['server'])
            self.statusbar.showMessage(f'(2/4)连接服务器{self.configs["server"]}:{self.configs["port"]}成功')  # 显示状态栏信息
        except Exception as e:
            self.statusbar.showMessage(f"(2/4)连接服务器{self.configs['server']}:{self.configs['port']}异常，请检查网络或设置")  # 显示状态栏信息
            self.conn_label.setText('连接服务器异常，请检查网络或设置')
            print(e)
            return None
        try:
            # 验证用户邮箱
            login_msg = self.email_server.login(self.configs['mail'], self.configs['pwd'])
            mail_id_set_msg = self.email_server.id_({"name": "IMAPClient", "version": "2.1.0"})
            self.statusbar.showMessage(f"(4/4)用户邮箱{self.configs['mail']}验证登陆成功")  # 显示状态栏信息
            self.conn_label.setText(f'验证成功')
            self.email_server.logout()
            self.mail_verified = True
        except:
            self.statusbar.showMessage(f"(4/4)用户邮箱{self.configs['mail']}密码错误，请检查设置")  # 显示状态栏信息
            self.conn_label.setText('密码错误，请检查设置')
            self.email_server.logout()
            return None
    def conn_server(self, server):
        self.email_server = server[0]
    ## 2.设置附件保存位置
    def set_file_path(self):
        """if self.save_file_path_info.toPlainText():
            path_old = self.save_file_path_info.toPlainText()
            if path_old[-1] != '/':
                path_old += '/'
        else:
            path_old = os.getcwd().replace('\\', '/')+'/download'"""
        path_old = self.save_file_path_info.toPlainText()
        path = QFileDialog.getExistingDirectory(None, "选取文件夹", path_old)
        if not path:
            self.save_file_path_info.setText(path_old)
        else:
            self.save_file_path_info.setText(path)
    ## 3.读取邮件列表
    def get_mail_list(self):
        try:
            if self.mail_verified:
                pass
        except:
            self.chk_label.setText('未验证！')
            msg_box = QMessageBox(QMessageBox.Critical, '邮箱尚未验证', '邮箱尚未验证，请点击“连接邮箱"')
            msg_box.exec_()
            return None
        #print(1, datetime.now().strftime("%H%M%S"))
        # ________设置部分________
        ## 设置收件关键词
        self.key_word = self.rcv_kwd_edit.text()
        if not self.key_word:
            self.key_word = False
        ## 设置筛选               待优化：目前尚不能加入subject，使用subject筛选后可加快处理速度
        self.mail_select = {}
        #print(2, datetime.now().strftime("%H%M%S"))
        # 是否设定起始时间
        set_start_time = self.start_time_check.isChecked()  # 是否设定起始时间
        if set_start_time:
            s_t = self.start_time_set.dateTime()
            start_time = public.process_qt_time(s_t)
            if start_time > datetime.now():
                msg_box = QMessageBox(QMessageBox.Critical, '时间设定错误', '设置筛选时间错误！起始时间不能晚于截止时间')
                msg_box.exec_()
                return None
        else:
            start_time = datetime(2022, 9, 3)
        #print(3, datetime.now().strftime("%H%M%S"))
        # 是否设定截止时间
        set_end_time = self.end_time_check.isChecked()
        if set_end_time:
            e_t = self.end_time_set.dateTime()
            end_time = public.process_qt_time(e_t)
            if start_time > end_time:
                msg_box = QMessageBox(QMessageBox.Critical, '时间设定错误', '设置筛选时间错误！起始时间不能晚于截止时间')
                msg_box.exec_()
                return None
        else:
            end_time = datetime.now()
        #print(4, datetime.now().strftime("%H%M%S"))
        self.mail_select = {'start': start_time, 'end': end_time}
        self.time_span = [start_time, end_time]
        self.statusbar.showMessage(f'收件参数设置完成，正在获取信件信息')  # 显示状态栏信息
        # ________测试部分________
        #return None
        # ________收件部分________
        try:
        # 在后台获取邮件列表
            m = True
            #self.login_thread = LoginMailThread()
            #self.login_thread.set_args(self.configs)
            #self.login_thread.sig_server.connect(self.login)
            #self.login_thread.start()
            self.email_server, login_status_msg, login_id_set_msg, sel_f_msg = public.login_mail(self.configs)
        except Exception as e:
            m = False
            self.chk_label.setText('登录失败')
            self.statusbar.showMessage('[错误]邮箱登录失败:'+str(e))  # 显示状态栏信息
            msg_box = QMessageBox(QMessageBox.Critical, '邮箱登录失败', '邮箱登录失败:\n'+str(e))
            msg_box.exec_()
        if m:
            self.mails_thread = GetMailsThread()  # 创建获取邮件列表的线程实例
            self.mails_thread.set_args([self.email_server, self.mail_select, self.key_word])
            #绑定线程信息流
            self.mails_thread.sig_process.connect(self.mail_process)
            self.mails_thread.sig_mails.connect(self.show_mails)  # 连接动作，将执行结果作为参数传给本类下的show_mails方法
            #self.mails_thread.connect(pyqtSignal("finished()"), self.mails_done)  # 连接动作，表明结束后执行本类下的mails_done方法
            self.mails_thread.start()  # 启动线程
            # 开始执行任务后，对主窗口的处理
            self.import_config_btn.setDisabled(True)
            self.rcv_mail_btn.setDisabled(True)
            self.down_file_btn.setDisabled(True)
            self.exp_file_info_btn.setDisabled(True)
            self.push2wxMsgSender_btn.setDisabled(True)



    def login(self, server):
        self.email_server = server[0]
    # 接收get_mails_list方法中，实例mails_thread下信号m_thread传来的参数mails_list，并显示
    def show_mails(self, mails_list):
        self.mails_list = mails_list
        n_mails = len(self.mails_list)
        note = f'获取邮件信息成功，共{n_mails}封邮件'
        self.statusbar.showMessage(note)
        self.import_config_btn.setEnabled(True)
        self.rcv_mail_btn.setEnabled(True)
        self.down_file_btn.setEnabled(True)
        self.exp_file_info_btn.setEnabled(True)
        self.push2wxMsgSender_btn.setEnabled(True)
        print(note)
        if n_mails == 0:
            return None
        self.mail_list_table.setRowCount(0)
        for i in range(n_mails):
            num = self.mail_list_table.rowCount()
            self.mail_list_table.insertRow(num)
            msg = mails_list[i]
            keys = ['uid', 'send_time', 'subject', 'sender', 'file_num']
            for j in range(5):
                self.mail_list_table.setItem(i, j, QTableWidgetItem(str(msg[keys[j]])) )
        self.mail_list_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
    def mail_process(self, note):
        self.statusbar.showMessage(note)
    # 4.附件下载
    def down_files(self):
        try:
            n_mails = len(self.mails_list)
        except:
            msg_box = QMessageBox(QMessageBox.Critical, '错误', '请先获取邮件列表')
            msg_box.exec_()
            return None
        if n_mails == 0:
            msg_box = QMessageBox(QMessageBox.Information, '提示', '没有附件可保存')
            msg_box.exec_()
            return None
        print(f'初始化附件保存')
        path = self.save_file_path_info.toPlainText()
        now = datetime.now().strftime("%Y%m%d %H%M%S")
        pathname = path + f'/{now}下载'
        start = self.time_span[0].strftime("%Y%m%d")
        end = (self.time_span[1]-timedelta(days=1)).strftime("%Y%m%d")
        pathname += f'__{start}-{end}'
        if self.key_word:
            pathname += f'__{self.key_word}'
        else:
            pathname += '__无主题'
        os.mkdir(pathname)
        for i in range(n_mails):
            print(f'正在保存附件，第{i}封，共{n_mails}封')
            self.statusbar.showMessage(f'正在保存附件，第{i}封，共{n_mails}封')
            m = self.mails_list[i]
            if self.multi_files_mode_check.isChecked():
                pathname1 = pathname + f'/{m["subject"]} [{m["sender"]}]'
                os.mkdir(pathname1)
            else:
                pathname1 = pathname
            messageObj = m['mail_obj']
            for part in messageObj.walk():
                name = part.get_filename()
                if name != None:
                    name = pathname1 + "/" + public.decode_str(name)
                    with open(name, 'wb') as f:
                        f.write(part.get_payload(decode=True))
        self.statusbar.showMessage(f'附件保存完成')
        msg_box = QMessageBox(QMessageBox.Information, '保存完成', f'附件保存完毕，共{n_mails}封邮件')
        msg_box.exec_()
    def open_dir(self):
        os.startfile(self.save_file_path_info.toPlainText())
        #os.system(f"start explorer {self.save_file_path_info.toPlainText()}")
    # 5.处理下载的文件
    def import_namelist(self):
        try:
            path_old = self.teamer_file_info.toPlainText()
        except:
            path_old = './'
        path,_ = QFileDialog.getOpenFileName(self, "选择成员名单", "./", "表格型(*.xlsx;*.xls;*.csv);;所有类型(*)")
        if not path:
            self.teamer_file_info.setText(path_old)
        else:
            self.teamer_file_info.setText(path)
    def export_names(self):
        msg_box = QMessageBox(QMessageBox.Information, '功能开发中', '抱歉，功能还在开发')
        msg_box.exec_()
    def push2wxMsgSender(self):
        msg_box = QMessageBox(QMessageBox.Information, '功能开发中', '抱歉，功能还在开发')
        msg_box.exec_()
    # -----------------------微信群发助手页面-----------------------
    def test_wx(self):
        msg_box = QMessageBox(QMessageBox.Information, '功能开发中', '抱歉，功能还在开发')
        msg_box.exec_()
    def import_receiver(self):
        msg_box = QMessageBox(QMessageBox.Information, '功能开发中', '抱歉，功能还在开发')
        msg_box.exec_()
    def match_wxFriend(self):
        msg_box = QMessageBox(QMessageBox.Information, '功能开发中', '抱歉，功能还在开发')
        msg_box.exec_()
    def send_wx_msg(self):
        msg_box = QMessageBox(QMessageBox.Information, '功能开发中', '抱歉，功能还在开发')
        msg_box.exec_()
    # -----------------------设置页面-----------------------
    ## 加载目前已有的设置
    def load_config(self):
        config_file = './config/config.json'
        if not os.path.exists(config_file):
            self.statusbar.showMessage('无邮箱信息，请填写后保存')  # 显示状态栏信息
        else:
            with open(config_file) as f:
                configs = json.load(f)
            self.mail_add_edit.setText(configs['mail'])
            self.mail_pwd_edit.setText(configs['pwd'])
            self.server_edit.setText(configs['server'])
            self.port_edit.setText(configs['port'])
            self.statusbar.showMessage('配置信息已加载')  # 显示状态栏信息
    ## 保存设置
    def save_config(self):
        config_file = './config/config.json'
        configs = {'mail': self.mail_add_edit.toPlainText(),
                   'pwd': self.mail_pwd_edit.toPlainText(),
                   'server': self.server_edit.toPlainText(),
                   'port': self.port_edit.toPlainText()
                   }
        with open(config_file, 'w') as f:
            f.write(json.dumps(configs))
        self.statusbar.showMessage('保存配置成功')  # 显示状态栏信息


#------------------------------------------------------------------------
if __name__ == '__main__':
    public.init_dev()  # 初始化软件

    app = QApplication(sys.argv)
    MainWindow = PyQtMainEntry()
    MainWindow.show()

    sys.exit(app.exec_())

