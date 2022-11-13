import re
from tkinter import *  # 图形界面
from tkinter import messagebox  # 对话框
from tkinter import filedialog  # 文件操作
import windnd
import os
import time
import sys
from tkinter import ttk  # 导入内部包
import configparser
import smtplib
import threading
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
from email.mime.application import MIMEApplication
from datajosn import save_data  # 用于保存数据
from datajosn import read_data  # 用于读取数据
from picter import icocb
from showemail import show_information
import pythoncom
from win32com.shell import shell

# json数据文件名称/地址
json_data_file = 'emaildata.json'  # 通讯录及已发送邮件存储
cong_ini = "userconfig.ini"     # 配置信息存储
# 初始化数据
data = {
    'email': [],  # 邮箱
    'record': [],
    'display_user': True,  # 是否显示用户名
    'program_path': sys.argv[0],  # 程序路径
    'author': '程甲第',  # 制作人
    'time': time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())  # 时间
}


class Right_Click_Menus:
    """创建一个弹出菜单"""

    def __init__(self, text, undo=True):
        menu = Menu(text, tearoff=False)
        menu.add_command(label="剪切", command=lambda: text.event_generate('<<Cut>>'))
        menu.add_command(label="复制", command=lambda: text.event_generate('<<Copy>>'))
        menu.add_command(label="粘贴", command=lambda: text.event_generate('<<Paste>>'))
        menu.add_command(label="删除", command=lambda: text.event_generate('<<Clear>>'))
        if undo:
            menu.add_command(label="撤销", command=lambda: text.event_generate('<<Undo>>'))
            menu.add_command(label="重做", command=lambda: text.event_generate('<<Redo>>'))

        def popup(event):
            menu.post(event.x_root, event.y_root)  # post在指定的位置显示弹出菜单

        text.bind("<Button-3>", popup)


def unite_login(usernames, password, smtp_server, port: int, opat: int = 0):
    smtp_server = smtp_server
    try:
        server = smtplib.SMTP(smtp_server, port)
        server.set_debuglevel(1)
    except smtplib.SMTPServerDisconnected:
        server = smtplib.SMTP_SSL(host=smtp_server, port=port)
        server.set_debuglevel(1)
        # 登陆邮箱
    except Exception as e:
        if str(e) == '[Errno 11001] getaddrinfo failed':
            messagebox.showerror(title="登录失败", message=f"网络连接异常，请检查网络连接\n错误：{e}")
        alter_config(option="have_log", values="False")
        server = False
    try:
        server.login(usernames, password)
        alter_config(option="have_log", values="True")
        if opat == 0:
            """用于软件打开时无感登录"""
            server.quit()
            return True
        elif opat == 1:
            """用于登录提示"""
            messagebox.showinfo(title="", message="登录成功!")
            server.quit()
            return True
        elif opat == 2:
            """用于发送邮件"""
            return server
    except smtplib.SMTPAuthenticationError:
        messagebox.showerror(title="错误",
                             message=f"登录失败，请检查邮箱和授权码是否正确\n您输入的授权码为<{read_config(option='password')}>")
        alter_config(option="have_log", values="False")
        return False
    except AttributeError:
        return False
    except smtplib.SMTPServerDisconnected:
        return False


def creat_config():
    """创建配置文件"""
    config_file = configparser.ConfigParser()
    config_file.add_section("User")
    config_file.set("User", "userName", "")
    config_file.set("User", "password", "")
    config_file.set("User", "senderName", "")
    config_file.set("User", "smtp_server", "")
    config_file.set("User", "port", "25")
    config_file.set("User", "have_log", "False")
    config_file.set("User", "pasw_start", "False")
    config_file.set("User", "pasw_word", "1234567")
    with open(cong_ini, 'w', encoding='utf-8') as configfileObj:
        config_file.write(configfileObj)
        configfileObj.flush()
        configfileObj.close()


def alter_config(option, values, section="User"):
    """修改配置文件"""
    config_file = configparser.ConfigParser()
    config_file.read(cong_ini, encoding='utf-8')
    config_file[section][option] = values
    with open(cong_ini, 'w', encoding='utf-8') as configfileObj:
        config_file.write(configfileObj)
        configfileObj.flush()
        configfileObj.close()


def read_config(option, section="User"):
    config_file = configparser.ConfigParser()

    config_file.read(cong_ini, encoding='utf-8')
    value = config_file.get(section=section, option=option)
    return value


def save_send_record(receiver, emtheme, emtype, content, addfiles, sender, accor):
    """记录发送的邮件"""
    send_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    rece_pers = ""
    rece_ema = ""
    all_receive = []
    recelis = receiver.split(",")
    recelis = [i for i in recelis if i != '']
    print(recelis)
    for item in tree_emile.get_children():
        print(item)
        one_rece = tree_emile.set(item, 1)
        print(one_rece, type(one_rece))
        all_receive.append(one_rece)
        if one_rece in recelis:
            print(one_rece, 1)
            xcp = f"({tree_emile.set(item, 0)}, {tree_emile.set(item, 1)}),"
            rece_ema += xcp
            rece_pers += tree_emile.set(item, 0)
            rece_pers += ","
    for ii in recelis:
        if ii not in all_receive:
            xcp = f"(未备注, {ii}),"
            rece_ema += xcp
            rece_pers += "未备注"
            rece_pers += ","

    record = {"发送时间": send_time, "收件人": rece_pers, "收件邮箱": rece_ema, "邮件主题": emtheme, "邮件类型": emtype,
              "邮件正文": content, "附件": addfiles, "发件人": sender, "抄送人": accor}
    data['record'].append(record)
    save_data(data, json_data_file)
    tree_emile_record.insert("", END, values=(send_time, rece_pers))


def show_detailed_email(*args):
    ssl = tree_emile_record.selection()
    fujian = ""
    if len(ssl) > 0:
        sl = ssl[0]
        itm = tree_emile_record.set(sl)
        for one_cord in data['record']:
            if one_cord["发送时间"] == itm["发送时间"] and one_cord["收件人"] == itm["收件人"]:
                if len(one_cord["附件"]) <= 0:
                    fujian = "无附件"
                else:
                    for ii in one_cord["附件"]:
                        xyu = f"({ii[0]},{ii[1]})"
                        fujian += xyu
                        fujian += "\n"
                show_information(information=f"""                                                 邮件详情
发件人：{one_cord["发件人"]}
收件人：{one_cord["收件邮箱"]}
抄送人：{one_cord["抄送人"]}\n
发送时间：{one_cord["发送时间"]}
邮件主题：{one_cord["邮件主题"]}
邮件类型：{one_cord["邮件类型"]}\n
邮件正文：\n\n{one_cord["邮件正文"]}
\n\n\n
附件：\n{fujian}
                """)


def open_log_in(*args):
    """打开软件自动登录，检测网络状态和账号密码的准确性"""
    unite_login(
        usernames=read_config(option="userName"),
        password=read_config(option="password"),
        smtp_server=read_config(option="smtp_server"),
        port=int(),
        opat=0)
    zlkl = eval(read_config(option="have_log"))
    if zlkl:
        window.update()
        login_button.forget()
        send_button.config(state="normal", text="发送", background="#8aa58b", fg="white")
        send_email_lable.pack(side=LEFT)
        sender_email_lable.pack(side=LEFT)
        kong_1_lable.pack(side=LEFT)
        send_name_lable.pack(side=LEFT)
        send_name_enter.pack(side=LEFT, expand=YES, fill=X)
        send_name_enter.delete(0, END)
        send_name_enter.insert(0, read_config(option="senderName"))
        updata_user_button.pack(side=LEFT, padx=5, pady=5)
        quit_user_button.pack(side=LEFT, padx=10, pady=5)

    else:
        window.update()
        login_button.pack(side=LEFT, expand=YES, fill=X)
        send_button.config(state="disabled", text="请先进行登录", background="#8aa58b", fg="white")
        send_email_lable.forget()
        sender_email_lable.forget()
        updata_user_button.forget()
        quit_user_button.forget()
        send_name_lable.forget()
        send_name_enter.delete(0, END)
        send_name_enter.forget()
        kong_1_lable.forget()


def dragged_files(files):
    for msge in files:
        msg = msge.decode('gbk')
        if os.path.exists(msg):
            os.path.realpath(msg)
            msg_on = msg.split("\\")[-1]
            tree_files.insert("", END, values=(msg_on, msg))
        else:
            messagebox.showerror(title="文件不存在", message=f"文件“{msg}“不存在，请重新选择！")


def choice():
    """选择姓名文件"""
    filenames = filedialog.askopenfilenames(title='打开--请选择你要发送的文件', filetypes=[('All Files', '*')], )
    for filename in filenames:
        if filename != '':
            file = os.path.realpath(filename)
            file_on = file.split("\\")[-1]
            tree_files.insert("", END, values=(file_on, file))


def delete_files():
    ss = tree_files.selection()
    for s in ss:
        tree_files.delete(s)


def delete_reco_email():
    ssl = tree_emile.selection()
    print(ssl)
    for sl in ssl:
        print(sl)
        itm = tree_emile.set(sl)
        tree_emile.delete(sl)
        data['email'].remove([itm["备注"], itm["邮箱"]])
        save_data(data, json_data_file)


def delete_send_record():
    ssl = tree_emile_record.selection()
    for sl in ssl:
        itm = tree_emile_record.set(sl)
        tree_emile_record.delete(sl)
        for one_cord in data['record']:
            if one_cord["发送时间"] == itm["发送时间"] and one_cord["收件人"] == itm["收件人"]:
                data['record'].remove(one_cord)
    save_data(data, json_data_file)


if not os.path.exists(cong_ini):
    creat_config()
if eval(read_config(option="pasw_start")):
    def open_window():
        open_win = Tk()
        open_win.geometry(
            "%dx%d+%d+%d" % (
                240, 80, (open_win.winfo_screenwidth() - 240) / 2, (open_win.winfo_screenheight() - 80) / 2))
        open_win.resizable(False, False)  # 禁止调节窗口大小
        open_win.attributes("-toolwindow", 2)  # 去掉窗口最大化最小化按钮，只保留关闭
        open_win.config(background="#8a99a5")
        open_win.title("打开应用")
        icocb(open_win)  # 设置图标
        logf = Frame(open_win, background="#8a99a5")
        logf.pack()
        logf2 = Frame(logf, background="#8a99a5")
        logf2.pack()
        remark_label = Label(logf2, text="请输入密码：", background="#8a99a5", fg="white")
        remark_label.pack(side=LEFT, pady=5)
        remark_enter = Entry(logf2, width=20, show="*", takefocus=True, background="#8a99a5")
        remark_enter.pack(side=LEFT, pady=5)
        logf3 = Frame(logf, background="#8a99a5")
        logf3.pack()

        def que_mm(*args):
            if remark_enter.get() == read_config(option="pasw_word"):
                open_win.destroy()
                return
            else:
                messagebox.showerror(title="", message="密码错误,请重新输入")

        def quite_mm():
            sys.exit()

        que_mm_button = Button(logf3, text="进入应用", command=que_mm, background="#8aa591", fg="white", width=13, bd=0)
        que_mm_button.pack(side=LEFT, padx=10, pady=7, expand=True, fill=X)
        que_mm_button.bind_all('<Return>', que_mm)
        quite_mm_button = Button(logf3, text="退出应用", command=quite_mm, background="#a58a9e", fg="white", width=13,
                                 bd=0)
        quite_mm_button.pack(side=LEFT, padx=10, pady=7, expand=True, fill=X)
        open_win.protocol("WM_DELETE_WINDOW", quite_mm)
        open_win.mainloop()


    open_window()


def display_useer(*args):
    """用于是否明文显示当前登录的邮箱"""
    if eval(str(data['display_user'])):
        kong_1_lable.config(text="🙈")
        sender_email_lable.config(text="*" * 23)
        data['display_user'] = False
    else:
        kong_1_lable.config(text="🙉")
        sender_email_lable.config(text=f"{read_config(option='userName')}")
        data['display_user'] = True
    save_data(data, json_data_file)


def create_shortcut():  # 如无需特别设置图标，则可去掉iconname参数
    try:
        bin_pathc = sys.argv[0]
        yuanshi = os.path.dirname(bin_pathc)
        usernamec = os.path.basename(bin_pathc).split(".exe")[0]
        desk_p = os.path.join(os.path.expanduser("~"), 'Desktop')
        lnkname = f"{desk_p}\\{usernamec}.lnk"
        shortcut = pythoncom.CoCreateInstance(shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER,
                                              shell.IID_IShellLink)
        shortcut.SetPath(bin_pathc)  # 设置文件的路径
        shortcut.SetDescription("云云超帅")  # 设置备注或者描述
        shortcut.SetWorkingDirectory(f"{yuanshi}")  # 设置快捷方式的起始位置, 不然会出现找不到辅助文件的情况
        shortcut.QueryInterface(pythoncom.IID_IPersistFile).Save(lnkname, 0)
        shortcut_bou.forget()
        return True

    except Exception as e:
        print(e.args)
        return False


def log_in_panle():
    log_win = Toplevel()
    screenwidthl = log_win.winfo_screenwidth()  # 获取显示屏宽度
    screenheightl = log_win.winfo_screenheight()  # 获取显示屏高度
    log_win.geometry("%dx%d+%d+%d" % (300, 150, (screenwidthl - 300) / 2, (screenheightl - 150) / 2))
    log_win.resizable(False, False)  # 禁止调节窗口大小
    # log_win.wm_attributes("-topmost", True)
    log_win.title("登录")
    icocb(log_win)  # 设置图标
    logf = ttk.LabelFrame(log_win, text="请输入信息进行登录", labelanchor="n")
    logf.pack()
    logf1 = Frame(logf)
    logf1.pack()
    zhanghao_label = Label(logf1, text="用户邮箱：")
    zhanghao_label.pack(side=LEFT)
    zhanghao_enter = Entry(logf1, width=30)
    zhanghao_enter.pack(side=LEFT)
    zhanghao_enter.insert(0, f"{read_config(option='userName')}")
    Right_Click_Menus(zhanghao_enter, undo=False)
    logf2 = Frame(logf)
    logf2.pack()
    password_label = Label(logf2, text="授 权 码：")
    password_label.pack(side=LEFT)
    password_enter = Entry(logf2, width=30, show="*")
    password_enter.pack(side=LEFT)
    password_enter.insert(0, f"{read_config(option='password')}")
    Right_Click_Menus(password_enter, undo=False)
    logf3 = Frame(logf)
    logf3.pack()
    smtp_server_label = Label(logf3, text="服务地址：")
    smtp_server_label.pack(side=LEFT)
    smtp_server_enter = ttk.Combobox(logf3, values=["smtp.163.com", "smtp.qq.com"], width=15)
    smtp_server_enter.pack(side=LEFT)
    smtp_server_enter.set(f"{read_config(option='smtp_server')}")
    Right_Click_Menus(smtp_server_enter, undo=False)
    duankou_server_label = Label(logf3, text="端口号：")
    duankou_server_label.pack(side=LEFT)
    duankou_server_enter = Entry(logf3, width=4)
    duankou_server_enter.pack(side=LEFT)
    duankou_server_enter.insert(0, f"{read_config(option='port')}")
    Right_Click_Menus(duankou_server_enter, undo=False)
    logf4 = Frame(logf)
    logf4.pack()
    name_label = Label(logf4, text="用户名称：")
    name_label.pack(side=LEFT)
    name_enter = Entry(logf4, width=30)
    name_enter.pack(side=LEFT)
    name_enter.insert(0, f"{read_config(option='senderName')}")
    Right_Click_Menus(name_enter, undo=False)

    def get_info():
        zhanghao = zhanghao_enter.get()
        mima = password_enter.get()
        fuwudizhi = smtp_server_enter.get()
        duankou = duankou_server_enter.get()
        sender = name_enter.get()
        if zhanghao == "" or mima == "" or fuwudizhi == "" or duankou == "":
            messagebox.showerror(title="登录", message="信息输入不全，请补充信息！")
            login_login_button.config(text="登  录", state="normal")
        else:
            if unite_login(usernames=zhanghao,
                           password=mima,
                           smtp_server=fuwudizhi,
                           port=int(duankou),
                           opat=1):
                alter_config(option="userName", values=zhanghao)
                alter_config(option="password", values=mima)
                alter_config(option="smtp_server", values=fuwudizhi)
                alter_config(option="port", values=duankou)
                alter_config(option="senderName", values=sender)
                alter_config(option="have_log", values="True")
                window.update()
                zlkl = eval(read_config(option="have_log"))
                if zlkl:
                    login_button.forget()
                    send_email_lable.pack(side=LEFT)
                    sender_email_lable.pack(side=LEFT)
                    sender_email_lable.config(text=read_config(option="userName"))
                    kong_1_lable.pack(side=LEFT)
                    send_name_lable.pack(side=LEFT)
                    send_name_enter.pack(side=LEFT, expand=YES, fill=X)
                    send_name_enter.insert(0, read_config(option="senderName"))
                    updata_user_button.pack(side=LEFT, padx=5, pady=5)
                    quit_user_button.pack(side=LEFT, padx=10, pady=5)
                    send_button.config(state="normal", text="发送", background="#8aa58b", fg="white")
                log_win.destroy()
            else:
                messagebox.showerror(title="登录", message="登录失败")
                login_login_button.config(text="登  录", state="normal")

    def start_log_in():
        login_login_button.config(text="正在登录", state="disabled")
        threading.Thread(target=get_info).start()

    logf5 = Frame(logf)
    logf5.pack()
    login_login_button = Button(logf5, text="登  录", command=start_log_in, background="#8aa598", fg="white")
    login_login_button.pack(side=LEFT, expand=True, fill=X, padx=10, pady=5)
    login_quit_button = Button(logf5, text="退  出", command=log_win.destroy, background="#a58a97", fg="white")
    login_quit_button.pack(side=RIGHT, expand=True, fill=X, padx=10, pady=5)
    log_win.mainloop()


def log_out_clear(w):
    """退出登录"""
    if clearvalu.get():
        os.rename(cong_ini, f'{time.strftime("%Y%m%d%H%M%S", time.localtime())}{cong_ini}HD')
        creat_config()
    alter_config(option="have_log", values="False")
    login_button.pack(side=LEFT, expand=YES, fill=X)
    send_email_lable.forget()
    sender_email_lable.forget()
    updata_user_button.forget()
    quit_user_button.forget()
    send_name_lable.forget()
    send_name_enter.delete(0, END)
    send_name_enter.forget()
    kong_1_lable.forget()
    send_button.config(state="disabled", text="请先进行登录", background="#8aa58b", fg="white")
    w.destroy()
    """"""


def get_files():
    """获取全部附件"""
    add_files = []
    for item in tree_files.get_children():
        one_file = (tree_files.set(item, 0), tree_files.set(item, 1))
        add_files.append(one_file)
    return add_files


def start_send_email(*args):
    global sending_email_count
    areceiver = ",".join(i for i in receive_email_checkbox.get().split("和"))  # 收件人邮箱
    acc = read_config(option='userName')  # 抄送人邮箱
    asubject = email_title_checkbox.get()  # 邮件主题
    from_addr = read_config(option='userName')  # 发件人地址
    password = read_config(option='password')  # 邮箱密码（授权码）
    sender_name = send_name_enter.get()
    # 邮件设置
    msg = MIMEMultipart()
    msg['Subject'] = asubject
    msg['to'] = areceiver
    msg['Cc'] = acc

    def _format_addr(s):
        addr = parseaddr(s)
        return formataddr(addr)

    # 自定义发件人名称
    msg['From'] = _format_addr(f'{sender_name}<{from_addr}>')
    body = email_title_text.get("1.0", "end")  # 邮件正文
    emtype = email_type_checkbox.get()  # 邮件类型
    if emtype == "默认":
        msg.attach(MIMEText(body, 'plain', 'utf-8'))  # 添加邮件正文
    else:
        try:
            msg.attach(MIMEText(body, f'{emtype}', 'utf-8'))  # 添加邮件正文
        except:
            if messagebox.askyesno(title="有误", message="邮件类型有误,是否使用默认文本类型发送邮件\n《否将取消发送》"):
                msg.attach(MIMEText(body, 'plain', 'utf-8'))  # 防止邮件类型错误
                emtype = "默认"
            else:
                return
    """添加附件"""
    send_add_files = get_files()
    if len(send_add_files) > 0:
        for saf in send_add_files:
            if os.path.exists(saf[1]):
                xlsxpart = MIMEApplication(open(os.path.realpath(saf[1]), 'rb').read())
                xlsxpart.add_header('Content-Disposition', 'attachment', filename=saf[0])
                msg.attach(xlsxpart)
    # 设置邮箱服务器地址以及端口
    try:
        smtp_server = read_config(option='smtp_server')
        port = int(read_config(option='port'))
        server = unite_login(usernames=from_addr, password=password, smtp_server=smtp_server, port=port, opat=2)
        server.sendmail(from_addr, areceiver.split(',') + acc.split(','), msg.as_string())
        # 断开服务器链接
        server.quit()
        threading.Thread(target=save_send_record,
                         args=(areceiver, asubject, emtype, body, send_add_files, from_addr, acc)).start()
        sending_email_count -= 1
        if sending_email_count < 1:
            send_status_label.config(text="空闲", fg="white", background="#627780")
        else:
            send_status_label.config(text=f"正在发送({sending_email_count})", fg="#fa8072", background="#72500e")
        messagebox.showinfo(title="", message="发送成功")

    except AttributeError:
        messagebox.showerror(title="", message="抱歉，邮件可能发送失败了")
        return


def send_email():
    global sending_email_count
    if not eval(read_config(option="have_log")):
        messagebox.showerror(title="", message="请先登录邮箱")
        return
    elif receive_email_checkbox.get() == "" or receive_email_checkbox.get() is None or "@" not in receive_email_checkbox.get():
        messagebox.showerror(title="", message="请填写收件人邮箱")
        return
    if messagebox.askyesno(title="发送", message="是否确认发送邮件？"):
        threading.Thread(target=start_send_email).start()
        sending_email_count += 1
        send_status_label.config(text=f"正在发送({sending_email_count})", fg="#fa8072", background="#72500e")


def close_yes_no():
    """检测关闭时是否还有电子邮件正在发送，如果正在发送则出现关闭提醒"""
    if sending_email_count < 1:
        window.destroy()
    else:
        if messagebox.askyesno(title="运行",
                               message=f"当前还有 {sending_email_count} 封电子邮件正在发送，\n关闭后将发送失败，是否关闭并终止邮件发送？"):
            window.destroy()
            sys.exit()
        else:
            return


def add_alter_panle(title, up_text="邮箱：", down_text="备注：", ema="", remark="", alter: bool = False, conum=""):
    aap_win = Toplevel()
    aap_win.geometry(
        "%dx%d+%d+%d" % (280, 110, (aap_win.winfo_screenwidth() - 280) / 2, (aap_win.winfo_screenheight() - 80) / 2))
    # aap_win.wm_attributes("-topmost", True)
    aap_win.resizable(False, False)  # 禁止调节窗口大小
    aap_win.title(title)
    icocb(aap_win)  # 设置图标
    logf = Frame(aap_win)
    logf.pack()
    logf2 = Frame(logf)
    logf2.pack()
    remark_label = Label(logf2, text=down_text)
    remark_label.pack(side=LEFT, pady=5)
    remark_enter = Entry(logf2, width=30)
    remark_enter.pack(side=LEFT, pady=5)
    remark_enter.insert(0, remark)
    Right_Click_Menus(remark_enter, undo=False)
    logf1 = Frame(logf)
    logf1.pack()
    add_email_label = Label(logf1, text=up_text)
    add_email_label.pack(side=LEFT, pady=5)
    add_email_enter = Entry(logf1, width=30)
    add_email_enter.pack(side=LEFT, pady=5)
    add_email_enter.insert(0, ema)
    Right_Click_Menus(add_email_enter, undo=False)

    def save_info():
        """验证邮箱是否合法"""
        regex = re.compile(r'([A-Za-z0-9]+[.-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+')
        get_ema = add_email_enter.get()
        get_ramark = remark_enter.get()
        for item in tree_emile.get_children():
            if (get_ramark == tree_emile.set(item, 0) or get_ema == tree_emile.set(item, 1)) and not alter:
                messagebox.showerror(title="", message="邮箱用户已存在")
                return
        if re.fullmatch(regex, get_ema) and not alter:
            tree_emile.insert("", END, values=(get_ramark, get_ema))
            data['email'].append((get_ramark, get_ema))
            save_data(data, json_data_file)
        elif re.fullmatch(regex, get_ema) and alter:
            tree_emile.item(conum, values=(get_ramark, get_ema))  # 修改数据
            # tree_emile.delete(conum)
            data['email'].remove([remark, ema])
            data['email'].append((get_ramark, get_ema))
            save_data(data, json_data_file)
            aap_win.destroy()
            messagebox.showinfo(title="", message="修改成功")
        else:
            messagebox.showerror(title="", message="邮箱格式不正确!")

    logf5 = Frame(logf)
    logf5.pack()
    login_login_button = Button(logf5, text="保    存", command=save_info, background="#8aa598", fg="white")
    login_login_button.pack(side=LEFT, expand=True, fill=X, padx=10, pady=5)
    login_quit_button = Button(logf5, text="退    出", command=aap_win.destroy, background="#a58a97", fg="white")
    login_quit_button.pack(side=RIGHT, expand=True, fill=X, padx=10, pady=5)

    aap_win.mainloop()


def log_out():
    global clearvalu
    quit_win = Toplevel()
    quit_win.attributes("-toolwindow", 2)  # 去掉窗口最大化最小化按钮，只保留关闭
    screenwidthl = quit_win.winfo_screenwidth()  # 获取显示屏宽度
    screenheightl = quit_win.winfo_screenheight()  # 获取显示屏高度
    quit_win.geometry("%dx%d+%d+%d" % (260, 150, (screenwidthl - 300) / 2, (screenheightl - 150) / 2))
    quit_win.resizable(False, False)  # 禁止调节窗口大小
    quit_win.wm_attributes("-topmost", True)
    icocb(quit_win)  # 设置图标
    quit_win.title("退出登录")
    fq1 = Frame(quit_win)
    fq1.pack()
    tips_lable = Label(fq1, text="是否退出登录?", font=("微软雅黑", 15, "bold"), fg="red")
    tips_lable.pack(pady=20, padx=40, expand=True, fill=X)
    clearquit_checkbox = ttk.Checkbutton(fq1, text="退出登录并清除账号密码", offvalue=False, onvalue=True,
                                         variable=clearvalu,
                                         takefocus=False)

    clearquit_checkbox.pack()
    fq2 = Frame(quit_win)
    fq2.pack()
    yesquit_button = Button(fq2, text="退  出", width=10, background="#844947", fg="white",
                            command=lambda: log_out_clear(w=quit_win))
    yesquit_button.pack(side=LEFT, pady=15, padx=20, expand=True, fill=X)
    noquit_button = Button(fq2, text="取  消", width=10, command=quit_win.destroy, background="#478284", fg="white")
    noquit_button.pack(side=LEFT, pady=15, padx=20, expand=True, fill=X)
    quit_win.mainloop()


def colse_pass_word():
    if colse_start_enter.get() == read_config(option="pasw_word"):
        pass_start_enter.forget()
        confirm_button.forget()
        alter_button.forget()
        colse_start_enter.delete(0, END)
        colse_start_enter.forget()
        colse_pasw_button.forget()
        alter_config(option="pasw_start", values="False")
        pastavalu.set(False)
    else:
        messagebox.showerror("", message="密码错误")
        colse_start_enter.delete(0, END)
        pastavalu.set(True)
        colse_start_enter.forget()
        colse_pasw_button.forget()


def yn_set_pasrt():
    global pastavalu
    if pastavalu.get():
        pass_start_enter.pack(expand=True, fill=X, pady=7)
        confirm_button.pack(side=LEFT, expand=True, fill=X, pady=7, padx=7)
        alter_button.pack(side=LEFT, expand=True, fill=X, pady=7, padx=7)
        alter_config(option="pasw_start", values="True")
        pastavalu.set(True)
    else:
        colse_start_enter.pack(expand=True, fill=X, pady=5, padx=5)
        colse_pasw_button.pack(expand=True, fill=X, pady=5, padx=5)
        pastavalu.set(True)


def alter_ema():
    """修改通讯录邮箱"""
    ss = tree_emile.selection()
    try:
        s = ss[0]
        itm = tree_emile.set(s)
        add_alter_panle(title="修改", ema=itm["邮箱"], remark=itm["备注"], alter=True, conum=s)
    except IndexError:
        return


def add_receive_email(*args):
    """双击添加收件人至收件人栏"""
    all_receivor = receive_email_checkbox.get().split("和")
    ss = tree_emile.selection()
    s = ss[0]
    itm = tree_emile.set(s)
    if itm["邮箱"] in all_receivor:
        messagebox.showerror(title="", message="该用户已添加，请勿重复添加")
    else:
        receive_email_checkbox.insert(END, f"{itm['邮箱']}和")


def querenmiama():
    if confie_mima_enter.get() == read_config(option="pasw_word"):
        confirm_button.config(state="normal")
        pass_start_enter.config(state="normal")
        querenmi_button.forget()
        querebmima_lable.forget()
        confie_mima_enter.forget()
        alter_button.config(state="disabled")
        confie_mima_enter.delete(0, END)
    else:
        messagebox.showerror(title="", message="密码错误，请重新输入")


def open_confie():
    pass_start_enter.forget()
    querenmi_button.forget()
    querebmima_lable.forget()
    confie_mima_enter.forget()
    confirm_button.forget()
    alter_button.forget()
    pass_start_enter.pack(expand=True, fill=X, pady=7)
    querenmi_button.pack(expand=True, fill=X, pady=7)
    querebmima_lable.pack(expand=True, fill=X, pady=7, side=LEFT)
    confie_mima_enter.pack(expand=True, fill=X, pady=7, side=LEFT)
    confirm_button.pack(side=LEFT, expand=True, fill=X, pady=7, padx=7)
    alter_button.pack(side=LEFT, expand=True, fill=X, pady=7, padx=7)


def save_pass_word():
    pass_word = pass_start_enter.get()
    if messagebox.askyesno(title="", message="是否确认修改？"):
        alter_config(option="pasw_word", values=pass_word)
        alter_button.config(state="normal")
        confirm_button.config(state="disabled")
        pass_start_enter.config(state="readonly")
    else:
        return


sending_email_count = 0
threading.Thread(target=open_log_in).start()  # 进行开机登录
data = read_data(data, json_data_file)
data['program_path'] = sys.argv[0]  # 更新程序路径
window = Tk()
clearvalu = BooleanVar()  # 是否清除登录信息
clearvalu.set(False)
pastavalu = BooleanVar()  # 是否清除登录信息
pastavalu.set(eval(read_config(option="pasw_start")))
# window.wm_attributes("-alpha", 0)  # 透明度(0.0~1.0)
window.title(f"云邮瀛瀛 YYMV {'  ' * 50}云里●悟理")
screenwidth = window.winfo_screenwidth()  # 获取显示屏宽度
screenheight = window.winfo_screenheight()  # 获取显示屏高度
window.geometry("%dx%d+%d+%d" % (825, 538, (screenwidth - 825) / 2, (screenheight - 538) / 2))
icocb(window)  # 设置图标
fll = Frame(window)
fll.pack(side=LEFT, expand=True, fill=BOTH)
f1n = ttk.Notebook(fll, takefocus=False)
f1l = ttk.LabelFrame(text="通 讯 录", labelanchor="n")
f2l = ttk.LabelFrame(text="已发邮件记录", labelanchor="n")
f3l = ttk.LabelFrame(text="隐私设置", labelanchor="n")
# f4l = ttk.LabelFrame(text="收件箱", labelanchor="n")
f1n.add(f2l, text="已发邮件记录")
f1n.add(f1l, text="通讯录")
# f1n.add(f4l, text="收件箱")
f1n.add(f3l, text="隐私设置")
f1n.pack(side=LEFT, expand=True, fill=BOTH)
f1l1 = Frame(f1l)
f1l1.pack(fill=X)
add_email_button = Button(f1l1, text="添 加", background="#8a97a5", fg="white",
                          command=lambda: add_alter_panle(title="新增"))
add_email_button.pack(side=LEFT, expand=True, fill=X, padx=7, pady=5)
alter_email_button = Button(f1l1, text="修 改", background="#a5998a", fg="white", command=alter_ema)
alter_email_button.pack(side=LEFT, expand=True, fill=X, padx=7, pady=5)
delete_email_button = Button(f1l1, text="删 除", background="#a58a90", fg="white", command=delete_reco_email)
delete_email_button.pack(side=LEFT, expand=True, fill=X, padx=7, pady=5)
shortcut_bou = Button(f1l1, text="快捷方式", command=create_shortcut, background="#90a58a", fg="white")
shortcut_bou.pack(side=LEFT, padx=7, pady=5)
"""判断是否存在快捷方式，如果存在就不显示创建快捷方式的按钮"""
bin_path = sys.argv[0]
username = os.path.basename(bin_path).split(".exe")[0]
zxp = os.path.join(os.path.expanduser("~"), 'Desktop')
shortcut = f"{zxp}\\{username}.lnk"
if os.path.exists(shortcut):
    shortcut_bou.forget()
"""----------------------------------------------------"""
yscroll_email = Scrollbar(f1l, orient=VERTICAL)
tree_emile = ttk.Treeview(f1l, show='headings', yscrollcommand=yscroll_email.set)  # 表格
yscroll_email.config(command=tree_emile.yview)
yscroll_email.pack(side=RIGHT, fill=Y)
tree_emile["columns"] = ("备注", "邮箱")
tree_emile.column("备注", width=50)  # 表示列,不显示
tree_emile.column("邮箱", width=160)
tree_emile.heading("备注", text="备注")  # 显示表头
tree_emile.heading("邮箱", text="邮箱")
if len(data['email']) > 0:
    for emile_one in data['email']:
        tree_emile.insert("", 0, values=emile_one)  # 插入数据，
tree_emile.pack(expand=YES, fill=BOTH)
tree_emile.bind('<Double-Button-1>', add_receive_email)
f2l1 = Frame(f2l)
f2l1.pack(fill=BOTH)
delete_email_button = Button(f2l1, text="删 除 选 中 记 录", background="#a58a90", fg="white",
                             command=delete_send_record,
                             takefocus=False, width=20)
delete_email_button.pack(side=LEFT, expand=True, fill=X, pady=5)
yscroll_email_recode = Scrollbar(f2l, orient=VERTICAL)
tree_emile_record = ttk.Treeview(f2l, show='headings', yscrollcommand=yscroll_email_recode.set)  # 表格
yscroll_email_recode.config(command=tree_emile_record.yview)
yscroll_email_recode.pack(side=RIGHT, fill=Y)
tree_emile_record["columns"] = ("发送时间", "收件人")
tree_emile_record.column("发送时间", width=130)  # 表示列,不显示
tree_emile_record.column("收件人", width=80)
tree_emile_record.heading("发送时间", text="发送时间")  # 显示表头
tree_emile_record.heading("收件人", text="收件人")
if len(data['record']) > 0:
    for emile_one in data['record']:
        tree_emile_record.insert("", 0, values=(emile_one["发送时间"], emile_one["收件人"]))  # 插入数据，
tree_emile_record.pack(expand=YES, fill=BOTH)
tree_emile_record.bind('<Double-Button-1>', show_detailed_email)



# f4l1 = Frame(f4l)
# f4l1.pack(fill=BOTH)
# refresh_email_button = Button(f4l1, text="刷新信箱", background="#a58a90", fg="white",
#                               command=delete_send_record,
#                               takefocus=False, width=20)
# refresh_email_button.pack(side=LEFT, expand=True, fill=X, pady=5)
# delrece_email_button = Button(f4l1, text="删除邮件", background="#a58a90", fg="white",
#                               command=delete_send_record,
#                               takefocus=False, width=20)
# delrece_email_button.pack(side=LEFT, expand=True, fill=X, pady=5)
# yscroll_recemail = Scrollbar(f4l, orient=VERTICAL)
# tree_emile_rece = ttk.Treeview(f4l, show='headings', yscrollcommand=yscroll_recemail.set)  # 表格
# yscroll_recemail.config(command=tree_emile_record.yview)
# yscroll_recemail.pack(side=RIGHT, fill=Y)
# tree_emile_rece["columns"] = ("收件时间", "邮件主题")
# tree_emile_rece.column("收件时间", width=130)  # 表示列,不显示
# tree_emile_rece.column("邮件主题", width=80)
# tree_emile_rece.heading("收件时间", text="收件时间")  # 显示表头
# tree_emile_rece.heading("邮件主题", text="邮件主题")
# # if len(data['record']) > 0:
# #     for emile_one in data['record']:
# #         tree_emile_rece.insert("", 0, values=(emile_one["发送时间"], emile_one["收件人"]))  # 插入数据，
# tree_emile_rece.pack(expand=YES, fill=BOTH)
# tree_emile_rece.bind('<Double-Button-1>', show_detailed_email)



f3l1 = Frame(f3l)
f3l1.pack()
set_passw_checkbutton = ttk.Checkbutton(f3l1, takefocus=False, text="是否设置启动密码", offvalue=False, onvalue=True,
                                        variable=pastavalu, command=yn_set_pasrt)
set_passw_checkbutton.pack()
f3l2 = Frame(f3l)
f3l2.pack(expand=YES, fill=BOTH)
pass_start_enter = Entry(f3l2, width=20, show="**", fg="blue")
pass_start_enter.pack(expand=YES, fill=X, pady=7, side=TOP)
pass_start_enter.insert(0, read_config(option="pasw_word"))
pass_start_enter.config(state="readonly", )
querenmi_button = Button(f3l2, takefocus=False, text="确  认", command=querenmiama, background="#8aa592", fg="white")
f3l21 = Frame(f3l2)
f3l21.pack(side=TOP, expand=YES, fill=BOTH)
querebmima_lable = Label(f3l21, text="请输入原密码:")
confie_mima_enter = Entry(f3l21, width=20, show="*", fg="red")

f3l3 = Frame(f3l)
f3l3.pack()
confirm_button = Button(f3l2, takefocus=False, text="保存密码", state="disabled", command=save_pass_word,
                        background="#8aa59f", fg="white")
confirm_button.pack(side=LEFT, expand=True, fill=X, pady=7, padx=7)
alter_button = Button(f3l2, takefocus=False, text="修改密码", command=open_confie, background="#a58a90", fg="white")
alter_button.pack(side=LEFT, expand=True, fill=X, pady=7, padx=7)
f3l4 = Frame(f3l)
f3l4.pack()

colse_start_enter = Entry(f3l4, width=20, show="*", fg="blue")
colse_pasw_button = Button(f3l4, takefocus=False, text="确 认 取 消", command=colse_pass_word, background="#8aa592",
                           fg="white")

if not pastavalu.get():
    pass_start_enter.forget()
    confirm_button.forget()
    alter_button.forget()

f1r = Frame(window)
f1r.pack(side=RIGHT, expand=YES, fill=BOTH)

f1r1 = Frame(f1r)
f1r1.pack(expand=YES, fill=BOTH, padx=5, pady=5)
send_email_lable = Label(f1r1, text="发件人邮箱：", width=10, anchor="w")
send_email_lable.pack(side=LEFT, expand=YES, fill=BOTH)
login_button = Button(f1r1, text="登录账户", command=log_in_panle)
login_button.pack(side=LEFT, expand=YES, fill=BOTH)
sender_email_lable = Label(f1r1, text=f"{read_config(option='userName')}", font=("微软雅黑", 10, "underline"),
                           fg="blue")
sender_email_lable.pack(side=LEFT, expand=YES, fill=BOTH)
if eval(str(data['display_user'])):
    kong1text = "🙉"
    sender_email_lable.config(text=f"{read_config(option='userName')}")
else:
    kong1text = "🙈"
    sender_email_lable.config(text="*" * 23)
kong_1_lable = Label(f1r1, text=kong1text, width=5)
kong_1_lable.pack(side=LEFT, expand=YES, fill=BOTH)
kong_1_lable.bind("<Button-1>", display_useer)
send_name_lable = Label(f1r1, text="发件人姓名：", anchor="w")
send_name_lable.pack(side=LEFT, expand=YES, fill=BOTH)
send_name_enter = Entry(f1r1, width=12)
send_name_enter.pack(side=LEFT, expand=YES, fill=BOTH)
Right_Click_Menus(send_name_enter, undo=False)
send_name_enter.insert(0, f"{read_config(option='senderName')}")
updata_user_button = Button(f1r1, text="更换用户", command=log_in_panle, background="#ab806c", fg="white")
updata_user_button.pack(side=LEFT, padx=5, pady=5, expand=YES, fill=X)
quit_user_button = Button(f1r1, text="退出账号", command=log_out, background="#955250", fg="white")
quit_user_button.pack(side=LEFT, padx=10, pady=5, expand=YES, fill=X)
"""判断是否登录"""
zlk = eval(read_config(option="have_log"))
if zlk:
    login_button.forget()
    send_email_lable.pack(side=LEFT)
    sender_email_lable.pack(side=LEFT)
    updata_user_button.pack(side=LEFT, padx=5, pady=5)
    quit_user_button.pack(side=LEFT, padx=10, pady=5)
    send_name_lable.pack(side=LEFT)
    send_name_enter.pack(side=LEFT, expand=YES, fill=X)
    kong_1_lable.pack(side=LEFT)
else:
    login_button.pack(side=LEFT, expand=YES, fill=X)
    send_email_lable.forget()
    sender_email_lable.forget()
    updata_user_button.forget()
    quit_user_button.forget()
    send_name_lable.forget()
    send_name_enter.forget()
    kong_1_lable.forget()
""""""
f1r2 = Frame(f1r)
f1r2.pack(expand=YES, fill=X, padx=5, pady=5)
receive_email_lable = Label(f1r2, text="收件人邮箱：", width=10, anchor="w")
receive_email_lable.pack(side=LEFT)
receive_email_checkbox = Entry(f1r2, )
receive_email_checkbox.pack(side=LEFT, expand=YES, fill=X)
Right_Click_Menus(receive_email_checkbox, undo=False)
f1r3 = Frame(f1r)
f1r3.pack(expand=YES, fill=X, padx=5, pady=5)
email_title_lable = Label(f1r3, text="邮 件 标 题：", width=10, anchor="w")
email_title_lable.pack(side=LEFT)
email_title_checkbox = Entry(f1r3, )
email_title_checkbox.pack(side=LEFT, expand=YES, fill=X)
Right_Click_Menus(email_title_checkbox, undo=False)
email_type_lable = Label(f1r3, text="邮件类型：", anchor="w")
email_type_lable.pack(side=LEFT)
email_type_checkbox = ttk.Combobox(f1r3, values=["默认", "html"], width=6)
email_type_checkbox.pack(side=LEFT, )
email_type_checkbox.set("默认")
Right_Click_Menus(email_type_checkbox, undo=False)
f1r4 = ttk.LabelFrame(f1r, text="邮件正文", labelanchor="n")
f1r4.pack(expand=YES, fill=BOTH, padx=5, pady=5)
email_title_text = Text(f1r4, undo=True, height=10)
email_title_text.pack(side=LEFT, expand=YES, fill=BOTH)
Right_Click_Menus(email_title_text, undo=True)
f1r5 = ttk.LabelFrame(f1r, text="附件（可拖拽导入)", labelanchor="n")
f1r5.pack(expand=YES, fill=X, padx=5, pady=5)
windnd.hook_dropfiles(f1r5, func=dragged_files)  # 拖拽导入文件
f1r51 = Frame(f1r5)
f1r51.pack(expand=YES, fill=X, padx=5, pady=5)
chico_file_button = Button(f1r51, text="添加附件", command=choice, background="#8aa59f", fg="white")
chico_file_button.pack(expand=YES, fill=X, padx=5, pady=5, side=LEFT)
delete_file_button = Button(f1r51, text="删除附件", background="#a58a90", fg="white", command=delete_files)
delete_file_button.pack(expand=YES, fill=X, padx=5, pady=5, side=RIGHT)
f1r52 = Frame(f1r5)
f1r52.pack(expand=YES, fill=X, padx=5, pady=5)
yscroll_files = ttk.Scrollbar(f1r52, orient=VERTICAL)
tree_files = ttk.Treeview(f1r52, show='headings', height=5, yscrollcommand=yscroll_files.set)  # 表格
yscroll_files.config(command=tree_files.yview)
yscroll_files.pack(side=RIGHT, fill=Y)
tree_files["columns"] = ("附件", "附件地址")
tree_files.column("附件", width=100)  # 表示列,不显示
tree_files.column("附件地址", width=300)
tree_files.heading("附件", text="附件")  # 显示表头
tree_files.heading("附件地址", text="附件地址")
tree_files.pack(expand=YES, fill=BOTH)

f1r6 = Frame(f1r)
f1r6.pack(expand=YES, fill=X, padx=5, pady=5)
send_button = Button(f1r6, text="请先进行登录", background="#8aa58b", fg="white", command=send_email, state="disabled")
send_button.pack(expand=YES, fill=BOTH, side=LEFT, padx=10)
send_status_label = Label(f1r6, takefocus=False, text="空闲", fg="white", background="#627780", width=10)
send_status_label.pack(side=LEFT, fill=Y, padx=10)
window.protocol("WM_DELETE_WINDOW", close_yes_no)
window.mainloop()
