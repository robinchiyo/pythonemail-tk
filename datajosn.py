import json
import time
import sys

# 用于文件操作
import os.path
# 图形界面
from tkinter import *
# 对话框
from tkinter import messagebox


def save_data(data, json_data_file):
    """保存数据到文件"""
    # 更新校验码，弃用
    # data['check_code'] = check_code(data)
    # 写入 JSON 数据
    with open(json_data_file, 'w') as f:
        json.dump(data, f, indent=4)
    return data


def read_data(data, json_data_file):
    """读数据"""
    # 判断数据文件是否存在
    if os.path.isfile(json_data_file):
        with open(json_data_file, 'r') as f:
            try:  # 尝试把json转换成Python的数据类型
                new_data = json.load(f)
                # print(new_data)
                return new_data
            except:
                top = Tk()
                top.withdraw()  # 为了防止弹出对话框时出现白框框--隐藏窗口

                messagebox.showerror(
                    title='数据损坏，无法读取！',
                    message='数据损坏，无法读取！\n已恢复默认数据')
                top.destroy()  # 销毁临时窗口
                del top
                save_data(data, json_data_file)
            return data
    else:
        # print("文件")
        save_data(data, json_data_file)
        return data


def change_data(new_data, old_data):
    """改变数据"""
    # 浅拷贝
    data = new_data.copy()
    # 获取旧数据的所有键
    old_keys = list(old_data.keys())
    # 如果没有则加上
    for key in old_keys:
        if key not in data:
            data[key] = old_data[key]

    return data


# json数据文件名称/地址
# json_data_file = 'data.json'
# # 初始化数据
# data = {
#     'email': [],  # 程序名
#     'record': [],
#     'display_user': True,  # 是否显示用户名
#     'program_path': sys.argv[0],  # 程序路径
#     'author': '程甲第',  # 制作人
#     'time': time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),  # 时间
# }
