import tkinter as tk
import tkinter.scrolledtext
import tkinter.filedialog
from picter import icocb


class Right_Click_Menus:
    """创建一个弹出菜单"""

    def __init__(self, text, undo=True):
        menu = tk.Menu(text, tearoff=False)
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


def show_information(information="", title="邮件详情"):
    """显示信息"""
    global information_window
    global information_scrolledtext

    def save_txt(information=information, title=title):
        filename = tkinter.filedialog.asksaveasfilename(
            title='请选择你要保存的地方', filetypes=[('TXT', '*.txt'), ('All Files', '*')],
            initialfile='%s' % title,
            defaultextension='txt',  # 默认文件的扩展名
        )  # 返回文件名--另存为
        if filename == '':
            return False
        else:
            with open(filename, 'w') as f:
                f.write(information)
                # f.close()
            return True

    try:
        information_window.deiconify()
        information_window.title(title)
        information_scrolledtext.delete(0.0, tk.END)
        information_scrolledtext.insert(tk.END, information)
        print(112)
    except:
        information_window = tk.Tk()
        information_window.title(title)
        information_window["bg"] = "#ddcdc6"
        screenwidth = information_window.winfo_screenwidth()  # 获取显示屏宽度
        screenheight = information_window.winfo_screenheight()  # 获取显示屏高度
        width = 530
        height = 600
        information_window.geometry(
            "%dx%d+%d+%d" % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2))
        icocb(information_window)

        information_scrolledtext = tkinter.scrolledtext.ScrolledText(
            information_window,
            width=70,
            height=23,
            undo=True,
            background="#f0e9e6",
            font=('微软雅黑', 10)
        )  # 滚动文本框（宽，高（这里的高应该是以行数为单位），字体样式）
        information_scrolledtext.pack(expand=tk.YES, fill=tk.BOTH, padx=5, pady=5)
        Right_Click_Menus(information_scrolledtext, undo=False)
        information_scrolledtext.insert(tk.INSERT, information)

        bottom_frame = tk.Frame(information_window, background="#ddcdc6")
        bottom_frame.pack()

        save_button = tk.Button(
            bottom_frame,
            text="另存邮件为文本文档(*.txt)",
            background="#86b893",
            fg="white",
            command=lambda: save_txt(information=information_scrolledtext.get('1.0', tk.END).rstrip()),
            width=20)
        save_button.pack(side=tk.LEFT, padx=18, pady=5)

        def copy_to_clipboard():
            """Copy current contents of text_entry to clipboard."""
            information_window.clipboard_clear()  # Optional.
            information_window.clipboard_append(information_scrolledtext.get('1.0', tk.END).rstrip())
        copy_button = tk.Button(
            bottom_frame,
            text="复制邮件内容到剪贴板",
            command=copy_to_clipboard,
            background="#86aab8",
            width=20,
            fg="white"
        )
        copy_button.pack(side=tk.LEFT, padx=18, pady=5)

        close_button = tk.Button(
            bottom_frame,
            text="关 闭 窗 口",
            background="#b89386",
            fg="white",
            width=20,
            command=information_window.destroy)
        close_button.pack(side=tk.LEFT, padx=18, pady=5)





        information_window.mainloop()


def main():
    show_information(information="我是用来显示信息的！", title="邮件详情")


if __name__ == "__main__":
    main()
