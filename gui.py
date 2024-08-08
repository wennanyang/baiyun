import threading
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from make_result import main
from tkinter import messagebox
import os
class GUI():
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.fang_var = tk.StringVar()
        self.fu_var = tk.StringVar()
        self.validate_var = tk.StringVar()
        self.progress_lable_var = tk.StringVar()
        self.root.title("放线验收数据程序")
        self.root.configure(bg="skyblue")
        self.root.minsize(200, 200)  # width, height
        self.root.maxsize(500, 500)
        self.root.geometry("600x300+250+250")
        fang_frame = tk.Frame(self.root, bg="#6FAFE7")
        # 设置第一行放线路径的label, entry, button
        fang_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        fang_label = tk.Label(fang_frame, text="放线路径", bg="#6FAFE7")
        fang_label.grid(row=0, column=0)
        self.fang_entry = tk.Entry(fang_frame, bd=3, width=50, textvariable=self.fang_var)
        self.fang_entry.grid(row=0, column=1)
        choose_fang_dir = tk.Button(fang_frame, text="选择目录", 
                                    command=lambda: self.select_directory(self.fang_var))
        choose_fang_dir.grid(row=0, column=2)
        # 设置第二行验收路径的label, entry, button
        fu_frame = tk.Frame(self.root, bg="#6FAFE7")
        fu_frame.grid(row=1, column=0, padx=10, pady=10, sticky='nsew')
        fu_lable = tk.Label(fu_frame, text="验收路径", bg="#6FAFE7")
        fu_lable.grid(row=0, column=0)
        self.fu_entry = tk.Entry(fu_frame, bd=3, width=50,textvariable=self.fu_var)
        self.fu_entry.grid(row=0, column=1)
        choose_fu_dir = tk.Button(fu_frame, text="选择目录", 
                                  command=lambda: self.select_directory(self.fu_var))
        choose_fu_dir.grid(row=0, column=2)
        # 设置第三行的验证excel路径
        validate_frame = tk.Frame(self.root, bg="#6FAFE7")
        validate_frame.grid(row=2, column=0, padx=10, pady=10, sticky='nsew')
        validate_lable = tk.Label(validate_frame, text="清单路径", bg="#6FAFE7")
        validate_lable.grid(row=0, column=0)
        self.validate_entry = tk.Entry(validate_frame, bd=3, width=50, textvariable=self.validate_var)
        self.validate_entry.grid(row=0, column=1)
        choose_validate = tk.Button(validate_frame, text="选择文件", 
                                    command=lambda: self.select_file(self.validate_var))
        choose_validate.grid(row=0, column=2)
        # 设置第四行执行
        self.description_var = tk.StringVar()
        excute_frame = tk.Frame(self.root, bg="#6FAFE7")
        excute_frame.grid(row=3, column=0, padx=10, pady=10, sticky='nsew')
        self.turn_on = tk.Button(excute_frame, text="开始执行", command=self.excute)
        self.turn_on.grid(row=0, column=0)
        self.description_label = tk.Label(excute_frame, bg="#6FAFE7", textvariable=self.description_var)
        self.description_label.grid(row=0, column=1)

        ## 第5行的进度条
        progress_fram = tk.Frame(self.root, bg="#6FAFE7")
        progress_fram.grid(row=4, column=0, padx=10, pady=10, sticky='nsew')
        self.progress_var = tk.IntVar()
        self.progressbar = ttk.Progressbar(progress_fram, orient="horizontal", length=400, mode="determinate", 
                                              variable=self.progress_var)
        self.progressbar["value"] = 0
        self.progressbar["maximum"] = 100
        self.pogress_label = tk.Label(progress_fram, text="进度", bg="#6FAFE7", textvariable=self.progress_lable_var)
        self.progressbar.grid(row=0, column=0)
        self.pogress_label.grid(row=0, column=1)
        self.thread = threading.Thread(target=self.long_running_task,args=(), daemon=True)
    def mainloop(self):
        self.root.mainloop()
    def select_directory(self, stringvar):
        directory = filedialog.askdirectory()
        if directory:
            stringvar.set(os.path.normpath(directory))
    def select_file(self, stringvar):
        # 打开文件选择对话框并获取文件路径
        file_path = filedialog.askopenfilename(
            title="选择一个xls文件",  # 对话框标题
            filetypes=[("excel文件", "*.xls"), ("excel文件", "*.xlsx") ]  # 过滤文件类型
        )
        if file_path:
            stringvar.set(os.path.normpath(file_path))
    def excute(self):
        self.thread.start()
    def update_progress(self, value, description):
        self.progress_var.set(value)
        self.progress_lable_var.set(f"{value : .2f}%")
        self.description_var.set(description)
        self.root.update_idletasks() 
    def long_running_task(self):
        main(fang_dir=self.fang_entry.get(), 
             fu_dir=self.fu_entry.get(), 
             validate_xls=self.validate_entry.get(),
             progress_callback=self.update_progress)
        self.on_success()
    def on_success(self):
        self.description_var.set("")
        self.progress_var.set(0)
        messagebox.showinfo("成功", "提取属性成功！")     
if __name__ == '__main__':
    gui = GUI()
    gui.mainloop()