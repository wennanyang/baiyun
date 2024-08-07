import threading
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from make_result import main


class GUI():
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.fang_var = tk.StringVar()
        self.fu_var = tk.StringVar()
        self.validate_var = tk.StringVar()
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
        choose_fang_dir = tk.Button(fang_frame, text="选择文件", 
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
        # 设置第四行验收路径的label, entry, button
        self.turn_on = tk.Button(self.root, text="开始执行", command=self.excute)
        self.turn_on.grid(row=3)

        ## 进度条
        self.progress_var = tk.IntVar()
        self.progressbar = ttk.Progressbar(self.root, orient="horizontal", length=300, mode="determinate", 
                                              variable=self.progress_var)
        self.progressbar.grid(row=4)
    def mainloop(self):
        self.root.mainloop()
    def select_directory(self, stringvar):
        directory = filedialog.askdirectory()
        if directory:
            stringvar.set(directory)
    def select_file(self, stringvar):
        # 打开文件选择对话框并获取文件路径
        file_path = filedialog.askopenfilename(
            title="选择一个xls文件",  # 对话框标题
            filetypes=[("excel文件", "*.xls"), ("excel文件", "*.xlsx") ]  # 过滤文件类型
        )
        if file_path:
            stringvar.set(file_path)
    def excute(self):
        print(self.fang_entry.get())
        print(self.fu_entry.get())
        print(self.validate_entry.get())
        threading.Thread(target=self.long_running_task,args=(), daemon=True).start()
    def update_progress(self, value):
        self.progress_var.set(value)
        self.root.update_idletasks() 
    def long_running_task(self):
        main(fang_dir=self.fang_entry.get(), 
             fu_dir=self.fu_entry.get(), 
             validate_xls=self.validate_entry.get(),
             progress_callback=self.update_progress)
        
if __name__ == '__main__':
    gui = GUI()
    gui.mainloop()