import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from utils import utils


def choose_file():
    old_file_path_length = len(entry1.get())
    entry1.delete(0, old_file_path_length)
    file_path = tk.filedialog.askopenfilename()
    # askopenfilename 1次上传1个；askopenfilenames 1次上传多个
    entry1.insert(0, file_path)


def count_data():
    file_path = entry1.get()
    count_time = entry2.get()
    if file_path == "":
        messagebox.showinfo(title='提示', message="错误：请先选择需要统计的 excel 文件")
        raise Exception("请先选择需要统计的 excel 文件")
    try:
        import driver.excel as drive
        data = drive.Data(file_path)
    except Exception:
        try:
            import driver.csv as drive
            data = drive.Data(file_path)
        except Exception:
            messagebox.showinfo(title='提示', message="错误：输入的 excel 文件不存在，或文件格式不为 xls、xlsx 或 csv")
            raise Exception("错误：输入的 excel 文件不存在，或文件格式不为 xls、xlsx 或 csv")
    medicine_info = data.medicine_info
    medicine_name = data.medicine_name
    '''
    {
        "药房": {
            "科室":{
                "数量":"",
                "医生": {
                    "a":""
                },
                "病人": {
                    "a": ""
                }
            }
        }
    }
    '''
    room_dict = data.room_dict
    try:
        utils.write_to_excel(room_dict, medicine_name, medicine_info, count_time)
    except Exception as e:
        messagebox.showinfo(title='提示', message="错误：%s" % e)
        raise Exception("错误：%s" % e)


if __name__ == '__main__':
    root = tk.Tk()
    version = 'v2.0.1.4'
    root.title("开单数量统计 %s" % version)
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    root_width = 300
    root_height = 150
    x = (screen_width - root_width) / 2
    y = (screen_height - root_height) / 2
    root.geometry("%dx%d+%d+%d" % (root_width, root_height, x, y))
    root.resizable(0, 0)
    # 定义选择提示 label
    label1 = tk.Label(root, text="请选择需要统计的 excel 文件")
    label1.place(x=30, y=10)
    # 定义输入框 1
    entry1 = tk.Entry(root, width='28', bd=5)
    entry1.place(x=30, y=30)
    # 定义选择按钮
    btn1 = tk.Button(root, text='...', width=3, command=choose_file)
    btn1.place(x=250, y=30)
    # 定义提示框 1
    label2 = tk.Label(root, text="请输入本次统计时间（可选项）")
    label2.place(x=30, y=60)
    # 定义输入框 2
    entry2 = tk.Entry(root, width='28', bd=5)
    entry2.place(x=30, y=80)
    # 定义统计按钮
    btn2 = tk.Button(root, text='统计数据', command=count_data)
    btn2.place(relx=0.4, y=110)
    # 启动窗口
    root.mainloop()
