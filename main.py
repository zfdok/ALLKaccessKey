import tkinter
from pandas_tools import readexcel 
import win32api
from os.path import dirname, abspath


class App:
    def __init__(self, master):
        frame = tkinter.Frame(master)
        frame.pack()
        master.title("安冷冷库设备鉴权码计算器")
        self.btn1 = tkinter.Button(
            frame, text="编辑设备清单", width=20, height=7, background="lightblue")
        self.btn1.grid(row=0, column=0)
        self.btn1.bind('<Button-1>', self.leftclick)
        self.btn2 = tkinter.Button(
            frame, text="生成设备鉴权码", width=20, height=7, background="lightgreen")
        self.btn2.grid(row=0, column=1)
        self.btn2.bind('<Button-1>', self.btn2_leftclick)
        pass

    def leftclick(self, event):
        readexcel.clear_AccessKey()
        project_path = dirname(abspath(__file__))
        print(project_path)
        self.handle = win32api.ShellExecute(0, 'open', project_path +
                                            '/excel/coding.xlsx', '', '', 1)

    def btn2_leftclick(self, event):
        project_path = dirname(abspath(__file__))
        print(project_path)
        readexcel.get_AccessKey()
        self.handle = win32api.ShellExecute(0, 'open', project_path +
                                            '/excel/coding.xlsx', '', '', 1)


if __name__ == "__main__":
    windows = tkinter.Tk()
    windows.geometry('320x140+1100+50')
    app = App(windows)

    tkinter.mainloop()