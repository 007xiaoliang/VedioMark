import datetime
import os
import random
import threading
import time
import tkinter
import math
import xlrd
import tkinter.messagebox
from tkinter import *
from tkinter import filedialog
from tkinter.font import Font
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import *
import subprocess
import cv2
from xlrd import xldate_as_tuple


class WaterMark():
    def __init__(self, ifadd, file_path, logo_path, new_file, watermark_position, watermark_rate, text6, text8var):
        # 是否需要添加水印（用于给片头片尾文件转码）
        self.ifadd = ifadd
        self.file_path = file_path
        self.logo_path = logo_path
        self.new_file = new_file
        self.watermark_position = watermark_position
        self.waterrate_list = [(0, 0), (60, 5), (180, 15), (300, 30)]
        self.rate = self.waterrate_list[watermark_rate]
        self.text6 = text6
        self.text8var = text8var
        # 初始化时创建temp文件夹
        if not os.path.exists(self.new_file + '/temp'):
            os.mkdir(self.new_file + '/temp')

    def add_mark(self):
        vidcap = cv2.VideoCapture(self.file_path)
        # 原视频宽度
        width = int(vidcap.get(3))
        # 原视频高度
        height = int(vidcap.get(4))
        # 原视频帧率
        fps = vidcap.get(5)
        # 原视频总帧率
        total_fps = vidcap.get(7)
        #  获取文件大小（M: 兆）
        file_byte = os.path.getsize(self.file_path)
        success, image = vidcap.read()
        count = 0
        size = (width, height)  # 需要转为视频的图片的尺寸
        # 可以使用cv2.resize()进行修改
        pic = cv2.imread(self.logo_path)
        rows, cols = pic.shape[:2]
        video = cv2.VideoWriter(self.new_file + '/temp/mv.mp4', cv2.VideoWriter_fourcc('X', 'V', 'I', 'D'), fps, size)
        # video = cv2.VideoWriter(new_file, cv2.VideoWriter_fourcc('F', 'L', 'V', '1'), fps, size)
        col, row = (width, height)
        try:
            mark_position_list = [(col - cols - 20, row - rows - 20), (col - cols - 20, 20), (20, row - rows - 20),
                                  (20, 20), (math.floor((col - cols) / 2), 20),
                                  (math.floor((col - cols) / 2), row - rows - 20),
                                  (random.randint(20, col - cols - 20), random.randint(20, row - rows - 20))]
        except ValueError:
            self.text6.insert(END, "水印文件尺寸与视频尺寸不匹配，水印尺寸过大")
            return
        x, y = (mark_position_list[self.watermark_position])
        # 读取水印产生时间与频率
        j = self.rate[0]
        k = self.rate[1]
        if self.ifadd:  # 需要添加水印
            if j:
                time = math.ceil(total_fps / (j * fps))
                start = random.randint(0, j - k)
                s_time = math.floor(start * fps)
                while True:
                    self.text8var.set(
                        self.file_path.split('/')[-1] + "添加水印完成进度:{0}%".format(math.ceil(count * 100 / total_fps)))
                    success, image = vidcap.read()
                    if not success:
                        break
                    # 优化循环参数，提高效率
                    i = 0
                    # 控制水印出现的时间,如果count在指定的区域内，就显示水印
                    for n in range(i, time):
                        s = s_time + j * fps * n
                        if s <= count <= s + k * fps:
                            image[y:rows + y, x:cols + x] = pic
                            break
                    video.write(image)
                    count += 1
            else:
                while True:
                    self.text8var.set(
                        self.file_path.split('/')[-1] + "添加水印完成进度:{0}%".format(int(count * 100 / total_fps)))
                    success, image = vidcap.read()
                    if not success:
                        break
                    image[y:rows + y, x:cols + x] = pic
                    video.write(image)
                    count += 1
        else:  # 不需要添加水印，只用于重新转码
            while True:
                self.text8var.set(self.file_path.split('/')[-1] + "转码进度:{0}%".format(int(count * 100 / total_fps)))
                success, image = vidcap.read()
                if not success:
                    break
                video.write(image)
                count += 1
        video.release()
        cv2.destroyAllWindows()
        # 提取视频中的声音
        self.text8var.set(self.file_path.split('/')[-1] + "：提取音频")
        self.text6.insert(END, self.file_path.split('/')[-1] + "：提取音频\t" + datetime.datetime.now().strftime(
            '%Y-%m-%d %H:%M:%S') + "\n")
        subprocess.call("ffmpeg -y -i " + self.file_path + " -vn -acodec copy " + self.new_file + '/temp/mv.m4a')
        # 视频与声音的合并
        self.text8var.set(self.file_path.split('/')[-1] + "：合并音频与视频")
        self.text6.insert(END, self.file_path.split('/')[-1] + "：合并音频与视频\t" + datetime.datetime.now().strftime(
            '%Y-%m-%d %H:%M:%S') + "\n")
        subprocess.call("ffmpeg -y -i " + self.new_file + '/temp/mv.mp4' + " -i " + self.new_file + '/temp/mv.m4a' + \
                        " -c:v copy -c:a copy " + self.new_file + '/' + self.file_path.split('/')[-1])
        # 删除多余的文件
        if os.path.exists(self.new_file + '/temp/mv.mp4'):
            os.remove(self.new_file + '/temp/mv.mp4')
        if os.path.exists(self.new_file + '/temp/mv.m4a'):
            os.remove(self.new_file + '/temp/mv.m4a')
        if os.path.exists(self.new_file + '/temp/fps_title.mp4'):
            os.remove(self.new_file + '/temp/fps_title.mp4')


class Application_ui(Frame):
    # 这个类仅实现界面生成功能，具体事件处理代码在子类Application中。
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master.title('视频批量处理')
        self.master.geometry('600x665+{0}+{1}'.format(math.ceil((master.winfo_screenwidth() - 600) / 2),
                                                      math.ceil((master.winfo_screenheight() - 665) / 2)))
        self.master.resizable(width=False, height=False)
        self.createWidgets()
        self.Frame2 = ''
        # 给按钮分组
        self.var1 = IntVar()  # 片头片尾
        self.var2 = IntVar()  # 水印位置
        self.var3 = IntVar()  # 水印频率

    def createWidgets(self):
        self.top = self.winfo_toplevel()

        self.style = Style()

        self.style.configure('Command5.TButton', font=('宋体', 9))
        self.Command5 = Button(self.top, text='日志处理', command=self.Command5_Cmd, style='Command5.TButton')
        self.Command5.place(relx=0.667, rely=0.012, relwidth=0.333, relheight=0.05)

        self.style.configure('Command4.TButton', font=('宋体', 9))
        self.Command4 = Button(self.top, text='视频批量截取', command=self.Command4_Cmd, style='Command4.TButton')
        self.Command4.place(relx=0.333, rely=0.012, relwidth=0.333, relheight=0.05)

        self.style.configure('Command1.TButton', font=('宋体', 9))
        self.Command1 = Button(self.top, text='批量片头/片尾/水印', command=self.Command1_Cmd, style='Command1.TButton')
        self.Command1.place(relx=0., rely=0.012, relwidth=0.333, relheight=0.05)

    # 创建首页面板
    def create_frame1(self):
        self.style.configure('Frame2.TLabelframe', font=('宋体', 9))
        self.Frame2 = LabelFrame(self.top, text='', style='Frame2.TLabelframe')
        self.Frame2.place(relx=0., rely=0.072, relwidth=1.082, relheight=0.94)

        self.Text6Font = Font(font=('宋体', 9))
        self.Text6 = ScrolledText(self.Frame2, font=self.Text6Font)
        # self.Text6.config(state=DISABLED)
        self.Text6.place(relx=0.025, rely=0.717, relwidth=0.815, relheight=0.194)

        self.style.configure('Command9.TButton', font=('宋体', 9))
        self.Command9 = Button(self.Frame2, text='开始批量处理', command=self.Command9_Cmd, style='Command9.TButton')
        self.Command9.place(relx=0.037, rely=0.666, relwidth=0.174, relheight=0.034)

        self.style.configure('Command8.TButton', font=('宋体', 9))
        self.Command8 = Button(self.Frame2, text='选择位置', command=self.Command8_Cmd, style='Command8.TButton')
        self.Command8.place(relx=0.419, rely=0.611, relwidth=0.1, relheight=0.034)

        self.Text5Var = StringVar(value='')
        self.Text5 = Entry(self.Frame2, textvariable=self.Text5Var, font=('宋体', 9))
        self.Text5.place(relx=0.247, rely=0.611, relwidth=0.162, relheight=0.034)

        self.style.configure('Option14.TRadiobutton', font=('宋体', 9))
        self.Option14 = Radiobutton(self.Frame2, text='5分钟随机出现30秒', variable=self.var3, value=3,
                                    style='Option14.TRadiobutton')
        self.Option14.place(relx=0.666, rely=0.544, relwidth=0.199, relheight=0.032)

        self.style.configure('Option13.TRadiobutton', font=('宋体', 9))
        self.Option13 = Radiobutton(self.Frame2, text='3分钟随机出现15秒', variable=self.var3, value=2,
                                    style='Option13.TRadiobutton')
        self.Option13.place(relx=0.468, rely=0.544, relwidth=0.199, relheight=0.032)

        self.style.configure('Option12.TRadiobutton', font=('宋体', 9))
        self.Option12 = Radiobutton(self.Frame2, text='1分钟随机出现5秒', variable=self.var3, value=1,
                                    style='Option12.TRadiobutton')
        self.Option12.place(relx=0.284, rely=0.544, relwidth=0.186, relheight=0.032)

        self.style.configure('Option11.TRadiobutton', font=('宋体', 9))
        self.Option11 = Radiobutton(self.Frame2, text='一直在', variable=self.var3, value=0, style='Option11.TRadiobutton')
        self.Option11.place(relx=0.185, rely=0.544, relwidth=0.09, relheight=0.032)

        self.style.configure('Option10.TRadiobutton', font=('宋体', 9))
        self.Option10 = Radiobutton(self.Frame2, text='位置随机', variable=self.var2, value=6,
                                    style='Option10.TRadiobutton')
        self.Option10.place(relx=0.715, rely=0.486, relwidth=0.113, relheight=0.032)

        self.style.configure('Option9.TRadiobutton', font=('宋体', 9))
        self.Option9 = Radiobutton(self.Frame2, text='下方', variable=self.var2, value=5, style='Option9.TRadiobutton')
        self.Option9.place(relx=0.641, rely=0.486, relwidth=0.077, relheight=0.032)

        self.style.configure('Option8.TRadiobutton', font=('宋体', 9))
        self.Option8 = Radiobutton(self.Frame2, text='上方', variable=self.var2, value=4, style='Option8.TRadiobutton')
        self.Option8.place(relx=0.567, rely=0.486, relwidth=0.072, relheight=0.032)

        self.style.configure('Option7.TRadiobutton', font=('宋体', 9))
        self.Option7 = Radiobutton(self.Frame2, text='左上角', variable=self.var2, value=3, style='Option7.TRadiobutton')
        self.Option7.place(relx=0.468, rely=0.486, relwidth=0.088, relheight=0.032)

        self.style.configure('Option6.TRadiobutton', font=('宋体', 9))
        self.Option6 = Radiobutton(self.Frame2, text='左下角', variable=self.var2, value=2, style='Option6.TRadiobutton')
        self.Option6.place(relx=0.375, rely=0.486, relwidth=0.088, relheight=0.032)

        self.style.configure('Option5.TRadiobutton', font=('宋体', 9))
        self.Option5 = Radiobutton(self.Frame2, text='右上角', variable=self.var2, value=1, style='Option5.TRadiobutton')
        self.Option5.place(relx=0.282, rely=0.486, relwidth=0.088, relheight=0.032)

        self.style.configure('Option4.TRadiobutton', font=('宋体', 9))
        self.Option4 = Radiobutton(self.Frame2, text='右下角', variable=self.var2, value=0, style='Option4.TRadiobutton')
        self.Option4.place(relx=0.185, rely=0.486, relwidth=0.088, relheight=0.032)

        self.Check2Var = StringVar(value='0')
        self.style.configure('Check2.TCheckbutton', font=('宋体', 9))
        self.Check2 = Checkbutton(self.Frame2, text='片头/片尾文件也添加水印', variable=self.Check2Var,
                                  style='Check2.TCheckbutton')
        self.Check2.place(relx=0.468, rely=0.435, relwidth=0.26, relheight=0.027)

        self.style.configure('Command7.TButton', font=('宋体', 9))
        self.Command7 = Button(self.Frame2, text='选择文件', command=self.Command7_Cmd, style='Command7.TButton')
        self.Command7.place(relx=0.32, rely=0.431, relwidth=0.1, relheight=0.034)

        self.Text4Var = StringVar(value='')
        self.Text4 = Entry(self.Frame2, textvariable=self.Text4Var, font=('宋体', 9))
        self.Text4.place(relx=0.173, rely=0.431, relwidth=0.137, relheight=0.034)

        self.Combo1List = [0, 1, 2, 3, 4, 5]
        self.Combo1 = Combobox(self.Frame2, values=self.Combo1List, font=('宋体', 9))
        self.Combo1.place(relx=0.173, rely=0.379, relwidth=0.066, relheight=0.032)
        self.Combo1.set(self.Combo1List[0])

        self.style.configure('Command6.TButton', font=('宋体', 9))
        self.Command6 = Button(self.Frame2, text='选择文件', command=self.Command6_Cmd, style='Command6.TButton')
        self.Command6.place(relx=0.678, rely=0.331, relwidth=0.112, relheight=0.034)

        self.Text3Var = StringVar(value='')
        self.Text3 = Entry(self.Frame2, textvariable=self.Text3Var, font=('宋体', 9))
        self.Text3.place(relx=0.493, rely=0.331, relwidth=0.174, relheight=0.034)

        self.style.configure('Option2.TRadiobutton', font=('宋体', 9))
        self.Option2 = Radiobutton(self.Frame2, text='片尾', variable=self.var1, value=2, style='Option2.TRadiobutton')
        self.Option2.place(relx=0.357, rely=0.333, relwidth=0.076, relheight=0.027)

        self.style.configure('Option1.TRadiobutton', font=('宋体', 9))
        self.Option1 = Radiobutton(self.Frame2, text='片头', variable=self.var1, value=1, style='Option1.TRadiobutton')
        self.Option1.place(relx=0.259, rely=0.333, relwidth=0.076, relheight=0.027)

        self.style.configure('Option15.TRadiobutton', font=('宋体', 9))
        self.Option15 = Radiobutton(self.Frame2, text='无', variable=self.var1, value=0, style='Option15.TRadiobutton')
        self.Option15.place(relx=0.173, rely=0.333, relwidth=0.063, relheight=0.032)

        self.Text2Font = Font(font=('宋体', 9))
        self.Text2 = ScrolledText(self.Frame2, font=self.Text2Font)
        # self.Text2.config(state=DISABLED)
        self.Text2.place(relx=0.012, rely=0.077, relwidth=0.827, relheight=0.232)

        self.style.configure('Command3.TButton', font=('宋体', 9))
        self.Command3 = Button(self.Frame2, text='列出视频文件', command=self.Command3_Cmd, style='Command3.TButton')
        self.Command3.place(relx=0.727, rely=0.025, relwidth=0.162, relheight=0.034)

        self.Check1Var = StringVar(value='0')
        self.style.configure('Check1.TCheckbutton', font=('宋体', 9))
        self.Check1 = Checkbutton(self.Frame2, text='显示子文件夹下视频', variable=self.Check1Var,
                                  style='Check1.TCheckbutton')
        self.Check1.place(relx=0.468, rely=0.028, relwidth=0.236, relheight=0.032)

        self.style.configure('Command2.TButton', font=('宋体', 9))
        self.Command2 = Button(self.Frame2, text='浏览', command=self.Command2_Cmd, style='Command2.TButton')
        self.Command2.place(relx=0.333, rely=0.026, relwidth=0.076, relheight=0.035)

        self.Text1Var = StringVar(value='')
        self.Text1 = Entry(self.Frame2, textvariable=self.Text1Var, font=('宋体', 9))
        self.Text1.place(relx=0.185, rely=0.026, relwidth=0.137, relheight=0.034)

        self.style.configure('Label7.TLabel', anchor='w', font=('宋体', 9))
        self.Label7 = Label(self.Frame2, text='生成新视频所在位置', style='Label7.TLabel')
        self.Label7.place(relx=0.025, rely=0.614, relwidth=0.174, relheight=0.032)

        self.style.configure('Label6.TLabel', anchor='w', font=('宋体', 9))
        self.Label6 = Label(self.Frame2, text='水印频率', style='Label6.TLabel')
        self.Label6.place(relx=0.068, rely=0.544, relwidth=0.09, relheight=0.032)

        self.style.configure('Label5.TLabel', anchor='w', font=('宋体', 9))
        self.Label5 = Label(self.Frame2, text='水印位置', style='Label5.TLabel')
        self.Label5.place(relx=0.068, rely=0.491, relwidth=0.09, relheight=0.027)

        self.style.configure('Label4.TLabel', anchor='w', font=('宋体', 9))
        self.Label4 = Label(self.Frame2, text='水印文件', style='Label4.TLabel')
        self.Label4.place(relx=0.068, rely=0.435, relwidth=0.09, relheight=0.027)

        self.style.configure('Label3.TLabel', anchor='w', font=('宋体', 9))
        self.Label3 = Label(self.Frame2, text='添加频率', style='Label3.TLabel')
        self.Label3.place(relx=0.068, rely=0.384, relwidth=0.09, relheight=0.027)

        self.style.configure('Label2.TLabel', anchor='w', font=('宋体', 9))
        self.Label2 = Label(self.Frame2, text='片头/片尾文件', style='Label2.TLabel')
        self.Label2.place(relx=0.025, rely=0.335, relwidth=0.137, relheight=0.027)

        self.style.configure('Label1.TLabel', anchor='w', font=('宋体', 9))
        self.Label1 = Label(self.Frame2, text='视频所在文件夹', style='Label1.TLabel')
        self.Label1.place(relx=0.025, rely=0.03, relwidth=0.137, relheight=0.032)

        self.style.configure('Label8.TLabel', anchor='w', font=('宋体', 9))
        self.Text8Var = StringVar(value='')
        self.Label8 = Label(self.Frame2, textvariable=self.Text8Var, style='Label8.TLabel')
        self.Label8.place(relx=0.025, rely=0.922, relwidth=0.715, relheight=0.034)

        self.style.configure('Label9.TLabel', anchor='w', font=('宋体', 11))
        self.labelVar = StringVar(value='')
        self.Label9 = Label(self.Frame2, textvariable=self.labelVar, style='Label9.TLabel')
        self.Label9.place(relx=0.750, rely=0.922, relwidth=0.1, relheight=0.034)

    # 创建视频截取面板
    def create_frame2(self):
        self.style.configure('Frame2.TLabelframe', font=('宋体', 9))
        self.Frame2 = LabelFrame(self.top, text='', style='Frame2.TLabelframe')
        self.Frame2.place(relx=0., rely=0.072, relwidth=1.082, relheight=0.94)

        self.style.configure('Label8.TLabel', anchor='w', font=('宋体', 9))
        self.Label8 = Label(self.Frame2, text='截取时间文件', style='Label8.TLabel')
        self.Label8.place(relx=0.025, rely=0.087, relwidth=0.118, relheight=0.027)

        self.style.configure('Command10.TButton', font=('宋体', 9))
        self.Command10 = Button(self.Frame2, text='加载截取时间文件', command=self.Command10_Cmd, style='Command10.TButton')
        self.Command10.place(relx=0.333, rely=0.085, relwidth=0.175, relheight=0.032)

        self.Text7Var = StringVar(value='')
        self.Text7 = Entry(self.Frame2, textvariable=self.Text7Var, font=('宋体', 9))
        self.Text7.place(relx=0.185, rely=0.085, relwidth=0.137, relheight=0.034)

        self.Text6Font = Font(font=('宋体', 9))
        self.Text6 = ScrolledText(self.Frame2, font=self.Text6Font)
        # self.Text6.config(state=DISABLED)
        self.Text6.place(relx=0.025, rely=0.717, relwidth=0.815, relheight=0.194)

        self.style.configure('Command12.TButton', font=('宋体', 9))
        self.Command12 = Button(self.Frame2, text='开始批量处理', command=self.Command12_Cmd, style='Command12.TButton')
        self.Command12.place(relx=0.037, rely=0.666, relwidth=0.174, relheight=0.034)

        self.style.configure('Command8.TButton', font=('宋体', 9))
        self.Command8 = Button(self.Frame2, text='选择位置', command=self.Command8_Cmd, style='Command8.TButton')
        self.Command8.place(relx=0.419, rely=0.611, relwidth=0.1, relheight=0.034)

        self.Text5Var = StringVar(value='')
        self.Text5 = Entry(self.Frame2, textvariable=self.Text5Var, font=('宋体', 9))
        self.Text5.place(relx=0.247, rely=0.611, relwidth=0.162, relheight=0.034)

        self.style.configure('Option14.TRadiobutton', font=('宋体', 9))
        self.Option14 = Radiobutton(self.Frame2, text='5分钟随机出现30秒', variable=self.var3, value=3,
                                    style='Option14.TRadiobutton')
        self.Option14.place(relx=0.666, rely=0.544, relwidth=0.199, relheight=0.032)

        self.style.configure('Option13.TRadiobutton', font=('宋体', 9))
        self.Option13 = Radiobutton(self.Frame2, text='3分钟随机出现15秒', variable=self.var3, value=2,
                                    style='Option13.TRadiobutton')
        self.Option13.place(relx=0.468, rely=0.544, relwidth=0.199, relheight=0.032)

        self.style.configure('Option12.TRadiobutton', font=('宋体', 9))
        self.Option12 = Radiobutton(self.Frame2, text='1分钟随机出现5秒', variable=self.var3, value=1,
                                    style='Option12.TRadiobutton')
        self.Option12.place(relx=0.284, rely=0.544, relwidth=0.186, relheight=0.032)

        self.style.configure('Option11.TRadiobutton', font=('宋体', 9))
        self.Option11 = Radiobutton(self.Frame2, text='一直在', variable=self.var3, value=0, style='Option11.TRadiobutton')
        self.Option11.place(relx=0.185, rely=0.544, relwidth=0.09, relheight=0.032)

        self.style.configure('Option10.TRadiobutton', font=('宋体', 9))
        self.Option10 = Radiobutton(self.Frame2, text='位置随机', variable=self.var2, value=6,
                                    style='Option10.TRadiobutton')
        self.Option10.place(relx=0.715, rely=0.486, relwidth=0.113, relheight=0.032)

        self.style.configure('Option9.TRadiobutton', font=('宋体', 9))
        self.Option9 = Radiobutton(self.Frame2, text='下方', variable=self.var2, value=5, style='Option9.TRadiobutton')
        self.Option9.place(relx=0.641, rely=0.486, relwidth=0.077, relheight=0.032)

        self.style.configure('Option8.TRadiobutton', font=('宋体', 9))
        self.Option8 = Radiobutton(self.Frame2, text='上方', variable=self.var2, value=4, style='Option8.TRadiobutton')
        self.Option8.place(relx=0.567, rely=0.486, relwidth=0.072, relheight=0.032)

        self.style.configure('Option7.TRadiobutton', font=('宋体', 9))
        self.Option7 = Radiobutton(self.Frame2, text='左上角', variable=self.var2, value=3, style='Option7.TRadiobutton')
        self.Option7.place(relx=0.468, rely=0.486, relwidth=0.088, relheight=0.032)

        self.style.configure('Option6.TRadiobutton', font=('宋体', 9))
        self.Option6 = Radiobutton(self.Frame2, text='左下角', variable=self.var2, value=2, style='Option6.TRadiobutton')
        self.Option6.place(relx=0.375, rely=0.486, relwidth=0.088, relheight=0.032)

        self.style.configure('Option5.TRadiobutton', font=('宋体', 9))
        self.Option5 = Radiobutton(self.Frame2, text='右上角', variable=self.var2, value=1, style='Option5.TRadiobutton')
        self.Option5.place(relx=0.282, rely=0.486, relwidth=0.088, relheight=0.032)

        self.style.configure('Option4.TRadiobutton', font=('宋体', 9))
        self.Option4 = Radiobutton(self.Frame2, text='右下角', variable=self.var2, value=0, style='Option4.TRadiobutton')
        self.Option4.place(relx=0.185, rely=0.486, relwidth=0.088, relheight=0.032)

        self.Check2Var = StringVar(value='0')
        self.style.configure('Check2.TCheckbutton', font=('宋体', 9))
        self.Check2 = Checkbutton(self.Frame2, text='片头/片尾文件也添加水印', variable=self.Check2Var,
                                  style='Check2.TCheckbutton')
        self.Check2.place(relx=0.468, rely=0.435, relwidth=0.26, relheight=0.027)

        self.style.configure('Command7.TButton', font=('宋体', 9))
        self.Command7 = Button(self.Frame2, text='选择文件', command=self.Command7_Cmd, style='Command7.TButton')
        self.Command7.place(relx=0.32, rely=0.431, relwidth=0.1, relheight=0.034)

        self.Text4Var = StringVar(value='')
        self.Text4 = Entry(self.Frame2, textvariable=self.Text4Var, font=('宋体', 9))
        self.Text4.place(relx=0.173, rely=0.431, relwidth=0.137, relheight=0.034)

        self.Combo1List = [0, 1, 2, 3, 4, 5]
        self.Combo1 = Combobox(self.Frame2, values=self.Combo1List, font=('宋体', 9))
        self.Combo1.place(relx=0.173, rely=0.379, relwidth=0.066, relheight=0.032)
        self.Combo1.set(self.Combo1List[0])

        self.style.configure('Command6.TButton', font=('宋体', 9))
        self.Command6 = Button(self.Frame2, text='选择文件', command=self.Command6_Cmd, style='Command6.TButton')
        self.Command6.place(relx=0.678, rely=0.331, relwidth=0.112, relheight=0.034)

        self.Text3Var = StringVar(value='')
        self.Text3 = Entry(self.Frame2, textvariable=self.Text3Var, font=('宋体', 9))
        self.Text3.place(relx=0.493, rely=0.331, relwidth=0.174, relheight=0.034)

        self.style.configure('Option2.TRadiobutton', font=('宋体', 9))
        self.Option2 = Radiobutton(self.Frame2, text='片尾', variable=self.var1, value=2, style='Option2.TRadiobutton')
        self.Option2.place(relx=0.357, rely=0.333, relwidth=0.076, relheight=0.027)

        self.style.configure('Option1.TRadiobutton', font=('宋体', 9))
        self.Option1 = Radiobutton(self.Frame2, text='片头', variable=self.var1, value=1, style='Option1.TRadiobutton')
        self.Option1.place(relx=0.259, rely=0.333, relwidth=0.076, relheight=0.027)

        self.style.configure('Option15.TRadiobutton', font=('宋体', 9))
        self.Option15 = Radiobutton(self.Frame2, text='无', variable=self.var1, value=0, style='Option15.TRadiobutton')
        self.Option15.place(relx=0.173, rely=0.333, relwidth=0.063, relheight=0.032)

        self.Text2Font = Font(font=('宋体', 9))
        self.Text2 = ScrolledText(self.Frame2, font=self.Text2Font)
        # self.Text2.config(state=DISABLED)
        self.Text2.place(relx=0.025, rely=0.128, relwidth=0.815, relheight=0.181)

        self.style.configure('Command3.TButton', font=('宋体', 9))
        self.Command3 = Button(self.Frame2, text='列出视频文件', command=self.Command3_Cmd, style='Command3.TButton')
        self.Command3.place(relx=0.727, rely=0.025, relwidth=0.162, relheight=0.034)

        self.Check1Var = StringVar(value='0')
        self.style.configure('Check1.TCheckbutton', font=('宋体', 9))
        self.Check1 = Checkbutton(self.Frame2, text='显示子文件夹下视频', variable=self.Check1Var,
                                  style='Check1.TCheckbutton')
        self.Check1.place(relx=0.468, rely=0.028, relwidth=0.236, relheight=0.032)

        self.style.configure('Command2.TButton', font=('宋体', 9))
        self.Command2 = Button(self.Frame2, text='浏览', command=self.Command2_Cmd, style='Command2.TButton')
        self.Command2.place(relx=0.333, rely=0.026, relwidth=0.076, relheight=0.035)

        self.Text1Var = StringVar(value='')
        self.Text1 = Entry(self.Frame2, textvariable=self.Text1Var, font=('宋体', 9))
        self.Text1.place(relx=0.185, rely=0.026, relwidth=0.137, relheight=0.034)

        self.style.configure('Label7.TLabel', anchor='w', font=('宋体', 9))
        self.Label7 = Label(self.Frame2, text='生成新视频所在位置', style='Label7.TLabel')
        self.Label7.place(relx=0.025, rely=0.614, relwidth=0.174, relheight=0.032)

        self.style.configure('Label6.TLabel', anchor='w', font=('宋体', 9))
        self.Label6 = Label(self.Frame2, text='水印频率', style='Label6.TLabel')
        self.Label6.place(relx=0.068, rely=0.544, relwidth=0.09, relheight=0.032)

        self.style.configure('Label5.TLabel', anchor='w', font=('宋体', 9))
        self.Label5 = Label(self.Frame2, text='水印位置', style='Label5.TLabel')
        self.Label5.place(relx=0.068, rely=0.491, relwidth=0.09, relheight=0.027)

        self.style.configure('Label4.TLabel', anchor='w', font=('宋体', 9))
        self.Label4 = Label(self.Frame2, text='水印文件', style='Label4.TLabel')
        self.Label4.place(relx=0.068, rely=0.435, relwidth=0.09, relheight=0.027)

        self.style.configure('Label3.TLabel', anchor='w', font=('宋体', 9))
        self.Label3 = Label(self.Frame2, text='添加频率', style='Label3.TLabel')
        self.Label3.place(relx=0.068, rely=0.384, relwidth=0.09, relheight=0.027)

        self.style.configure('Label2.TLabel', anchor='w', font=('宋体', 9))
        self.Label2 = Label(self.Frame2, text='片头/片尾文件', style='Label2.TLabel')
        self.Label2.place(relx=0.025, rely=0.335, relwidth=0.137, relheight=0.027)

        self.style.configure('Label1.TLabel', anchor='w', font=('宋体', 9))
        self.Label1 = Label(self.Frame2, text='视频所在文件夹', style='Label1.TLabel')
        self.Label1.place(relx=0.025, rely=0.03, relwidth=0.137, relheight=0.032)

        self.style.configure('Label8.TLabel', anchor='w', font=('宋体', 9))
        self.Text8Var = StringVar(value='')
        self.Label8 = Label(self.Frame2, textvariable=self.Text8Var, style='Label8.TLabel')
        self.Label8.place(relx=0.025, rely=0.922, relwidth=0.715, relheight=0.034)

        self.style.configure('Label9.TLabel', anchor='w', font=('宋体', 11))
        self.labelVar = StringVar(value='')
        self.Label9 = Label(self.Frame2, textvariable=self.labelVar, style='Label9.TLabel')
        self.Label9.place(relx=0.750, rely=0.922, relwidth=0.1, relheight=0.034)

    # 创建日志处理面板
    def create_frame3(self):
        self.style.configure('Frame2.TLabelframe', font=('宋体', 9))
        self.Frame2 = LabelFrame(self.top, text='', style='Frame2.TLabelframe')
        self.Frame2.place(relx=0., rely=0.072, relwidth=1.082, relheight=0.94)

        self.Text7Font = Font(font=('宋体', 9))
        self.Text7 = ScrolledText(self.Frame2, font=self.Text7Font)
        self.Text7.place(relx=0.037, rely=0.181, relwidth=0.815, relheight=0.424)

        self.style.configure('Label11.TLabel', anchor='w', font=('宋体', 11))
        self.Label11 = Label(self.Frame2, text='选择日志', style='Label11.TLabel')
        self.Label11.place(relx=0.049, rely=0.075, relwidth=0.103, relheight=0.04)

        self.Combo2List = ['请选择...']
        # 加载日志文件
        fiel_path = (os.getcwd() + '/log/').replace('\\', '/')
        file_list = os.listdir(fiel_path)
        for file in file_list:
            self.Combo2List.append(file)
        self.Combo2 = Combobox(self.Frame2, values=self.Combo2List, font=('宋体', 9), state='readonly')
        self.Combo2.bind('<<ComboboxSelected>>', self.on_select)
        self.Combo2.place(relx=0.166, rely=0.075, relwidth=0.29, relheight=0.04)
        self.Combo2.set(self.Combo2List[0])

        self.style.configure('Command14.TButton', font=('宋体', 9))
        self.Command14 = Button(self.Frame2, text='删除此条日志', command=self.Command14_Cmd, style='Command14.TButton')
        self.Command14.place(relx=0.493, rely=0.075, relwidth=0.14, relheight=0.04)

        self.style.configure('Command15.TButton', font=('宋体', 9))
        self.Command15 = Button(self.Frame2, text='删除所有日志', command=self.Command15_Cmd, style='Command15.TButton')
        self.Command15.place(relx=0.641, rely=0.075, relwidth=0.14, relheight=0.04)

    def on_select(self, event=None):
        file_name = event.widget.get()
        if file_name.split('.')[-1] == 'txt':
            path = (os.getcwd() + '/log/').replace('\\', '/') + file_name
            if os.path.exists(path):
                file_context = open(path, encoding='UTF-8').read()
                self.Text7.delete(1.0, END)
                self.Text7.insert(END, file_context)


class Application(Application_ui):
    # 这个类实现具体的事件处理回调函数。界面生成代码在Application_ui中。
    def __init__(self, master=None):
        Application_ui.__init__(self, master)
        # 处理视频的格式限制
        self.video_type = 'mp4'
        self.title_type = 'mp4'
        # 添加水印格式的限制
        self.mark_type = 'png'
        # 添加时间截取文件格式限制
        self.excel_type = 'xls'
        # 判断加水印线程是否执行完
        self.flag = False
        self.flag1 = False
        self.count_file = 0  # 计算当前处理数
        # 成功个数
        self.success = 0
        # 失败个数
        self.fail = 0
        self.create_frame1()
        # 初始化时创建log文件夹
        if not os.path.exists(os.getcwd() + '\\log\\'):
            os.mkdir(os.getcwd() + '\\log\\')

    def Command5_Cmd(self):
        if not self.flag and not self.flag1:
            if self.Frame2:
                self.Frame2.destroy()
            self.create_frame3()
        else:
            tkinter.messagebox.showinfo('提示', '请等待当前程序执行完成')

    def Command10_Cmd(self):
        fname = filedialog.askopenfilename(title='选择水印文件', filetypes=[('文件类型', '*.' + self.excel_type)])
        self.Text7Var.set(fname)

    def Command4_Cmd(self):
        if not self.flag and not self.flag1:
            if self.Frame2:
                self.Frame2.destroy()
            self.create_frame2()
        else:
            tkinter.messagebox.showinfo('提示', '请等待当前程序执行完成')

    def Command1_Cmd(self):
        if not self.flag and not self.flag1:
            if self.Frame2:
                self.Frame2.destroy()
            self.create_frame1()
        else:
            tkinter.messagebox.showinfo('提示', '请等待当前程序执行完成')

    # 删除此条日志
    def Command14_Cmd(self, event=None):
        # 得到combobox选中的文件
        file_name = self.Combo2.get()
        if file_name.split('.')[-1] == 'txt':
            flag = tkinter.messagebox.askquestion('提示', '确认删除' + file_name + "?")
            print(flag)
            if flag == 'yes':
                path_f = (os.getcwd() + '/log/').replace('\\', '/') + file_name
                if os.path.exists(path_f):
                    os.remove(path_f)
                    self.Combo2List.remove(file_name)
                else:
                    flag = tkinter.messagebox.showinfo('提示', '文件不存在或已删除')

    # 删除所有日志
    def Command15_Cmd(self, event=None):
        path = (os.getcwd() + '/log/').replace('\\', '/')
        if os.listdir(path):
            flag = tkinter.messagebox.askquestion('警告', '确认删除所有日志？此操作不可撤销！！')
            if flag == 'yes':
                if tkinter.messagebox.askquestion('警告', '确认删除') == 'yes':
                    for i in os.listdir(path):
                        path_file = os.path.join(path, i)  # 取文件路径
                        if os.path.isfile(path_file):
                            os.remove(path_file)
                    self.Combo2List.clear()

    # 视频截取面板中获得视频处理的所有参数，开始处理视频
    def Command12_Cmd(self):
        # 先获取文本框与Checkbutton的值
        fname = self.Text1Var.get()
        if fname:
            # 查找文件夹下的文件
            video_list = self.search_file(fname, self.Check1Var.get())
            # 获取片头/片尾文件选项 无是0，片头是1，片尾是2,返回值为int
            title_file = self.var1.get()
            # 片头片尾是否也要添加水印,返回值为string，0为不添加，1为添加
            add_mark = int(self.Check2Var.get())
            # 获取片头/片尾文件地址
            title_addr = self.Text3Var.get()
            # 获取截取时间文件地址
            time_addr = self.Text7Var.get()
            # 获取添加频率,返回的为string类型
            combo_text = self.Combo1.get()
            # 获取水印文件地址
            watermark_addr = self.Text4Var.get()
            # 获取水印位置
            # 0：右下角 1：右上角 2：左下角 3：左上角 4：上方 5：下方 6：位置随机 ,返回值为int
            watermark_position = self.var2.get()
            # 获取水印频率
            # 0：一直在 1：1分钟随机出现5秒 2：3分钟随机出现15秒 3：5分钟随机出现30秒 ,返回值为int
            watermark_rate = self.var3.get()
            # 获取新视频所在位置
            new_video_addr = self.Text5Var.get()

            # 进行逻辑判断
            if combo_text.isdigit():
                if time_addr:
                    if watermark_addr:
                        if new_video_addr:
                            if title_file:  # 需要添加片头或者片尾文件
                                if not title_addr:  # 片头片尾文件没有选中，需要报错
                                    tkinter.messagebox.showinfo('提示', '需要选择片头或者片尾文件')
                                    return
                            if not self.flag1:
                                self.count_file = 0  # 清零
                                # 写入日志文件
                                f = open(os.getcwd() + '\\log\\' + datetime.datetime.now().strftime(
                                    '%Y%m%d %H%M%S' + '.txt'), 'w', encoding='utf-8')
                                # 修改flag属性，当前方法没执行完，不允许新线程进来
                                self.flag1 = True
                                # 清空文本框
                                self.Text6.delete('1.0', END)
                                self.Text8Var.set('')
                                # 读取Excel文件，将数据保存到列表中
                                excel_list = self.read_excel(time_addr)
                                if not excel_list:
                                    self.flag1 = False
                                    return
                                # 对选择文件夹里的视频逐一操作
                                t = threading.Thread(target=self.execute_core1, args=(
                                    video_list, title_file, title_addr, combo_text,
                                    watermark_addr, watermark_position, watermark_rate,
                                    new_video_addr, add_mark, f, excel_list))
                                t.start()
                            else:
                                tkinter.messagebox.showinfo('提示', '请等待当前程序执行完成')
                        else:
                            tkinter.messagebox.showinfo('提示', '需要选择处理后文件所放文件夹')
                    else:
                        tkinter.messagebox.showinfo('提示', '需要选择水印文件')
                else:
                    tkinter.messagebox.showinfo('提示', '需要加载截取时间文件')
            else:
                tkinter.messagebox.showinfo('提示', '添加频率需输入正整数')
        else:
            tkinter.messagebox.showinfo('提示', '需要选择处理视频文件夹')

    # 获得视频处理的所有参数，开始处理视频
    def Command9_Cmd(self):
        # 先获取文本框与Checkbutton的值
        fname = self.Text1Var.get()
        if fname:
            # 查找文件夹下的文件
            video_list = self.search_file(fname, self.Check1Var.get())
            # 获取片头/片尾文件选项 无是0，片头是1，片尾是2,返回值为int
            title_file = self.var1.get()
            # 片头片尾是否也要添加水印,返回值为string，0为不添加，1为添加
            add_mark = int(self.Check2Var.get())
            # 获取片头/片尾文件地址
            title_addr = self.Text3Var.get()
            # 获取添加频率,返回的为string类型
            combo_text = self.Combo1.get()
            # 获取水印文件地址
            watermark_addr = self.Text4Var.get()
            # 获取水印位置
            # 0：右下角 1：右上角 2：左下角 3：左上角 4：上方 5：下方 6：位置随机 ,返回值为int
            watermark_position = self.var2.get()
            # 获取水印频率
            # 0：一直在 1：1分钟随机出现5秒 2：3分钟随机出现15秒 3：5分钟随机出现30秒 ,返回值为int
            watermark_rate = self.var3.get()
            # 获取新视频所在位置
            new_video_addr = self.Text5Var.get()

            # 进行逻辑判断
            if combo_text.isdigit():
                if watermark_addr:
                    if new_video_addr:
                        if title_file:  # 需要添加片头或者片尾文件
                            if not title_addr:  # 片头片尾文件没有选中，需要报错
                                tkinter.messagebox.showinfo('提示', '需要选择片头或者片尾文件')
                                return
                        if not self.flag:
                            self.count_file = 0  # 清零
                            # 修改flag属性，当前方法没执行完，不允许新线程进来
                            self.flag = True
                            # 清空文本框
                            self.Text6.delete('1.0', END)
                            self.Text8Var.set('')
                            # 写入日志文件
                            f = open(os.getcwd() + '\\log\\' + datetime.datetime.now().strftime(
                                '%Y%m%d %H%M%S' + '.txt'), 'w', encoding='utf-8')
                            # 操作开始时间
                            start = time.time()
                            self.Text6.insert(END, "开始执行...\n")
                            t = threading.Thread(target=self.cerete_new_video, args=(
                                start, video_list, title_file, title_addr, combo_text, watermark_addr,
                                watermark_position, watermark_rate, new_video_addr, add_mark, f))
                            t.start()
                        else:
                            tkinter.messagebox.showinfo('提示', '请等待当前程序执行完成')
                    else:
                        tkinter.messagebox.showinfo('提示', '需要选择处理后文件所放文件夹')
                else:
                    tkinter.messagebox.showinfo('提示', '需要选择水印文件')
            else:
                tkinter.messagebox.showinfo('提示', '添加频率需输入正整数')
        else:
            tkinter.messagebox.showinfo('提示', '需要选择处理视频文件夹')

    def Command8_Cmd(self):
        fname = filedialog.askdirectory(title="选择处理后视频存放文件夹")
        self.Text5Var.set(fname)

    def Command7_Cmd(self):
        fname = filedialog.askopenfilename(title='选择水印文件', filetypes=(('文件类型', '*.' + self.mark_type), ('所有类型', '*.*')))
        self.Text4Var.set(fname)

    def Command6_Cmd(self):
        fname = filedialog.askopenfilename(title='选择片头或片尾文件',
                                           filetypes=(('文件类型', '*.' + self.title_type), ('所有类型', '*.*')))
        self.Text3Var.set(fname)

    def Command3_Cmd(self):
        # 先获取文本框与Checkbutton的值
        fname = self.Text1Var.get()
        if fname:
            # 查找文件夹下的文件
            video_list = self.search_file(fname, self.Check1Var.get())
            # 将查出的文件写入到Text2
            num = 1
            # 先清空文本框
            self.Text2.delete('1.0', END)
            for file in video_list:
                file = file.replace('\\', '/')
                self.Text2.insert(END, str(num) + ':' + file)
                num += 1
        else:
            tkinter.messagebox.showinfo('提示', '需要选择处理视频文件夹')

    def Command2_Cmd(self):
        fname = filedialog.askdirectory(title=u"选择要处理视频所在文件夹")
        self.Text1Var.set(fname)

    # 查找文件夹下的文件
    def search_file(self, fname, flag):
        video_list = []
        # 根据路径查找文件
        if int(flag):
            for root, dirs, files in os.walk(fname):
                for name in files:
                    if name.split('.')[-1] == self.video_type or name.split('.')[-1] == 'flv' or name.split('.')[
                        -1] == 'avi':
                        video_list.append(os.path.join(root, name) + '\n')
        else:
            list_dir = os.listdir(fname)
            for f in list_dir:
                if f.split('.')[-1] == self.video_type or f.split('.')[-1] == 'flv' or f.split('.')[-1] == 'avi':
                    filepath = os.path.join(fname, f)
                    if os.path.isfile(filepath):
                        video_list.append(filepath + '\n')

        return video_list

    def execute_core1(self, video_list, title_file, title_addr, combo_text, watermark_addr, watermark_position,
                      watermark_rate, new_video_addr, add_mark, f, excel_list):
        # 操作开始时间
        start = time.time()
        # 总共要操作的视频个数--》分割后的视频
        count = 0
        for video_path in video_list:
            for x in excel_list:
                if x[0] == video_path.replace('\\', '/').rstrip('\n').split('/')[-1]:
                    count += 1
        for video_path in video_list:
            if not os.path.exists(new_video_addr + '/temp'):
                os.mkdir(new_video_addr + '/temp')
            path = video_path.replace('\\', '/').rstrip('\n')
            # 获取文件名
            str_name = path.split('/')[-1]
            # 文件名与excel里文件名比对，符合的放在一个列表里，进行统一处理
            file_list = []
            self.Text6.insert(END, "开始执行...\n")
            for x in excel_list:
                if x[0] == str_name:
                    split_path = self.slice_video(path, new_video_addr, x)
                    if split_path != '':
                        file_list.append(split_path)
            # 解除ffempg对文件的控制
            os.system('taskkill /F /IM ffmpeg-win64-v4.1.exe')
            # 视频批量加水印
            if len(file_list) == 0:
                tkinter.messagebox.showinfo('提示', '所选截取时间文件里不包含' + str_name + '的信息')
                self.flag1 = False
                return
            self.cerete_new_video(start, file_list, title_file, title_addr, combo_text, watermark_addr,
                                  watermark_position, watermark_rate, new_video_addr, add_mark, f, count)
            # 删除生成的分割文件
            for file in file_list:
                if os.path.exists(file):
                    os.remove(file)
        self.flag1 = False

    # 核心代码，根据传入参数生成既定要求的文件，并写入text6文本框，生成相应的日志文件
    def cerete_new_video(self, start, video_list, title_file, title_addr, combo_text, watermark_addr,
                         watermark_position, watermark_rate, new_video_addr, add_mark, f, length=0):
        if len(video_list) == 0:
            tkinter.messagebox.showinfo('提示', '所选文件夹里不包含要操作视频')
            self.flag = False
            return 'error'
        if length == 0:
            length = len(video_list)
        # 单次操作开始时间
        start_time = 0
        # 操作结束时间
        end_time = 0
        temp = 0
        if title_file:  # 需要添加片头片尾文件
            # 将片头片尾文件通过opencv-python转码
            WaterMark(add_mark, title_addr, watermark_addr, new_video_addr, watermark_position, watermark_rate,
                      self.Text6, self.Text8Var).add_mark()
            for video_path in video_list:
                self.labelVar.set("{0}/{1}".format(self.count_file + 1, length))
                if not int(combo_text) or temp % (int(combo_text) + 1) == 0:
                    try:
                        start_time = time.time()
                        path = video_path.replace('\\', '/').rstrip('\n')
                        insert_text = path.split('/')[-1] + '\t正在处理\t开始时间' + datetime.datetime.now().strftime(
                            '%Y-%m-%d %H:%M:%S') + '\n'
                        self.Text6.insert(END, insert_text)
                        # 添加水印
                        WaterMark(True, path, watermark_addr, new_video_addr, watermark_position, watermark_rate,
                                  self.Text6, self.Text8Var).add_mark()
                        # 合并视频
                        # 将片头片尾文件与视频合并
                        self.Text8Var.set('将片头片尾文件' + title_addr.split('/')[-1] + '与' + path.split('/')[-1] + '合并\t')
                        self.Text6.insert(END, '将片头片尾文件' + title_addr.split('/')[-1] + '与' + path.split('/')[
                            -1] + '合并\t' + datetime.datetime.now().strftime(
                            '%Y-%m-%d %H:%M:%S') + "\n")
                        self.concat_video(title_file, title_addr.split('/')[-1], path.split('/')[-1], new_video_addr)
                        self.Text6.insert(END, '成功合并\t' + datetime.datetime.now().strftime(
                            '%Y-%m-%d %H:%M:%S') + "\n")
                        end_time = time.time()
                        self.success += 1
                    except Exception as e:
                        end_time = time.time()
                        self.Text6.insert(END, e)
                        self.Text6.insert(END, '\n')
                        self.fail += 1
                        continue
                    finally:
                        # 当次操作耗时
                        per_time = end_time - start_time
                        h = per_time / 3600
                        m = per_time % 3600 / 60
                        s = per_time % 3600 % 60
                        str1 = "{0}:{1}:{2}".format(int(h), int(m), int(s))
                        insert_text = path.split('/')[-1] + '\t处理完毕' + datetime.datetime.now().strftime(
                            '%Y-%m-%d %H:%M:%S') + '\n当次总共耗时' + str1 + '\n'
                        self.Text6.insert(END, insert_text)
                        self.labelVar.set("{0}/{1}".format(temp, length))
                else:
                    self.success += 1
                temp += 1
                self.labelVar.set("{0}/{1}".format(self.count_file + 1, length))
                self.count_file += 1
            # 删除生成的片头片尾文件
            if os.path.exists(new_video_addr + '/' + title_addr.split('/')[-1]):
                os.remove(new_video_addr + '/' + title_addr.split('/')[-1])
            # 删除生成的msg文件
            if os.path.exists(new_video_addr + '/temp/1.mpg'):
                os.remove(new_video_addr + '/temp/1.mpg')
            if os.path.exists(new_video_addr + '/temp/2.mpg'):
                os.remove(new_video_addr + '/temp/2.mpg')
            self.flag = False
            if self.count_file == length:
                # 总耗时
                total_time = end_time - start
                h1 = total_time / 3600
                m1 = total_time % 3600 / 60
                s1 = total_time % 3600 % 60
                str2 = "{0}:{1}:{2}".format(int(h1), int(m1), int(s1))
                self.Text8Var.set(
                    "操作完成！一共处理{0}个文件，已成功处理{1}个，处理失败{2}个，总耗时{3}".format(length, self.success, self.fail, str2))
                f.write(self.Text6.get('1.0', 'end-1c'))
                f.write(self.Text8Var.get())
                f.close()
        else:  # 不需要添加片头片尾文件
            for video_path in video_list:
                self.labelVar.set("{0}/{1}".format(self.count_file + 1, length))
                if not int(combo_text) or temp % (int(combo_text) + 1) == 0:
                    try:
                        start_time = time.time()
                        path = video_path.replace('\\', '/').rstrip('\n')
                        insert_text = path.split('/')[-1] + '\t正在处理\t开始时间' + datetime.datetime.now().strftime(
                            '%Y-%m-%d %H:%M:%S') + '\n'
                        self.Text6.insert(END, insert_text)
                        # 添加水印
                        WaterMark(True, path, watermark_addr, new_video_addr, watermark_position, watermark_rate,
                                  self.Text6, self.Text8Var).add_mark()
                        end_time = time.time()
                        self.success += 1
                    except Exception as e:
                        end_time = time.time()
                        self.Text6.insert(END, e)
                        self.Text6.insert(END, '\n')
                        self.fail += 1
                        continue
                    finally:
                        # 当次操作耗时
                        per_time = end_time - start_time
                        h = per_time / 3600
                        m = per_time % 3600 / 60
                        s = per_time % 3600 % 60
                        str1 = "{0}:{1}:{2}".format(int(h), int(m), int(s))
                        insert_text = path.split('/')[-1] + '\t处理完毕' + datetime.datetime.now().strftime(
                            '%Y-%m-%d %H:%M:%S') + '\n当次总共耗时' + str1 + '\n'
                        self.Text6.insert(END, insert_text)
                else:
                    self.success += 1
                temp += 1
                self.labelVar.set("{0}/{1}".format(self.count_file + 1, length))
                self.count_file += 1
            self.flag = False
            if self.count_file == length:
                # 总耗时
                total_time = end_time - start
                h1 = total_time / 3600
                m1 = total_time % 3600 / 60
                s1 = total_time % 3600 % 60
                str2 = "{0}:{1}:{2}".format(int(h1), int(m1), int(s1))
                self.Text8Var.set("一共处理{0}个文件，已成功处理{1}个，处理失败{2}个，总耗时{3}".format(length, self.success, self.fail, str2))
                f.write(self.Text6.get('1.0', 'end-1c'))
                f.write(self.Text8Var.get())
                f.close()

    # 合并视频
    def concat_video(self, title_file, title_name, file_name, new_video_addr):
        # 判断视频帧率是否大于40fps，大于此帧率更改帧率28
        fps_title = cv2.VideoCapture(new_video_addr + '/' + title_name).get(5)
        fps_file = cv2.VideoCapture(new_video_addr + '/' + file_name).get(5)
        if fps_title > 40 or fps_title < 20:
            self.Text8Var.set('更改片头片尾文件' + title_name + '帧率,如果文件过大，将耗费相当长时间')
            self.Text6.insert(END, '更改片头片尾文件' + title_name + '帧率\t' + datetime.datetime.now().strftime(
                '%Y-%m-%d %H:%M:%S') + "\n")
            subprocess.call(
                'ffmpeg -i ' + new_video_addr + '/' + title_name + ' -r 25 ' + new_video_addr + '/new_title_name.mp4')
            os.remove(new_video_addr + '/' + title_name)
            os.rename(new_video_addr + '/new_title_name.mp4', new_video_addr + '/' + title_name)
        if fps_file > 40 or fps_file < 20:
            self.Text8Var.set('更改文件:' + file_name + '帧率,如果文件过大，将耗费相当长时间')
            self.Text6.insert(END, '更改文件:' + file_name + '帧率\t' + datetime.datetime.now().strftime(
                '%Y-%m-%d %H:%M:%S') + "\n")
            subprocess.call(
                'ffmpeg -i ' + new_video_addr + '/' + file_name + ' -r 25 ' + new_video_addr + '/new_file_name.mp4')
            os.remove(new_video_addr + '/' + file_name)
            os.rename(new_video_addr + '/new_file_name.mp4', new_video_addr + '/' + file_name)
        # 生成mpg文件
        if not os.path.exists(new_video_addr + '/temp/1.mpg'):
            self.Text8Var.set('片头片尾文件' + title_name + '生成mpg格式流，将耗费一部分时间')
            self.Text6.insert(END, '片头片尾文件' + title_name + '生成mpg格式流\t' + datetime.datetime.now().strftime(
                '%Y-%m-%d %H:%M:%S') + "\n")
            subprocess.call(
                'ffmpeg -i ' + new_video_addr + '/' + title_name + ' -q:v 4 -q:a 4 ' + new_video_addr + '/temp/1.mpg')
        self.Text8Var.set('文件:' + file_name + '生成mpg格式流，将耗费一部分时间')
        self.Text6.insert(END, '文件:' + file_name + '生成mpg格式流\t' + datetime.datetime.now().strftime(
            '%Y-%m-%d %H:%M:%S') + "\n")
        subprocess.call(
            'ffmpeg -y -i ' + new_video_addr + '/' + file_name + ' -q:v 4 -q:a 4 ' + new_video_addr + '/temp/2.mpg')
        # 按照要求合并生成最终的MP4文件
        self.Text8Var.set('正在合并文件...')
        self.Text6.insert(END, '合并文件\t' + datetime.datetime.now().strftime(
            '%Y-%m-%d %H:%M:%S') + "\n")
        name = file_name.split('.')[0]
        if title_file == 1:
            cmd = 'ffmpeg -y -i concat:"' + new_video_addr + '/temp/1.mpg|' + new_video_addr + '/temp/2.mpg" -c copy ' + new_video_addr + '/' + name + '.mpg'
            subprocess.call(cmd)
        elif title_file == 2:
            cmd = 'ffmpeg -y -i concat:"' + new_video_addr + '/temp/2.mpg|' + new_video_addr + '/temp/1.mpg" -c copy ' + new_video_addr + '/' + name + '.mpg'
            subprocess.call(cmd)
        os.remove(new_video_addr + '/' + file_name)

    # 根据传入的路径读取excel文件，返回保存有excel数据的列表
    def read_excel(self, time_addr):
        try:
            wb = xlrd.open_workbook(filename=time_addr)  # 打开文件
            sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
            # 获得总行数
            rows = sheet1.get_rows()
            list_excel = []
            for x in rows:
                if x and x[0].value != '' and x[1].value != '' and x[2].value != '' and x[3].value != '':
                    obj_list = []
                    obj_list.append(x[0].value)
                    obj_list.append(x[1].value)
                    obj_list.append(x[2].value)
                    obj_list.append(x[3].value)
                    list_excel.append(obj_list)
            return list_excel
        except Exception as e:
            self.Text6.insert(END, e + '\n')
            tkinter.messagebox.showinfo('警告', '请确认所选截取时间文件格式正确')
            return

    # 根据传入参数分割视频
    def slice_video(self, path, new_video_addr, x):
        cap = cv2.VideoCapture(path)  # path即视频文件的路径
        file_time = cap.get(7) / cap.get(5)
        cap.release()
        tup1 = xldate_as_tuple(x[1], 0)
        tup2 = xldate_as_tuple(x[2], 0)
        start = tup1[3] * 3600 + tup1[4] * 60 + tup1[5]
        end = tup2[3] * 3600 + tup2[4] * 60 + tup2[5]
        t_start = "{0}:{1}:{2}".format(tup1[3], tup1[4], tup1[5])
        t_end = "{0}:{1}:{2}".format(math.floor((end - start) / 3600), math.floor((end - start) % 3600 / 60),
                                     math.floor((end - start) % 3600 % 60))
        if start < file_time:
            if start < 0:
                t_start = "00:00:00"
            if end > file_time:
                t_end = "{0}:{1}:{2}".format(math.floor((file_time - start) / 3600),
                                             math.floor((file_time - start) % 3600 / 60),
                                             math.floor((file_time - start) % 3600 % 60))
            # 生成新文件名
            new_name = new_video_addr + '/temp/' + x[3]
            # 执行切割

            self.Text8Var.set('正在切割文件:{0}-->{1}'.format(path.split('/')[-1], x[3]))

            self.Text6.insert(END, '正在切割文件:{0}-->{1}'.format(path.split('/')[-1],
                                                             x[3]) + '\t' + datetime.datetime.now().strftime(
                '%Y-%m-%d %H:%M:%S') + "\n")
            subprocess.call(
                "ffmpeg -y -ss " + t_start + " -t " + t_end + " -i " + path + " -vcodec copy -acodec copy " + new_name)
            return new_name
        return ''


if __name__ == "__main__":
    top = Tk()
    Application(top).mainloop()
