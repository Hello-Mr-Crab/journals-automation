import base64
import os
from PIL import Image,ImageTk
import win32com.client
from io import BytesIO
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
import tkinter as tk
from tkinter import BOTH
import ctypes
import catppuccin
import matplotlib.pyplot as plt
import matplotlib as mpl
mpl.style.use(catppuccin.PALETTE.mocha.identifier)
mpl.use('TkAgg')#tkinter中显示matplotlib图像
plt.rcParams['font.sans-serif']=['SimHei']
def Convert_digitstring(string_number:str):
    '''
    string_number:字符串类型数字\n
    这个函数用来将CNKI关键字搜索结果数量的字符串转为int或float类型,
    便于绘制图像'''
    #搜索结果是digit类型，数量未超过10000
    if string_number:
        if string_number.isdigit():
            return int(string_number)
        else:
            #CNKI关键字搜索结果数量超过10000时会使用8.9万这样的字符串来表示搜索结果
            #故需要定位到结果中万字的位置，然后将万字之前的digit类型字符串转为float并*10000来得到最后的结果
            for  index in range(len(string_number)):
                if string_number[index]=='万':
                    string_number=float(string_number[0:index])*10000
                    break
            return string_number
    else:
        return 0


def picture_Transform(picture_src_path:str,save_path:str,py_name:str):
    ''' 
    picture_src_path:图片原地址
    save_path:py文件存放地址
    pitcture_name:py文件名称
    pyintstaller打包不支持图片直接打包,所以将图片转换为base64字符串存放到py文件中后再打包。
    这个函数将图片转化为base64字符串,并将其写入为本地的py文件
    '''
    #将图片转换问base64码
    open_pic = open(picture_src_path, 'rb')
    b64str = base64.b64encode(open_pic.read())
    open_pic.close()
    #注意这边b64str一定要加上.decode() 
    write_data = 'img = "%s"' % b64str.decode()
    if py_name:
        file_path=os.path.join(save_path,py_name+'.py')
    else:
        raise('缺少py文件名称!')
    f = open(file_path, 'w+') 
    f.write(write_data)
    f.close()

def Base64_to_Image(base64_string:str,width:float,height:float):
    '''
    base64_string:图片文件的base64字符串\n
    with:图片的宽\n
    height:图片的高\n
    这个函数将base64字符串编码的图片文件转化为PhothImage类型\n
    pyintstaller打包不支持图片直接打包,打包时打包的是图片的base64编码的py文件\n
    为了显示图片需要将其转换为PhotoImage类型'''
    image_data = base64.b64decode(base64_string)
    image = Image.open(BytesIO(image_data)).resize((width,height))
    return ImageTk.PhotoImage(image)
##############################################################################

def Speak(message:str):
      '''
      message:语音播报内容。\n
      这个函数调用windows浏览器中语音播报的api'''
      speaker=win32com.client.Dispatch('SAPI.SpVoice')  
      speaker.Speak(message)  

class Drawsystem():
    '''将matplotlib的图嵌入tkinter gui界面'''
    def __init__(self,root,figure):
         self.figure=figure
         self.root=root
         canvas=FigureCanvasTkAgg(self.figure,master=self.root)
         canvas_widget=canvas.get_tk_widget()
         canvas_widget.pack(fill=BOTH,expand=True)
         canvas.draw()


 
def get_windows_scaling_factor():
    ''' ctypes调用 Windows API 函数获取缩放比例'''
    try:
        # 调用 Windows API 函数获取缩放比例
        user32 = ctypes.windll.user32
        user32.SetProcessDPIAware()
        scaling_factor = user32.GetDpiForSystem()
        scaling_factor=scaling_factor/96*100
        # 计算缩放比例
        return scaling_factor
    except Exception as e:
        print("获取缩放比例时出错:", e)

