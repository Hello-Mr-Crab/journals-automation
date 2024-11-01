import ctypes
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
ES_SYSTEM_REQUIRED=0x00000001
ES_DISPLAY_REQUIRED=0x00000002
ES_AWAYMODE_REQUIRED=0x00000040 
ES_CONTINUOUS=0x80000000

def set_Volume_to_100():
    '''将windows系统音量设置为100,语音播报中会用到'''
# 判断是否静音，mute为1代表是静音，为0代表不是静音
    mute=volume.GetMute()
    if mute==1:
        volume.SetMute(False,None)
    volume.SetMasterVolumeLevel(0.0, None)

def open_Listening_mode():
    '''调用此函数后,只要不关机,屏幕保持常亮'''
    #调用win32内核dll库中的SetThreadExecutionState函数这个函数顾名思义就是启动一个轻量级线程，启动后不阻塞的话会一直运行,因此电脑屏幕会常亮。
    #详情见这个网址'https://learn.microsoft.com/zh-cn/windows/win32/api/winbase/nf-winbase-setthreadexecutionstate'
    #这里实现的是屏幕常亮，让ES_CONTINUOUS | ES_DISPLAY_REQUIRED这两个
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_DISPLAY_REQUIRED)
    
def close_Listening_mode():
    '''与open_Listening_mode函数功能相反,用来关闭屏幕常亮\n
    需要与open_Listening_mode函数成对使用,单独使用无意义'''
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

