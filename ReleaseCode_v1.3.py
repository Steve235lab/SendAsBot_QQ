# -*- coding: utf-8 -*-
"""
Open source edition of Send As Bots for QQ

Created on Sat Jul  4 02:38:54 2020

@author: Steve.D.Jobs

Copyright (c) 2020 Steve D. J..All rights reserved.
"""

"""
Based on released verson 1.3.
"""

"""
声明

此程序最初由作者SteveDJobs发布于知乎

源码仅供学习交流使用，严禁用于商业用途。

业余练手之作，如有不足，欢迎改进指教。
"""


import win32gui
import win32con
import win32clipboard as w
import win32api
import time
from pandas import read_excel 


def setText(massage):   #重设剪贴板文本
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, massage)
    w.CloseClipboard()


def send_qq(massage):
    setText(massage)
    
    hwnd_title = dict()    
    def get_all_hwnd(hwnd,mouse):
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
            hwnd_title.update({hwnd:win32gui.GetWindowText(hwnd)})
    
    win32gui.EnumWindows(get_all_hwnd, 0)
    for h,t in hwnd_title.items():
        if t != "":                    
            hwnd = win32gui.FindWindow('TXGuiFoundation', t)    # 获取qq窗口句柄    
            if hwnd != 0:
                #win32gui.SendMessage(hwnd, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)    #虽然可以还原最小化的会话窗口，但经过测试发现并不能解决还原后不发送消息的问题。
                win32gui.ShowWindow(hwnd,win32con.SW_SHOW)
                time.sleep(0.1)
                win32gui.SetForegroundWindow(hwnd)
                win32gui.SetActiveWindow(hwnd)
                time.sleep(0.1)
                win32gui.SendMessage(hwnd,770, 0, 0)    # 将剪贴板文本发送到QQ窗体
                win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)  #模拟按下回车键
                win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)  #模拟松开回车键


def get_contects(mod):    #录入联系人
    global name_list
    name_list = []
    if mod == 'A':
        print("建议先创建一个txt文档将要发送信息的好友名称录入后按提示粘贴入命令窗口内，以免出现意外情况浪费输入时间。\n")
        print('输入要发送消息的好友名称（不必精确，但要确保键入联系人搜索框中时，要发送信息的联系人处在第一个）\n')
        i = ''
        while i != 'exit':
            i = input('输入一个联系人后回车以继续输入，不再输入请键入“exit”并回车\n')
            if i != 'exit':
                name_list.append(i)
    elif mod == 'B':
        flag = ''
        while flag != 'Y':
            flag = input('将联系人名称复制到"contects.xlsx"中"名称"列后，键入"Y"并回车以开始读取。\n')
        name_list = read_excel('contects.xlsx',index_col = None,colNames = 0).iloc[:,0].values
        if name_list[0] != '':
            print("已成功读取联系人，联系人名称如下:\n")
            for j in name_list:
                print(j)
        else:
            print("未成功读取联系人！")
        time.sleep(3)
    elif mod == 'C':
        print("建议先创建一个txt文档将要发送信息的QQ号录入后按提示粘贴入命令窗口内，以免出现意外情况浪费输入时间。\n")
        print('输入要发送消息的QQ号\n')
        i = ''
        while i != 'exit':
            i = input('输入一个QQ号后回车以继续输入，不再输入请键入“exit”并回车\n')
            if i != 'exit':
                name_list.append(i)
    elif mod == 'D':
        flag = ''
        while flag != 'Y':
            flag = input('将QQ号每行一个复制到"contects.xlsx"中"QQ号"列后，键入"Y"并回车以开始读取。\n')
        name_list = read_excel('contects.xlsx',index_col = None,colNames = 0).iloc[:,1].values        
        if name_list[0] != '':
            print("已成功读取联系人，联系人QQ号如下:\n")
            for j in name_list:
                print(j)
        else:
            print("未成功读取联系人！")
        time.sleep(3)
       

def open_windows():     #打开QQ会话窗口
    qq_hwnd = win32gui.FindWindow(None, 'QQ') 
    print("捕捉到QQ主窗体的句柄为:"+str(qq_hwnd))
    win32gui.ShowWindow(qq_hwnd,win32con.SW_SHOW)
    print("正在打开会话窗口...\n")
    time.sleep(1)
    for i in name_list:
        massage = i
        setText(massage)
        win32gui.SetForegroundWindow(qq_hwnd)
        win32gui.SetActiveWindow(qq_hwnd)
        time.sleep(1)
        win32gui.SendMessage(qq_hwnd,770, 0, 0)
        time.sleep(1)
        win32gui.SetForegroundWindow(qq_hwnd)
        win32gui.SetActiveWindow(qq_hwnd)
        win32api.keybd_event(0x0D, win32api.MapVirtualKey(0x0D, 0), 0, 0)   
        win32api.keybd_event(0x0D, win32api.MapVirtualKey(0x0D, 0), win32con.KEYEVENTF_KEYUP, 0)


def set_time():     #定时发送
    print("是否要启用定时发送？")
    key_time = ''
    key_time = input('Y:启用！\nN:立即发送！\n')
    if key_time == 'Y':
        loc_time = time.localtime()
        print('输入发送信息的时间:\n例如:' + time.strftime("%Y-%m-%d_%H:%M:%S", loc_time) + '\n')
        send_time_str = input()
        send_time = time.mktime(time.strptime(send_time_str,"%Y-%m-%d_%H:%M:%S"))   #预计输入格式不当会导致该处报错
        loc_time = time.time()
        delta_t = send_time - loc_time
        print('将于' + str(int(delta_t/3600)) + '小时 ' + str(int((delta_t%3600)/60)) + '分钟 ' + str(int(delta_t%60)) + '秒后发送消息！\n')
        print("期间请勿关闭已打开的会话窗口！\n")
        while loc_time <= send_time:
            loc_time = time.time()
            time.sleep(1)    


#main()
print("Copyright (c) 2020 Steve D. J..All rights reserved.\n")

print("欢迎使用SendAsBots_QQ_1.3 QQ信息批量发送消息工具！\n在开始前请确认以下几点:\n1. 你已经打开并登陆了Windows QQ客户端;\n2. 打开一个会话窗口，点击其右上角 'V' 按钮，使'合并会话窗口'选项保持没有勾选的状态;\n3. 已经关闭所有QQ会话窗口。\n")
key = input('键入字符以继续:\nY:我准备好了！\n其他任意字符:退出程序。\n')

while key == 'Y':
    print("请选择录入联系人的方式:")
    mode = 'E'
    while (mode != 'A') and (mode != 'B') and (mode != 'C') and (mode != 'D'): 
        mode = input('A:手动单个输入联系人名称;\nB:将联系人名称复制到"contects.xlsx"中"名称"列后一并读取;\nC:手动单个输入QQ号;\nD:将QQ号(支持同群非好友)复制到"联系人.xlsx"中"QQ号"列后一并读取。\n')
    
    get_contects(mode)
    
    open_windows()
    
    massage = input('输入要发送的消息:\n')
    
    ct = int(input('输入发送次数(每位联系人n次):\n'))
    
    set_time()
    
    print("发送信息中...\n")
    
    while ct > 0:
        send_qq(massage)
        ct = ct - 1
        print("消息已发出！\n")
        time.sleep(1)
    
    key = input('键入字符以继续:\nY:我还要继续发！\n其他任意字符:退出程序。\n')

w.EmptyClipboard()
