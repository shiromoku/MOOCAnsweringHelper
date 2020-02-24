# -*- coding: utf-8 -*-
# __Author__: Shiromoku
# __Email__ : shiromoku@outlook.com

import os
import sys
import time
import pyperclip
import webbrowser
import win32com.client as win32
from urllib import parse
# import keyboard


exam_bank = 'bank.doc'                          # 题库文件名

def search(str):                                # 发起搜索(使用bing搜索)
    keyWord = {'q':str}
    data = parse.urlencode(keyWord)
    webbrowser.open('https://www.bing.com/search?'+data)

def really_main():
    history_value = ""
    tmp_value=""
    word = win32.gencache.EnsureDispatch('word.Application')    #打开word
    word.Visible = True                                         #word窗口可见
    word.DisplayAlerts = False                                  #关闭警告信息
    doc = word.Documents.Open(os.getcwd() + '/' + exam_bank)    #打开题库
    seletion = word.Selection
    tmp_value = pyperclip.paste()                               #使用剪切板原本存在的内容初始化内容记录
    history_value = tmp_value
    try:
        while True:
            tmp_value = pyperclip.paste()               # 读取剪切板复制的内容
            if tmp_value != history_value:                   #如果检测到剪切板改动，进入文本的搜索
                history_value = tmp_value
                seletion.SetRange(0, 0)
                if seletion.Find.Execute(FindText=history_value, Forward=True) :
                    print("success")                        #在题库中查找成功
                else :
                    print("fail")
                    search(history_value)            #在题库中查找失败
                # print(recent_value)
                time.sleep(1.0)
            
    except KeyboardInterrupt:  # 使用ctrl+c以退出程序
    # except BaseException :
        # print("end1")
        pass
    # print("end")
    print("感谢使用")
    doc.Close()                # 关闭打开的文档
    word.Quit()

if __name__ == "__main__":
    # keyboard.add_hotkey('ctrl+q', exit(1))
    really_main()
