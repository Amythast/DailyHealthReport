import pyautogui
import time
import xlrd
import pyperclip
import random
from apscheduler.schedulers.blocking import BlockingScheduler
from datetime import datetime


#定义鼠标事件

#pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159

def mouseClick(clickTimes,lOrR,img,reTry):#retry=1不重试,-1重试，>1次数重试
    if reTry == 1:
        while True:
            #print(pyautogui.size()) mac屏幕1440*900
            location=pyautogui.locateCenterOnScreen(img)#可加confidence置信度参数
            #找到对应点击位置
            if location is not None :
                print(str(location.x/2)+" "+str(location.y/2-2))
                pyautogui.click(location.x/2,location.y/2-2,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                break
            print("未找到匹配图片,0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                print("重复")
                i += 1
            time.sleep(0.1)

def moveToImg(img,reTry):
    if reTry == 1:
        while True:
            #print(pyautogui.size()) mac屏幕1440*900
            location=pyautogui.locateCenterOnScreen(img)#可加confidence置信度参数
            #找到对应点击位置
            if location is not None :
                print(str(location.x/2)+" "+str(location.y/2-2))
                pyautogui.moveTo(location.x/2,location.y/2,duration=1)  
                break
            print("未找到匹配图片,0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.moveTo(location.x/2,location.y/2,duration=1)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.moveTo(location.x/2,location.y/2,duration=1)
                print("重复")
                i += 1
            time.sleep(0.1)




# 数据检查
# cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
# ctype     空：0
#           字符串：1
#           数字：2
#           日期：3
#           布尔：4
#           error：5
def dataCheck(sheet1):
    checkCmd = True
    #行数检查
    if sheet1.nrows<2:
        print("没数据啊哥")
        checkCmd = False
    #每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0 
        and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0 and cmdType.value != 7.0):
            print('第',i+1,"行,第1列数据有毛病")
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value ==1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd = False
        # 输入类型，内容不能为空
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd = False
        #鼠标移动事件  事件内容必须为字符串类型 
        if cmdType.value == 7.0:
            if cmdValue.ctype != 1:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd = False
        i += 1
    return checkCmd

#任务
def mainWork(sheet1):
    i = 1
    while i < sheet1.nrows:
        #取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        #1代表单机左键
        if cmdType.value == 1.0:
            #取图片名称
            img = sheet1.row(i)[1].value
            reTry = 1#retry是重复次数
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"left",img,reTry)
            print("单击左键",img)
        #2代表双击左键
        elif cmdType.value == 2.0:
            #取图片名称
            img = sheet1.row(i)[1].value
            #取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2,"left",img,reTry)
            print("双击左键",img)
        #3代表右键
        elif cmdType.value == 3.0:
            #取图片名称
            img = sheet1.row(i)[1].value
            #取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"right",img,reTry)
            print("右键",img) 
        #4代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value    
            pyperclip.copy(inputValue)
            pyautogui.keyDown('command')
            pyautogui.press('v')
            pyautogui.keyUp('command')
            #pyautogui.hotkey('command','v')
            pyautogui.hotkey('return')
            time.sleep(0.5)
            print("输入:",inputValue)                                        
        #5代表等待
        elif cmdType.value == 5.0:
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("等待",waitTime,"秒")
        #6代表滚轮
        elif cmdType.value == 6.0:
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动",int(scroll),"距离")    
        #7代表光标移动    
        elif cmdType.value == 7.0:
             #取图片名称
            img = sheet1.row(i)[1].value
            reTry = 1#retry是重复次数
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            moveToImg(img,reTry)
            print("光标移动到",img)        
        i += 1

def healthReport() :
    file = '/Users/feifeixia/Desktop/DailyHealth/cmd.xls'
    #打开文件
    wb = xlrd.open_workbook(filename=file)
    #通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    print('欢迎使用DailyHealth')
    #数据检查
    checkCmd = dataCheck(sheet1)
    if checkCmd:
        ##key=input('选择功能: 1.做一次 2.循环 \n')
        key = '1'
        if key=='1':
            #循环拿出每一行指令
            mainWork(sheet1)
        elif key=='2':
            while True:
                mainWork(sheet1)
                time.sleep(0.1)
                print("等待0.1秒")    
    else:
        print('输入有误或者已经退出!')

if __name__ == '__main__':
    healthReport()
    

