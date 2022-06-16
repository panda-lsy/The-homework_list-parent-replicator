# -*- coding: utf-8 -*- 
import os
import shutil
import easygui as g
import time
import webbrowser                                       #自动打开源码网站
import win32file

datetime=time.strftime("%b %d", time.localtime()) 
weekday=time.strftime("%a",time.localtime())            #检测今天是星期几
weekdayDisplay=time.strftime("%A",time.localtime()) 

moveDir=r'older_versions/'
sourceDir=r'older_versions/master/'
sourceDirtwo=r'older_versions/'
targetDir=r'/'
listDir=os.getcwd()
listDirtwo=os.path.join(listDir,"older_versions")
imageDir=r'older_versions/images/'

listDir=listDir+'/'

mylink="https://github.com/panda-lsy/The-homework_list-parent-replicator"

waiting="true"                                          #防止直接进行复制进程

datanamesfinally=[]

def is_used(file_name):                                 #检测文件占用
    try:
        v_handle = win32file.CreateFile(file_name, win32file.GENERIC_READ, 0, None, win32file.OPEN_EXISTING, 
                                        win32file.FILE_ATTRIBUTE_NORMAL, None)
        result = bool(int(v_handle) == win32file.INVALID_HANDLE_VALUE)
        win32file.CloseHandle(v_handle)
    except Exception:
        return True
    return result

def backup():                                           #回档
    global sourceDirtwo
    global datanamesfinally
    docxcount = 0                                       #检测文件数
    success_state = ""
    text="作者:LSY\n未经作者授权随意转载\n开源是一种美德。"
    imagename = "successful"
    def successGUI():                                   #回档成功或失败界面
        choice2=g.buttonbox(text,image=imageDir+imagename+".png",title="复制作业副本ver7.1:回档"+success_state,choices=("好的","查看使用说明"))
        if choice2 == '好的':
            os._exit(0)
        if choice2 == '查看使用说明':
            choice3 = g.buttonbox(msg=
            '''使用说明：双击本应用即可立刻复制作业母本并设置好今日日期\n复制完毕后你可以再次开启应用回档之前的作业清单。
            待更新内容:
            1.设置Json与基础设置
            2.支持回档多选
            3.当文件被占用时添加解除占用功能
            4.通用功能:添加自动安装,自动打包EXE功能
            ''',title="复制作业副本ver7.1:使用说明",choices=("好的","打开本项目的Github网址(你可能需要科学上网)"))
            if choice3 == '好的':
                os._exit(0)
            if choice3 == '打开本项目的Github网址(你可能需要科学上网)':
                webbrowser.open(mylink, new=0, autoraise=True)
                os._exit(0)
    datanamestwo = os.listdir(listDirtwo)
    for datanametwo in datanamestwo:
        if os.path.splitext(datanametwo)[1] == '.docx':                         #目录下包含.docx的文件
            datanamesfinally.append(datanametwo)
            docxcount = docxcount + 1 
    if docxcount >=2:  
        back=g.choicebox("请选择需要回档的文件", "文件回档", datanamesfinally)
        if back == None:
            os._exit(0)
        sourceDirtwo = os.path.join(sourceDirtwo,back)
        if is_used(sourceDirtwo) == True:
            success_state = "失败"
            text = "移动回档库文件失败!文件"+back+"被占用,请尝试关闭Word里面你要回档的文件名."
            imagename = "error"
            successGUI()
        else:
            shutil.move(sourceDirtwo,listDir)
            success_state = "成功"
            successGUI()   
    else:
        success_state = "失败"
        text = "移动回档文件失败!你需要检测older_versions这个库文件夹里是否拥有两个以上旧文件."
        imagename = "error"
        successGUI()
        
def day_check(weekday, weekdayDisplay):                 #检测今天日期，防误删
    global iffilemove
    global waiting
    while True:
        if waiting =="true":
            if weekday=="Sat" or weekday=="Sun":
                    choice1=g.indexbox("今天是"+weekdayDisplay+",如果进行作业复制将会自动将旧作业移动到older_versions文件夹",title="复制作业副本ver7.1:防误触",choices=("无视,继续生成","生成,但是不移动旧作业","回档之前的作业清单","我点错了,退出"))
                    if choice1==0:
                        iffilemove=True
                        waiting="false"

                    if choice1==1:
                        iffilemove=False
                        waiting="false"

                    if choice1==2:
                        backup()
                    
                    if choice1==3:
                        os._exit(0)

            else:
                waiting = "false"
        if waiting=="false":
            break

Checkdate = day_check(weekday, weekdayDisplay)

def move_old_file(datetime, moveDir, listDir, weekday):            #移动和复制
    global moveFile
    
    listFiles=[]
    NewMovePaths=[]
    DocxNames=[]
    
    movement="移动了:\n"                                    #移动的文件名
    
    datanames = os.listdir(listDir)

    has_copyfile = False
    has_txtfile = False

    def GUI():
        choice2=g.buttonbox(text,image=imageDir+image_name+".png",title="复制作业副本ver7.1:"+interaction,choices=("好的","查看使用说明","回档之前作业"))
        if choice2 == '好的':
                os._exit(0)
        if choice2 == "查看使用说明":
            choice3 = g.buttonbox(msg='''使用说明：双击本应用即可立刻复制作业母本并设置好今日日期\n复制完毕后你可以再次开启应用回档之前的作业清单。
            待更新内容:
            1.设置Json与基础设置
            2.支持回档多选
            3.当文件被占用时添加解除占用功能
            4.通用功能:添加自动安装,自动打包EXE功能
            最终目标:通过Python的QQAPI来进行各个学科群作业关键字自动更新作业。''',title="复制作业副本ver7.1:使用说明",choices=("好的","打开本项目的Github网址(你可能需要科学上网)","回档作业"))
            if choice3 == "好的":
                os._exit(0)
            if choice3 == "打开本项目的Github网址(你可能需要科学上网)":
                webbrowser.open(mylink, new=0, autoraise=True)
            if choice3 == "回档作业":
                backup()
        if choice2 == "回档之前作业":
            backup()
        os._exit(0)
    #检测是否复制或者移动
    for dataname in datanames:
        if os.path.splitext(dataname)[1] == '.docx':        #目录下包含.docx的文件
            has_txtfile = True                              #存在DOCX文件
            DocxNames.append(dataname)
            listFile = os.path.join(listDir,dataname)       #把文件夹名和文件名称链接起来
            listFiles.append(listFile)                      #将文件列表保存,便于后续移动
            NewMovePath = os.path.join(moveDir,dataname)    #如果移动,它就是目标路径
            NewMovePaths.append(NewMovePath)
            if listFile == (listDir+str(datetime)+".docx"): 
                has_copyfile = True                         #检测是否存在作业母本
        
    if has_txtfile == False and has_copyfile == False:      #如果没有 进行复制 
        shutil.copy(sourceDir+"master.docx",targetDir)
        os.rename(targetDir+"master.docx",str(datetime)+".docx")
        text="作者:LSY\n未经作者授权随意转载\n开源是一种美德。"
        image_name="successful"
        interaction="复制成功"
        movement="\n"
        GUI()
            
    elif has_txtfile == True and has_copyfile == False:     #有旧文件没有新文件,那么移动+复制

        if weekday == 'Sat' or weekday == 'Sun':
            if iffilemove == True:
                for i in range(0,len(listFiles)):
                    #移动和添加移动的文件名到GUI
                    if is_used(listFiles[i]) == True:       #文件被占用,退出
                        text="移动文件失败!程序将在你进行交互后退出。文件"+DocxNames[i]+"被其它占用,请尝试关闭Word里面你要回档的文件名."
                        image_name="error"
                        interaction="移动失败"
                        GUI()
                        
                    else:                                   #没有检测文件被占用,那么移动
                        shutil.move(listFiles[i],NewMovePaths[i])  
                        movement = movement+DocxNames[i]+"\n"
                        
        else:                                               #没有防误删,直接移动
            for i in range(0,len(listFiles)):
                if is_used(listFiles[i]) == True:       #文件被占用,退出
                        text="移动文件失败!程序将在你进行交互后退出。文件"+DocxNames[i]+"被其它占用,请尝试关闭Word里面你要回档的文件名."
                        image_name="error"
                        interaction="移动失败"
                        GUI()
                else:
                    shutil.move(listFiles[i],NewMovePaths[i])
                    movement = movement+DocxNames[i] +"\n"             
        #复制文件
        shutil.copy(sourceDir+"master.docx",targetDir)
        os.rename(targetDir+"master.docx",str(datetime)+".docx")
        if weekday=="Sat" or weekday=="Sun":
            if iffilemove == False:
                movement=""
        text="作者:LSY\n未经作者授权随意转载\n"+movement+"开源是一种美德。"
        image_name="successful"
        interaction="复制成功"
        GUI()

    elif has_copyfile == True:
        text="已经存在今日作业母本(错误代码:ERROR#114514)"
        image_name="error"
        interaction="复制失败"
        GUI()
                
MoveandCopyFile = move_old_file(datetime, moveDir, listDir, weekday)

def main():

    Checkdate                                #检查日期
    MoveandCopyFile                              #复制或移动作业
    os._exit(0)

if __name__ == '__main__':
    main()
    os.system('pause')
