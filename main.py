# -*- coding: utf-8 -*- 
import os
import shutil
import easygui as g
import time
import webbrowser                                       #自动打开源码网站


datetime=time.strftime("%b %d", time.localtime()) 
weekday=time.strftime("%a",time.localtime())            #检测今天是星期几
weekdayDisplay=time.strftime("%A",time.localtime()) 

#print(weekday)

moveDir=r'older_versions/'
sourceDir=r'older_versions/master/'
sourceDirtwo=r'older_versions/'
targetDir=r'/'
listDir=os.getcwd()
listDirtwo=os.path.join(listDir,"older_versions")
imageDir=r'older_versions/images/'

listDir=listDir+'/'

iffilemove='true'                                       #是否移动旧作业

mylink="https://github.com/panda-lsy/The-homework_list-parent-replicator"

waiting="true"                                          #防止直接进行复制进程

moveFilelist=[]                                         #被移动文件的集合

datanamesfinally=[]

copyfile='true'

#print(listDir)
def day_check(weekday, weekdayDisplay):                 #检测今天日期，防误删
    global iffilemove
    global waiting
    while True:
        if waiting =="true":
            if weekday=="Sat" or weekday=="Sun":
                    choice1=g.indexbox("今天是"+weekdayDisplay+",如果进行作业复制将会自动将旧作业移动到older_versions文件夹",title="复制作业副本ver6.0:防误触",choices=("无视,继续生成","生成,但是不移动旧作业","回档之前的作业清单","我点错了,退出"))
                    print(choice1)
                    if choice1==0:
                        iffilemove='true'
                        waiting="false"

                    if choice1==1:
                        iffilemove='false'
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

def backup():              #回档
    global sourceDirtwo
    global datanamesfinally
    datanamestwo = os.listdir(listDirtwo)
    for datanametwo in datanamestwo:
        if os.path.splitext(datanametwo)[1] == '.docx':#目录下包含.docx的文件
            datanamesfinally.append(datanametwo)   
    back=g.choicebox("请选择需要回档的文件", "文件回档", datanamesfinally)
    sourceDirtwo = os.path.join(sourceDirtwo,back)
    print(sourceDirtwo)
    shutil.copy(sourceDirtwo,listDir)
    choice2=g.indexbox("作者:LSY\n未经作者授权随意转载\n开源是一种美德。",image=imageDir+"successful.png",title="复制作业副本ver6.0:回档成功",choices=("好的","查看使用说明"))
    if choice2 == 0:
        os._exit(0)
    if choice2 == 1:
        choice3 = g.indexbox(msg="使用说明：双击本应用即可立刻复制作业母本并设置好今日日期\n复制完毕后你可以再次开启应用回档之前的作业清单。\n待更新内容:\n最终目标:通过Python的QQAPI来进行各个学科群作业关键字自动更新作业。",title="复制作业副本ver6.0:使用说明",choices=("好的","打开本项目的Github网址(你可能需要科学上网)"))
        if choice3 == 0:
            os._exit(0)
        if choice3 == 1:
            webbrowser.open(mylink, new=0, autoraise=True)
            os._exit(0)

def move_old_file(datetime, moveDir, listDir, moveFilelist):            #移动
    global moveFile
    global copyfile
    datanames = os.listdir(listDir)

    has_copyfile = False
    has_txtfile = False

    for dataname in datanames:
        print(dataname)
        if os.path.splitext(dataname)[1] == '.docx':#目录下包含.docx的文件
            has_txtfile = True
        #print(dataname)
            listFile = os.path.join(listDir,dataname)   #把文件夹名和文件名称链接起来
            moveFile = os.path.join(moveDir,dataname)
        #print(listFile)
            if listFile == (listDir+str(datetime)+".docx"): #要求名字是今日作业母本,则无法复制
                has_copyfile = True
    if has_txtfile == False and has_copyfile == False:
        for files in os.listdir(sourceDir):
            sourceFile = os.path.join(sourceDir,files)   #把文件夹名和文件名称链接起来
            if os.path.isfile(sourceFile) and sourceFile.find('.docx')>0 and copyfile=='true': #要求是文件且后缀是docx
                shutil.copy(sourceDir+"master.docx",targetDir)
                os.rename(targetDir+"master.docx",str(datetime)+".docx")
        movement = '\n'
        choice2=g.indexbox("作者:LSY\n未经作者授权随意转载\n"+movement+"开源是一种美德。",image=imageDir+"successful.png",title="复制作业副本ver6.0:复制成功",choices=("好的","查看使用说明","回档之前作业"))

        if choice2 == 0:
            os._exit(0)
        if choice2 == 1:
            choice3 = g.indexbox(msg="使用说明：双击本应用即可立刻复制作业母本并设置好今日日期\n待更新内容:\n最终目标:通过Python的QQAPI来进行各个学科群作业关键字自动更新作业。",title="复制作业副本ver6.0:使用说明",choices=("好的","打开本项目的Github网址(你可能需要科学上网)","回档作业"))
            if choice3 == 0:
                os._exit(0)
            if choice3 == 1:
                webbrowser.open(mylink, new=0, autoraise=True)
            if choice3 == 2:
                backup()
        if choice2 == 2:
            backup()
        os._exit(0)

    #print(has_copyfile)
    if has_copyfile == True:
            copyfile='false'
            title = g.ccbox("已经存在今日作业母本(错误代码:ERROR#114514)",title="复制作业副本ver6.0:复制失败",choices=("好的","回档之前作业"))
            if title == True:
                os._exit(0)
            if title == False:
                backup()
                
    else:
        for dataname in datanames:
            if os.path.splitext(dataname)[1] == '.docx':#目录下包含.docx的文件
            #print(dataname)
                listFile = os.path.join(listDir,dataname)   #把文件夹名和文件名称链接起来
                moveFile = os.path.join(moveDir,dataname)
            if iffilemove == 'true':
                shutil.move(listFile,moveFile)
                moveFilelist=moveFilelist.append(dataname)
                return moveFilelist


moveFiles = move_old_file(datetime, moveDir, listDir, moveFilelist)

def Last(imageDir, moveFile):                           
    if iffilemove == 'true':
        moveFilelist_str=''.join(moveFilelist)
        movement = "移动了"+moveFilelist_str+'\n'
    else:
        movement = '\n'
    choice2=g.indexbox("作者:LSY\n未经作者授权随意转载\n"+movement+"开源是一种美德。",image=imageDir+"successful.png",title="复制作业副本ver6.0:复制成功",choices=("好的","查看使用说明","回档之前作业"))

    if choice2 == 0:
        os._exit(0)
    if choice2 == 1:
        choice3 = g.indexbox(msg="使用说明：双击本应用即可立刻复制作业母本并设置好今日日期\n待更新内容:\n最终目标:通过Python的QQAPI来进行各个学科群作业关键字自动更新作业。",title="复制作业副本ver6.0:使用说明",choices=("好的","打开本项目的Github网址(你可能需要科学上网)","回档作业"))
        if choice3 == 0:
            os._exit(0)
        if choice3 == 1:
            webbrowser.open(mylink, new=0, autoraise=True)
        if choice3 == 2:
            backup()
    if choice2 == 2:
        backup()

#复制
def copy_new_file(datetime, sourceDir, targetDir):
    for files in os.listdir(sourceDir):
        sourceFile = os.path.join(sourceDir,files)   #把文件夹名和文件名称链接起来
        if os.path.isfile(sourceFile) and sourceFile.find('.docx')>0 and copyfile=='true': #要求是文件且后缀是docx
            shutil.copy(sourceDir+"master.docx",targetDir)
            os.rename(targetDir+"master.docx",str(datetime)+".docx")

copyFile = copy_new_file(datetime, sourceDir, targetDir)

def main():

    Checkdate                                #检查日期


    for listfiles in os.listdir(listDir):
        #print(listfiles)
        if listfiles.find('.docx')>0 and iffilemove=="true":
            moveFiles()                      #如果docx数量＞0,则移动旧文件
        else:
            iffilemove="false"

    if iffilemove == 'true':

        LastDisplay = Last(imageDir, moveFile)

    if iffilemove == 'false':
        moveFile="\n"
        copyFile="\n"                       #修复每两次打开报错的BUG
        LastDisplay = Last(imageDir,moveFile)

    copyFile
    LastDisplay
    os._exit(0)


if __name__ == '__main__':
    main()
    os.system('pause')
