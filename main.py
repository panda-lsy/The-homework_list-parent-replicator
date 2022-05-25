# -*- coding: utf-8 -*- 
import os
import shutil
import time


datetime=time.strftime("%b %d", time.localtime()) 
weekday=time.strftime("%a",time.localtime())            #检测今天是星期几
weekdayDisplay=time.strftime("%A",time.localtime()) 

#print(weekday)

moveDir=r'older_versions/'
sourceDir=r'older_versions/master/'
targetDir=r'/'
listDir=os.getcwd()
imageDir=r'older_versions/images/'

listDir=listDir+'/'

iffilemove='true'                                       #是否移动旧作业

mylink="https://github.com/panda-lsy/The-homework_list-parent-replicator"

waiting="true"                                          #防止直接进行复制进程

moveFilelist=[]                                         #被移动文件的集合

copyfile='true'

#print(listDir)
def day_check(weekday, weekdayDisplay):                 #检测今天日期，防误删
    global iffilemove
    global waiting
    while True:
        if waiting =="true":
            if weekday=="Sat" or weekday=="Sun":
                    print("今天是"+weekdayDisplay+",如果进行作业复制将会自动将旧作业移动到older_versions文件夹")
                    print("A.无视,继续生成")
                    print("B.生成,但是不移动旧作业")
                    print("C.我点错了,退出")
                    while True:
                        choice1=input("请输入你的选择\n")
                        print(choice1)
                        if choice1=='A':
                            iffilemove='true'
                            waiting="false"
                            break

                        if choice1=='B':
                            iffilemove='false'
                            waiting="false"
                            break

                        if choice1=='C':
                            exit()
                        
                        else:
                            print("请输入A,B,C")

            else:
                waiting = "false"
        if waiting=="false":
            break

Checkdate = day_check(weekday, weekdayDisplay)

#移动

def move_old_file(datetime, moveDir, listDir, moveFilelist):
    global moveFile
    global copyfile
    datanames = os.listdir(listDir)

    for dataname in datanames:
        if os.path.splitext(dataname)[1] == '.docx':#目录下包含.docx的文件
        #print(dataname)
            listFile = os.path.join(listDir,dataname)   #把文件夹名和文件名称链接起来
            moveFile = os.path.join(moveDir,dataname)
        #print(listFile)
            if listFile == (listDir+str(datetime)+".docx"): #要求名字是今日作业母本
                copyfile='false'
                print("已经存在今日作业母本(错误代码:ERROR#114514)")
                time.sleep(5)
                exit()
            else:
                if iffilemove == 'true':
                    shutil.move(listFile,moveFile)
                    moveFilelist=moveFilelist.append(dataname)
                    return moveFilelist


moveFiles = move_old_file(datetime, moveDir, listDir, moveFilelist)

#复制
def copy_new_file(datetime, sourceDir, targetDir):
    for files in os.listdir(sourceDir):
        sourceFile = os.path.join(sourceDir,files)   #把文件夹名和文件名称链接起来
        if os.path.isfile(sourceFile) and sourceFile.find('.docx')>0 and copyfile=='true': #要求是文件且后缀是docx
            shutil.copy(sourceDir+"master.docx",targetDir)
            os.rename(targetDir+"master.docx",str(datetime)+".docx")

copyFile = copy_new_file(datetime, sourceDir, targetDir)



def Last(imageDir, moveFile):
    if iffilemove == 'true':
        moveFilelist_str=''.join(moveFilelist)
        movement = "移动了"+moveFilelist_str+'\n'
    else:
        movement = '\n'
    print('''
    ┏☆━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━☆┓\n
    ┃☆^ǒ^*☆*^ǒ^*★*^ǒ^*☆*^ǒ^*★*^ǒ^**┃\n
    ┃作者:LSY   未经作者授权随意转载_____┃\n
    ┃☆^ǒ^*★*^ǒ^*☆*^ǒ^*★*^ǒ^*☆*^ǒ^**┃\n
    ┃^ǒ^*★*^ǒ^*☆*开源是一种美德^ǒ^*★*^┃\n
    ┗☆━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━☆┛,\n''',
    movement,'''\n
    ''')

    a=input("按下右上角'x'退出,回车查看使用说明>>>>")

    print("使用说明：双击本应用即可立刻复制作业母本并设置好今日日期\n待更新内容:\n最终目标:通过Python的QQAPI来进行各个学科群作业关键字自动更新作业。\n项目地址:",mylink,"\n")

    b=input("按下任意键继续>>>")



def main():

    Checkdate


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
        LastDisplay = Last(imageDir,moveFile)

    copyFile
    LastDisplay
    exit()


if __name__ == '__main__':
    main()
    os.system('pause')
