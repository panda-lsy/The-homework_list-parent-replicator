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
targetDir=r'/'
listDir=os.getcwd()
imageDir=r'older_versions/images/'

listDir=listDir+'/'

iffilemove='true'                                       #是否移动旧作业

mylink="https://github.com/panda-lsy/The-homework_list-parent-replicator"

waiting="true"                                          #防止直接进行复制进程

moveFilelist=[]                                         #被移动文件的集合

#print(listDir)
def day_check(weekday, weekdayDisplay):                 #检测今天日期，防误删
    global iffilemove
    global waiting
    while True:
        if waiting =="true":
            if weekday=="Sat" or weekday=="Sun":
                    choice1=g.indexbox("今天是"+weekdayDisplay+",如果进行作业复制将会自动将旧作业移动到older_versions文件夹",title="复制作业副本ver3.0:防误触",choices=("无视,继续生成","生成,但是不移动旧作业","我点错了,退出"))
                    print(choice1)
                    if choice1==0:
                        iffilemove='true'
                        waiting="false"
                    
                    if choice1==1:
                        iffilemove='false'
                        waiting="false"
                    
                    if choice1==2:
                        exit()
            
            else:
                waiting = "false"
        if waiting=="false":
            break

Checkdate = day_check(weekday, weekdayDisplay)

#移动

def move_old_file(datetime, moveDir, listDir, moveFilelist):
    global moveFile
    datanames = os.listdir(listDir)

    for dataname in datanames:
        if os.path.splitext(dataname)[1] == '.docx':#目录下包含.docx的文件
        #print(dataname)
            listFile = os.path.join(listDir,dataname)   #把文件夹名和文件名称链接起来
            moveFile = os.path.join(moveDir,dataname)
        #print(listFile)
            if listFile == (listDir+str(datetime)+".docx"): #要求名字是今日作业母本
                title = g.msgbox(msg="已经存在今日作业母本(错误代码:ERROR#114514)",title="复制作业副本ver3.0:复制失败",ok_button="OK")
                time.sleep(5)
                exit()
            else:
                if iffilemove == 'true':
                    shutil.move(listFile,moveFile)
                    moveFilelist=moveFilelist.append(moveFile)
                    

moveFiles = move_old_file(datetime, moveDir, listDir, moveFilelist)

#复制
def copy_new_file(datetime, sourceDir, targetDir):
    for files in os.listdir(sourceDir):
        sourceFile = os.path.join(sourceDir,files)   #把文件夹名和文件名称链接起来
        if os.path.isfile(sourceFile) and sourceFile.find('.docx')>0: #要求是文件且后缀是docx
            shutil.copy(sourceDir+"master.docx",targetDir)
            os.rename(targetDir+"master.docx",str(datetime)+".docx")

copyFile = copy_new_file(datetime, sourceDir, targetDir)



def Last(imageDir, moveFile):
    if iffilemove == 'true':
        moveFilelist_str=''.join(moveFilelist)
        movement = "移动了"+moveFilelist_str+''
    else:
        movement = '\n'
    choice2=g.ccbox("作者:LSY\n未经作者授权随意转载\n"+movement+"开源是一种美德。",image=imageDir+"successful.png",title="复制作业副本ver3.0:复制成功",choices=("好的","查看使用说明"))

    if choice2 == True:
        exit()
    if choice2 == False:
        choice3 = g.indexbox(msg="使用说明：双击本应用即可立刻复制作业母本并设置好今日日期\n待更新内容:\n最终目标:通过Python的QQAPI来进行各个学科群作业关键字自动更新作业。",title="复制作业副本ver3.0:使用说明",choices=("好的","打开本项目的Github网址(你可能需要科学上网)"))
        if choice3 == 0:
            exit()
        if choice3 == 1:
            webbrowser.open(mylink, new=0, autoraise=True)
            

    
def main():
    
    Checkdate
    

    for listfiles in os.listdir(listDir):
        print(listfiles)
        if listfiles.find('.docx')>0 and iffilemove=="true":
            moveFiles()                      #如果docx数量＞0,则移动旧文件
        else:
            iffilemove="false"
    
    if iffilemove == 'true':
    
        LastDisplay = Last(imageDir, moveFile)
    
    if iffilemove == 'false':
        moveFile="\n"
        LastDisplay = Last(imageDir,moveFile)
            
    copyFile()
    LastDisplay()
    exit()
    
    
if __name__ == '__main__':
    main()

    
