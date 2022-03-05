import os
import shutil
import time
import easygui as g
import time


datetime=time.strftime("%b %d", time.localtime()) 

moveDir=r'older_versions/'
sourceDir=r'older_versions/master/'
targetDir=r'/'
listDir=os.getcwd()
imageDir=r'older_versions/images/'

listDir=listDir+'/'

#print(listDir)

datanames = os.listdir(listDir)

for dataname in datanames:
    if os.path.splitext(dataname)[1] == '.docx':#目录下包含.docx的文件
        #print(dataname)
        listFile = os.path.join(listDir,dataname)   #把文件夹名和文件名称链接起来
        moveFile = os.path.join(moveDir,dataname)

#移动
#print(listFile)
if listFile == (listDir+str(datetime)+".docx"): #要求名字是今日作业母本
    title = g.msgbox(msg="已经存在今日作业母本(错误代码:ERROR#114514)",title="复制作业副本ver2.0:复制失败",ok_button="OK")
    time.sleep(5)
    exit()
else:
        shutil.move(listFile,moveFile)

#复制
for files in os.listdir(sourceDir):
    sourceFile = os.path.join(sourceDir,files)   #把文件夹名和文件名称链接起来
    targetFile = os.path.join(targetDir,files)
    if os.path.isfile(sourceFile) and sourceFile.find('.docx')>0: #要求是文件且后缀是docx
        shutil.copy(sourceDir+"master.docx",targetDir)
        os.rename(targetDir+"master.docx",str(datetime)+".docx")



a=g.ccbox("作者:LSY\n未经作者授权随意转载\n移动了"+moveFile+"开源是一种美德。",image=imageDir+"successful.png",title="复制作业副本ver2.0:复制成功",choices=("好的","查看使用说明"))

if a==True:
    exit()
if a==False:
    title = g.msgbox(msg="使用说明：双击本应用即可立刻复制作业母本并设置好今日日期\n待更新内容:\n1.自动将源码Pull到Github上\n2.自动安装软件(win7/win10:exe)\n最终目标:通过Python的QQAPI来进行各个学科群作业关键字自动更新作业。",title="复制作业副本ver2.0:使用说明",ok_button="好的")
