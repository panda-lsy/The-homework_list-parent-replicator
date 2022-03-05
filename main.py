import os
import shutil
import time
import easygui as g


datetime=time.strftime("%b %d", time.localtime()) 

sourceDir=r'older_versions/master/'
targetDir=r'/'


for files in os.listdir(sourceDir):
    sourceFile = os.path.join(sourceDir,files)   #把文件夹名和文件名称链接起来
    targetFile = os.path.join(targetDir,files)
    if os.path.isfile(sourceFile) and sourceFile.find('.docx')>0: #要求是文件且后缀是docx
        shutil.copy(sourceDir+"master.docx",targetDir)


os.rename(targetDir+"master.docx",str(datetime)+".docx")

title = g.msgbox(msg="作者:LSY\n未经作者授权随意转载\n开源是一种美德。",title="复制成功",ok_button="我tm谢谢你")