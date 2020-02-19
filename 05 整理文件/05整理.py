# --** coding="UTF-8" **--
# 
import os
import re
import sys



#要修改什么类型的文件


fileType = '.txt'
#前缀
front = ''
#后缀
back = ''
#是否保留原文件名字 True False
old = True


currentpath = os.getcwd()
fileList = os.listdir(currentpath)

# 名称变量
num = 1
# 遍历文件夹中所有文件
for fileName in fileList:
    prefix = os.path.splitext(fileName)[0]
    fix = os.path.splitext(fileName)[1]
    # 文件重新命名
    if fix ==fileType and os.path.isfile(os.path.join(currentpath,fileName)) :
        if (num < 10):
            num_str = '0' + str(num)
        else:
            num_str = str(num)
        if old == False:
            newName = front + num_str + back + fix
        else:
            newName = front + num_str + back +fileName
        os.rename(fileName, newName)
        # 改变编号，继续下一项
        num = num + 1
print("***************************************")

# 刷新
sys.stdin.flush()
print("修改后：" + str(os.listdir(currentpath)))
# 输出修改后文件夹中包含的文件名称
#print("修改后：" + str(os.listdir(r"./neteasy_playlist_data3"))[1])
