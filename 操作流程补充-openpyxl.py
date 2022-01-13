import openpyxl as openpyxl
import os as os

#文件扫描
def filescanner(path):
    try:
        for dirpath,dirnames,filenames in os.walk(path):
            return filenames
    except Exception as e:
        print("路径报错：",e)
#目录下文件绝对路径
def getfilepaths(dirpath,filenames):
    filepaths=[]
    for filename in filenames:
        filepaths.append(dirpath+'\\'+filename)
    return filepaths

if __name__ == '__main__':
    #文件操作目录
    workdirpath=r'C:\Users\97759\Desktop\project\20220111'
    filenames=filescanner(workdirpath)
    filepaths=getfilepaths(workdirpath,filenames)
    for filepath in filepaths:
        print(filepath)
