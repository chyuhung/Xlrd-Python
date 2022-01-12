import xlrd
import xlsxwriter
import os

#文件扫描
def fileScanner(path):
    for dirPath,dirNames,fileNames in os.walk(path):
        #print(dirPath,dirNames,fileNames)
        return fileNames

#文件生成
def fileCreater(path,filenames):
    for filename in filenames:
        filePath=path+'\\'+filename
        print(filePath)
        try:
            desWorkBook=xlsxwriter.Workbook(filePath)
            desWorkBook.add_worksheet(filename)
            desWorkBook.close()
        except Exception as e:
            print('创建错误：',e)

if __name__ == '__main__':
    # 文件操作目录
    # 生成的文件目录
    desWorkBookDirPath = r'C:\Users\97759\Desktop\project\20220111\results'
    # 源文件目录
    sourceWorkBookDirPath = r'C:\Users\97759\Desktop\project\20220111'

    filenames=fileScanner(sourceWorkBookDirPath)
    #print(filenames)
    fileCreater(desWorkBookDirPath,filenames)
