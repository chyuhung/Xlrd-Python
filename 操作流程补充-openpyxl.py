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

class WorkBook():
    def __init__(self,dirpath,filename):
        self.dirpath=dirpath
        self.filename=filename
        self.workbook_path=self.get_workbook_path()
    def get_workbook_path(self):
        workbook_path = self.dirpath + '\\' + self.filename
        return workbook_path
    def get_workbook_data(self):
        workbook=openpyxl.load_workbook(self.workbook_path,read_only=True,data_only=True,keep_links=True,keep_vba=True)
        data=[]
        datalist=[]
        for sheet in workbook:
            for row in sheet.rows:
                data.append(row)
            datalist.append(data)
        return datalist


if __name__ == '__main__':
    #文件操作目录
    workdirpath=r'C:\Users\97759\Desktop\project\20220111'
    filenames=filescanner(workdirpath)
    filepaths=getfilepaths(workdirpath,filenames)
    for filename in filenames:
        print('yes')
        workbook=WorkBook(workdirpath,filename)
        datalist=workbook.get_workbook_data()
        for data in datalist:
            for line in data:
                print(line)
