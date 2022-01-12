import xlrd
import xlsxwriter as xw
import os

if __name__ == '__main__':
    # 指定生成的文件完整路径
    desWorkBookPath=r'C:\Users\97759\Desktop\project\myexcel.xlsx'
    try:
        # 创建新工作簿
        desWorkbook = xw.Workbook(desWorkBookPath)
        # 创建新工作表
        desSheet = desWorkbook.add_worksheet('desSheet')
    except Exception as e:
        print('文件创建错误：'+e)
        try:
            desWorkbook = xlrd.open_workbook(desWorkBookPath)
            desSheet = desWorkbook.sheet_loaded(1)
        except Exception as e:
            print("文件打开错误："+e)


    #初始化行列标记为0
    pointRow=0
    pointCols=0

    #指定文件所在的目录路径
    dirPath=r'C:\Users\97759\Desktop\project\hb'
    for root,dirnames,filenames in os.walk(dirPath):
        for filename in filenames:
            #数据缓存列表
            data = []
            workFilePath=root+'\\'+filename
            print(filename+"----------------------1")

            # 打开源工作簿
            try:
                sourceWorkbook = xlrd.open_workbook(workFilePath)
                sourceSheet = sourceWorkbook.sheets()[1]
            except Exception as e:
                print("文件打开错误：" + e)
                exit(1)


            #获取工作表的行数
            rows=sourceSheet.nrows
            for row in range(rows):
                #将工作表数据添加到data
                data.append(sourceSheet.row_values(row))
                #print(sourceSheet.row_values(row),end='')

            dataRows=len(data)
            dataCols=len(data[0])

            for row in range(dataRows):
                for cols in range(dataCols):
                    #前两行数据是标题，不进行写操作
                    if row > 1:
                        #所有列位置+1再写，空出第一列用于记录数据来源工作簿名称
                        desSheet.write(pointRow+row,cols+1,data[row][cols])
                        #在第一列写入来源工作簿名称
                        desSheet.write(pointRow+row, 0, filename)
                        print("行：",pointRow+row,"列：",cols,"数据：",data[row][cols],"数据行数:",dataRows)
            #需要添加到同一个Sheet，加上前面的记录行位置
            pointRow=pointRow+dataRows
            print("当前记录行：",pointRow)
    desWorkbook.close()