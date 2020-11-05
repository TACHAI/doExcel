import os
import xlrd,xlwt
from xlutils.copy import copy

# path = os.getcwd()
path = 'F:/江西省图书馆20201016镜像视频（257）'

excelPath = 'C:/Users/Administrator/Desktop/江西省公共数字文化服务/视频信息表.xls'



def read_excel(path):
    old_wb = xlrd.open_workbook(excelPath)

    sheet = old_wb.sheets()[0]
    high = sheet.nrows

    sum=0
    for w in range(5,high):
        name = sheet.cell(w,13).value
        videoPath=sheet.cell(w,17).value[0:-4]+".mp4"
        print("name:"+name +"==videoPath:"+videoPath)
        nameList = videoPath.split("/")
        newPath=path
        for i in range(0,len(nameList)-1):
            newPath=newPath +"/"+nameList[i]

        newName=newPath+"/"+name+".mp4"
        oldName=path +"/"+videoPath
        print("oldName:"+oldName)
        print("newName:"+newName)
        sum=sum+1
        os.rename(oldName,newName)

    print("修改了："+str(sum)+"条数据")









if(__name__=='__main__'):
    read_excel(path)
