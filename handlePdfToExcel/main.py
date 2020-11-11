import os
import xlrd,xlwt
from xlutils.copy import copy

# pdf地址写入excel
# path = os.getcwd()
path = 'F:/南康区图书馆古籍展示文化长廊/'
pathName = 'pdf'
excelPath = 'C:/Users/Administrator/Desktop/江西书院/书院视频图片/图片/'


if(__name__=='__main__'):
    excelPath1 = excelPath+"/pdf.xls"
    f=xlwt.Workbook()
    # 创建一个名为学生的sheet
    sheet1=f.add_sheet('Sheet0',cell_overwrite_ok=True)
    row0=["id","name","path"]
    #写第一行
    for i in range(0,len(row0)):
        sheet1.write(0,i,row0[i],set_style('Times New Roman',220,True))
    listAddr = os.listdir(path)
    temp=1
    #遍历第一层目录
    for i in range(len(listAddr)):
        name = listAddr[i]
        path2 = path+"/"+name
        listPath = os.listdir(path2)
        for j in range(len(listPath)):
            pdfPathName = listPath[j]
            sheet1.write(temp,0,temp)
            sheet1.write(temp,1,name)
            pdfPath="/static/pdf/"+name+"/"+pdfPathName
            sheet1.write(temp,2,pdfPath)
            print(pdfPath)
            temp =temp+1
    f.save(excelPath1)


