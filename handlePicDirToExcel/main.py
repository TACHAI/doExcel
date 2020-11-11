import os
import xlrd,xlwt
from xlutils.copy import copy

# path = os.getcwd()
path = 'F:/南康区图书馆古籍展示文化长廊/'
pathName = '象山书院'
excelPath = 'C:/Users/Administrator/Desktop/江西书院/书院视频图片/图片/'

#构造数组对象
class Image:
    def __init__(self,name,text,uri):
        self.name = name
        self.text = text
        self.uri = uri





# 设置表格样式
def set_style(name,height,bold=0):
    style=xlwt.XFStyle()
    font=xlwt.Font()
    font.name=name
    font.colour_index=0
    font.height=height
    style.font=font
    return style


def write_excel():
    try:
        # 把目录中的文件名生成 list
        listAddr = os.listdir(path)
        #遍历第一层目录
        for i in range(len(listAddr)):
            name = listAddr[i]
            path2 = path+"/"+name
            listPath = os.listdir(path2)


            #新键excel文件
            excelPath1 = excelPath+"/"+ pathName+"/"+name+".xls"
            # 创建一个excel文件
            f=xlwt.Workbook()
            # 创建一个名为学生的sheet
            sheet1=f.add_sheet('Sheet0',cell_overwrite_ok=True)
            row0=["栏目","标题","发布时间","描述","标题图","图片集"]
            #写第一行
            for i in range(0,len(row0)):
                sheet1.write(0,i,row0[i],set_style('Times New Roman',220,True))
            #写第一列

            #写excel文件
            # old_wb = xlrd.open_workbook(excelPath)
            # new_wb = copy(old_wb)
            # new_ws = new_wb.get_sheet(0)
            # # 获取excel条数
            # high = old_wb.sheets()[0].nrows

            #遍历第二层目录
            for j in range(len(listPath)):

                picPathName = listPath[j]
                picName = picPathName[:-4]
                picPath = "/uploads/1/image/public/"+pathName+"/"+name+"/"+picPathName
                # 对象字典化
                print(picPath)
                sheet1.write(j+1,1,picName)
                sheet1.write(j+1,4,picPath)
                # new_ws.write(high+i, 1, name)
                # new_ws.write(high+i, 5, picPath)
            f.save(excelPath1)
        # new_wb.save(excelPath )
    except Exception as e:
        print(e)


if(__name__=='__main__'):
    for x in range(1,7):
        print(x)
    # write_excel()

