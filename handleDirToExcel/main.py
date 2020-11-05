import os
import xlrd,re
from xlutils.copy import copy

# path = os.getcwd()
path = 'F:/五楼红色图书馆文化长廊/图书'
imagpath = 'F:/五楼红色图书馆文化长廊/封面'

excelPath = 'C:/Users/Administrator/Desktop/省图文化长廊/Resource.xls'



try:
    old_wb = xlrd.open_workbook(excelPath)
    new_wb = copy(old_wb)
    new_ws = new_wb.get_sheet(0)
    # 获取excel条数
    high = old_wb.sheets()[0].nrows
    # 把目录中的文件名生成 list
    listAddr = os.listdir(path)
    for i in range(len(listAddr)):
        # 写 行列值
        m = re.search('\d+',listAddr[i])
        id = m.group(0)
        print(id)
        new_ws.write(high+i, 0, id)
        name = listAddr[i].replace(id,'')
        print(name[:-4])
        new_ws.write(high+i, 3, name)
        filePath= path+'/'+listAddr[i]
        print(filePath)
        new_ws.write(high+i, 5, filePath)
        image = listAddr[i].replace('pdf','png')
        imagePath=imagpath+'/'+image
        print(imagePath)
        new_ws.write(high+i, 6, imagePath)
    new_wb.save(excelPath )
except Exception as e:
    print(e)


