from openpyxl import load_workbook
import os
import shutil


def handler_excel(filename=r'./photo.xlsx'):
    # 根据文件路径加载一个excel表格，这里包含所有的sheet
    excel = load_workbook(filename)
    # 根据sheet名称加载对应的table
    table = excel.get_sheet_by_name('Sheet1')
    imgnames = []
    # 读取所有列
    for column in table.columns:
        for cell in column:
            imgnames.append(cell.value)
    # 选择图片
    pickImg(imgnames)


def pickImg(pickImageNames):
    # 遍历所有图片集合的文件名
    for image in os.listdir(r'./car-model'):
        print(image)
        if image in pickImageNames:
            oldname = r"./car-model/"+image
            newname = r"./target/"+image
            # 文件拷贝
            shutil.copyfile(oldname, newname)


handler_excel()
