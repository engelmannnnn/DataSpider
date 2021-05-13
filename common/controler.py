from openpyxl import Workbook, load_workbook
from common.public import getDir
import traceback
import time


class xls_control():
    def __init__(self):
        pass
    def creatXls(self):
        pass

    def openXls(self, fileDir="result", fileName="Marketplace_microsoft.xlsx", sheet="Microsoft AppSource"):
        self.fileDir, self.fileName = fileDir, fileName
        filePath = getDir(fileDir, fileName)
        try:
            self.workbook = load_workbook(filePath)
            self.worksheet = self.workbook[sheet]

        except FileNotFoundError:
            print("文件打开错误, 文件未找到! 文件路径: {}".format(filePath))
        except Exception:
            print("文件打开错误!")
            print(traceback.format_exc())

    def setTitle(self, titles):
        self.title = {}
        for index in range(0, len(titles)):
            self.title[titles[index]] = index+1

    def writeXls(self,title, content, row):
        if title in self.title:
            col = chr( self.title[title] + 64 )
            cell = col+str(row)
            self.worksheet[cell] = content
        else:
            print("数据写入错入,未找到对应列: {} , 数据内容: {}".format(title, content))

    def readXls(self,column: int, row: int):
        col = chr(column + 64)
        cell = str(col+str(row))
        return self.worksheet[cell].value

    def delRows(self,row: int):
        self.worksheet.delete_rows(row)

    def saveXls(self):
        now = str(round(time.time()))
        filePath = getDir(self.fileDir, now + self.fileName)
        self.workbook.save(filePath)
        print("文件保存路径: {}".format(filePath))






if __name__ == "__main__":
    tst = xls_control(row=2)
    tst.openXls("result","Marketplace_microsoft.xlsx")
    title = ["product_label1","product_label2"]
    tst.setTitle(title)
    tst.writeXls("product_label1","hahaha")
    tst.saveXls()
