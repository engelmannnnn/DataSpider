from common.controler import xls_control
from tqdm import tqdm

def repeatClean_xls(fileDir, fileName, sheet, cols: [int]):
    xls_op = xls_control()
    xls_op.openXls(fileDir=fileDir, fileName=fileName, sheet=sheet)

    row = 2
    content_map = {}

    while xls_op.readXls(cols[0], row) != None:
        checkContent = [None for i in range(len(cols))]
        for index in range(len(cols)):
            checkContent[index] = xls_op.readXls(cols[index], row)
        if str(checkContent) in content_map:
            xls_op.delRows(row)
            print("row: {}".format(row), checkContent)
        else:
            content_map[str(checkContent)] = 1
            row +=1
    xls_op.saveXls()

# repeatClean_xls("result", "[2021]Marketplace_microsoft.xlsx", "Microsoft AppSource", [3,4])
def run():
    ls = [20170419, 20170415, 20170124, 20161207]
    for item in tqdm(ls):
        fileName = "[{}]Marketplace_microsoft.xlsx".format(item)
        print("当前处理的文件: ",fileName)
        repeatClean_xls("result", fileName,"Microsoft AppSource" ,[3,4])


run()