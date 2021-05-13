from common.controler import xls_control
from common.public import request
import time
import traceback
from tqdm import tqdm


class microsoft_appSource():
    def __init__(self, row=2):
        # excel 记录的起始行数, row = 1 是 title行
        self.row = row
        self.xls_op = xls_control()
        self.xls_op.openXls(fileDir="result", fileName="Marketplace_microsoft.xlsx", sheet="Microsoft AppSource")
        title = ["product_label1", "product_label2", "app_name", "publisher", "type1", "type2", "type3", "star"]
        self.xls_op.setTitle(title)

    def OutApiData(self):
        apiUrlOut = "https://appsource.microsoft.com/view/tiledata/"
        params = {
            "FilteredLoading": "true",
            "country": "US",
            "region": "ALL",
            "version": "2017-04-24",
            "page": "1",
            "product": "dynamics-365-business-central"
        }

        product_label = {
            "Dynamics 365": ["dynamics-365-business-central", "dynamics-365-commerce",
                             "dynamics-365-for-customer-services",
                             "dynamics-365-customer-voice", "dynamics-365-for-field-services",
                             "dynamics-365-for-finance-and-operations",
                             "dynamics-365-human-resources", "dynamics-365-marketing", "dynamics-365-mixed-reality",
                             "dynamics-365-project-operations", "dynamics-365-for-project-service-automation",
                             "dynamics-365-for-sales", "dynamics-365-supply-chain-management"],
            "Microsoft 365": ["excel", "onenote", "outlook", "powerpoint", "project", "sharepoint", "teams", "word"],
            "Power Platform": ["powerapps", "power-automate", "power-bi", "power-bi-visuals", "power-virtual-agents"],
            "web-apps": ["web-apps"]
        }

        Headers = {
            "accept-encoding": "gzip, deflate, br",
            "accept-language": "zh-CN,zh;q=0.9,en;q=0.8",
            "content-type": "application/json",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.128 Safari/537.36"
        }

        for label1 in product_label:
            for label2 in product_label[label1]:
                params["product"] = label2
                page = 1
                try:
                    while True:
                        params["page"] = page
                        res = request(url=apiUrlOut, method="GET", headers=Headers, params=params)
                        retry = 0
                        while res.status_code != 200 and retry < 3:
                            time.sleep(2)
                            res = request(url=apiUrlOut, method="GET", headers=Headers, params=params)
                            retry +=1
                        try :
                            app_list = res.json()["apps"]["dataList"]
                        except Exception:
                            print("json解析失败: {}".format(res.text))
                            print(traceback.format_exc())
                            break

                        if  app_list == []:
                            break
                        else:
                            print("当前内容: {} 产品, {} 类, 第 {} 页".format(label1, label2, page))
                            for item in tqdm(app_list):
                                # print(item)
                                self.xls_op.writeXls(title="product_label1", content=label1, row=self.row)
                                self.xls_op.writeXls(title="product_label2", content=label2, row = self.row)
                                self.xls_op.writeXls(title="app_name", content=item["title"], row=self.row)
                                self.xls_op.writeXls(title="publisher", content=item["publisher"], row=self.row)
                                self.xls_op.writeXls(title="star", content=round(item["AverageRating"],2), row=self.row)
                                self.row += 1
                        page +=1
                        time.sleep(0.5)

                except Exception:
                    print(traceback.format_exc())
                finally:
                    self.xls_op.saveXls()

    def test(self):
        apiUrlOut = "https://appsource.microsoft.com/view/tiledata/"
        params = {
            "FilteredLoading": "true",
            "country": "US",
            "region": "ALL",
            "version": "2020-09-22",
            "page": "1",
            "product": "dynamics-365-business-central"
        }
        Headers = {
            "accept-encoding": "gzip, deflate, br",
            "accept-language": "zh-CN,zh;q=0.9,en;q=0.8",
            "content-type": "application/json",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.128 Safari/537.36"
        }
        res = request(url=apiUrlOut, method="GET", headers=Headers, params=params)
        print("test")
        app_list = res.json()["apps"]["dataList"]
        print(  round(app_list[0]["AverageRating"]),2)
        print(round(2.32323, 2))
        star = app_list[0]["AverageRating"]
        print(round(star, 2))


if __name__ == "__main__":
    tst = microsoft_appSource(row=2)
    tst.OutApiData()
