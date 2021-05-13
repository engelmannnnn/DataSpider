from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
import time
from common.public import getDir
from common.controler import xls_control
import re
import xml.etree.ElementTree as ET
import collections
from lxml import etree
import traceback


class salesforce():
    def __init__(self):
        self.xls_op = xls_control()
        self.xls_op.openXls(fileDir="result", fileName="Marketplace_salesforce.xlsx", sheet="Salesforce AppExchange")
        title = ["app_name", "label1", "label2", "publisher", "price", "rating", "info"]
        self.xls_op.setTitle(title)
        self.appCount = 1
        # appInfo 使用队列存储
        self.app = collections.deque()

        self.url = "https://appexchange.salesforce.com/appxStore"
        self.browser = webdriver.Chrome()
        self.browser.maximize_window()
        self.browser.get(self.url)
        self.findWait = WebDriverWait(driver=self.browser, timeout=5, poll_frequency=0.5, ignored_exceptions=None)
        self.page = 0

        # 如果页面提示授权cookie , 则点击一下
        cookie_accept = "//*[@id=\"onetrust-accept-btn-handler\"]"
        judge_cookie = self.judgeElement(cookie_accept, "cookie_accept")
        if judge_cookie:
            self.browser.find_element_by_xpath(cookie_accept).click()

    def judgeElement(self, xpath, elemName=""):
        getItem = True
        try:
            self.findWait.until(method=lambda driver: driver.find_element_by_xpath(xpath),
                                message="未找到元素: {}".format(elemName))
        except Exception:
            getItem = False
        return getItem

    def findMore(self):
        # while True:
        # 页面滑倒底部
        js = "var q=document.documentElement.scrollTop=100000"
        self.browser.execute_script(js)
        time.sleep(1)
        show_more = "//*[@id=\"appx-load-more-button-id\"]"
        judge_show_more = self.judgeElement(show_more, "show_more")
        # 如果有show more按钮 就点击一下
        if judge_show_more:
            try:
                self.browser.find_element_by_xpath(show_more).click()
                self.page += 1
                print("已加载 {} 页".format(self.page))
                time.sleep(8)
            except Exception:
                print("点击[查看更多]按钮失败, 正在进行重试...")
                self.findMore()
        else:
            print("未找到[查看更多]按钮, 共加载 {} 页".format(self.page))
        #     break

    def saveData(self):
        try:
            app = {"app_name": "",
                   "publisher": "",
                   "price": "",
                   "rating": "",
                   "info": "",
                   "label1": "",
                   "label2": "",
                   "no": ""}
            page_code = self.browser.page_source
            html = etree.HTML(page_code)
            cur_no = len(self.app)

            app_names_xpath = '//*[@id="appx-table-results"]/li[*]/a/span[2]/span[2]/span[1]/span[2]/span[1]'
            app_names = html.xpath(app_names_xpath)

            while cur_no < len(app_names):
                app["no"] = cur_no + 1
                self.appCount = app["no"]
                app["app_name"] = app_names[cur_no].text.replace("\n", "").replace("\t", "")

                publishers_xpath = '/html/body/div[1]/div[1]/div[1]/div/div/span[2]/div/ul/li[{}]/a/span[2]/span[2]/span[1]/span[2]/span[2]'.format(
                    app["no"])
                publishers = html.xpath(publishers_xpath)
                app["publisher"] = publishers[0].text

                pricesFree_xpath = '/html/body/div[1]/div[1]/div[1]/div/div/span[2]/div/ul/li[{}]/a/span[2]/span[2]/span[2]/span'.format(
                    app["no"])
                pricesPaid_xpath = '/html/body/div[1]/div[1]/div[1]/div/div/span[2]/div/ul/li[{}]/a/span[2]/span[2]/span[2]/div/div/span'.format(
                    app["no"])
                pricesFree = html.xpath(pricesFree_xpath)
                pricesPaid = html.xpath(pricesPaid_xpath)
                if len(pricesFree):
                    app["price"] = pricesFree[0].text
                elif len(pricesPaid):
                    app["price"] = pricesPaid[0].text
                else:
                    print("APP: {} 没有找到价格标签".format(app["app_name"]))

                ratings_path = '/html/body/div[1]/div[1]/div[1]/div/div/span[2]/div/ul/li[{}]/a/span[1]/span[3]/span[2]/span[1]/span/span/@class'.format(app["no"])
                               # '/html/body/div[1]/div[1]/div[1]/div/div/span[2]/div/ul/li[8]/a/span[1]/span[3]/span[2]/span[1]/span/span'
                ratings = html.xpath(ratings_path)
                try:
                    app["rating"] = int(ratings[0][-2:])/10
                except:
                    app["rating"] = 0

                infos_path = '/html/body/div[1]/div[1]/div[1]/div/div/span[2]/div/ul/li[{}]/a/span[2]/span[2]/span[5]/span'.format(
                    app["no"])
                infos = html.xpath(infos_path)
                app["info"] = infos[0].text.replace("\n", "").replace("\t", "")

                labels1_xpath = '/html/body/div[1]/div[1]/div[1]/div/div/span[2]/div/ul/li[{}]/a/span[2]/span[2]/span[4]/span/span[2]'.format(app["no"])
                labels2_xpath = '/html/body/div[1]/div[1]/div[1]/div/div/span[2]/div/ul/li[{}]/a/span[2]/span[2]/span[4]/span/span[4]'.format(app["no"])
                labels1 = html.xpath(labels1_xpath)
                labels2 = html.xpath(labels2_xpath)
                if len(labels1):
                    app["label1"] = labels1[0].text.replace("\n", "").replace("\t", "")
                else:
                    app["label1"] = " \ "
                if len(labels2):
                    app["label2"] = labels2[0].text.replace("\n", "").replace("\t", "")
                else:
                    app["label2"] = " \ "

                print(app)
                # 注意, dict是一个指针存储单元, 存进队列里的时候, 如果dict的值改了, 之前的也改了, 所以要复制一个存进去
                self.app.append(app.copy())
                # print(self.app)
                cur_no = len(self.app)

        except Exception:
            print(traceback.format_exc())

        finally:
            while self.app:
                app = self.app.popleft()
                self.xls_op.writeXls(title="app_name", content=app["app_name"], row=int(app["no"]) + 1)
                self.xls_op.writeXls(title="label1", content=app["label1"], row=int(app["no"]) + 1)
                self.xls_op.writeXls(title="label2", content=app["label2"], row=int(app["no"]) + 1)
                self.xls_op.writeXls(title="publisher", content=app["publisher"], row=int(app["no"]) + 1)
                self.xls_op.writeXls(title="rating", content=app["rating"], row=int(app["no"]) + 1)
                self.xls_op.writeXls(title="price", content=app["price"], row=int(app["no"]) + 1)
                self.xls_op.writeXls(title="info", content=app["info"], row=int(app["no"]) + 1)
            self.xls_op.saveXls()
            print("\n当前进度: 已抓取数据 {} 条, 进度: {}%\n".format(self.appCount, round(int(self.appCount)/39.13, 3)))

    def run(self):
        while True:
            try:
                self.saveData()
                self.findMore()
            except Exception:
                print("#######抓取结束########")
                print(traceback.format_exc())



    def test(self):
        filePath = getDir("result", "test.txt")
        f = open(filePath, mode="a")
        f.write(str(self.browser.page_source.encode("gbk", "ignore")))
        f.close()


tst = salesforce()
# tst.saveData()
tst.run()
