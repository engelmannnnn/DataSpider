from common.public import request,getDir
import time, datetime
import traceback


class findDate():
    def __init__(self):
        self.params = {
            "FilteredLoading": "true",
            "country": "US",
            "region": "ALL",
            "version": "2020-09-22",
            "page": "1",
            "product": "dynamics-365-business-central"
        }
        self.Headers = {
            "accept-encoding": "gzip, deflate, br",
            "accept-language": "zh-CN,zh;q=0.9,en;q=0.8",
            "content-type": "application/json",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.128 Safari/537.36"
        }
        self.url = "https://appsource.microsoft.com/view/tiledata/"


        self.start_time = datetime.datetime.strptime("2014-01-01", "%Y-%m-%d")
        self.timeDelta = datetime.timedelta(days=1)
        self.end_time = datetime.datetime.now()

        self.save = []

    def findParams(self):
        params = self.params.copy()
        Ptime = self.start_time

        path = getDir("result","date.txt")
        file = open(path, encoding="utf-8", mode="a")
        file.write(time.ctime()+"\n")

        while Ptime < self.end_time:
            try:
                strTime = Ptime.strftime("%Y-%m-%d")
                print(strTime)
                params["version"] = strTime
                res = request(url=self.url, params=params, headers=self.Headers)
                if res.status_code == 200:
                    self.save.append(strTime)
                    print("find: {}".format(strTime))
                time.sleep(0.1)
            except Exception:
                print(traceback.format_exc())
            finally:
                Ptime += self.timeDelta
                file.write(str(self.save))
                print(self.save)
        file.close()






tst = findDate()
tst.findParams()

