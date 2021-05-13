from pathlib import Path
import requests
import time
# 网络请求
def request(url, method='GET', **kwargs):
    if method == 'GET':
        return requests.request(url=url, method=method, **kwargs)
    elif method == 'POST':
        return requests.request(url=url, method=method, **kwargs)
    elif method == 'PUT':
        return requests.request(url=url, method='put', **kwargs)
    elif method == 'DELETE':
        return requests.request(url=url, method='delete', **kwargs)


# 返回运行时间
def get_time(func):
    def wraper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        print("run time: {}s".format(round(end_time - start_time, 2)))
        return result

    return wraper

# 返回文件目录
def getDir(fileDir, fileName):
    root = Path().cwd().parent
    filePath = Path.joinpath(root, fileDir)
    return  Path.joinpath(filePath, fileName)



if __name__ == "__main__":
    a = getDir("result","a.txt")
    print(a)