import os
import re
import sys
import winreg
import zipfile

import requests
from seleniumwire import webdriver

# from seleniumwire.webdriver.common.desired_capabilities import DesiredCapabilities
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option('detach', True)
chrome_options.add_experimental_option('w3c', False)
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
# chrome_options.add_argument('--headless')
# chrome_options.add_argument('--disable-gpu')
# chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument("--auto-open-devtools-for-tabs")

# d = DesiredCapabilities.CHROME
# d['loggingPrefs'] = { 'performance':'ALL' }

url = 'http://npm.taobao.org/mirrors/chromedriver/'  # chromedriver download link


def get_path():
    return os.path.dirname(os.path.realpath(sys.argv[0]))
    # return os.path.dirname(os.path.realpath(__file__))


def get_Chrome_version():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Google\Chrome\BLBeacon')
    version, types = winreg.QueryValueEx(key, 'version')
    return version


def get_server_chrome_versions():
    '''return all versions list'''
    versionList = []
    url = "http://npm.taobao.org/mirrors/chromedriver/"
    rep = requests.get(url).text
    result = re.compile(r'\d.*?/</a>.*?Z').findall(rep)
    for i in result:
        version = re.compile(r'.*?/').findall(i)[0]  # 提取版本号
        versionList.append(version[:-1])  # 将所有版本存入列表
    return versionList


def download_driver(download_url):
    '''下载文件'''
    file = requests.get(download_url)
    with open("chromedriver.zip", 'wb') as zip_file:  # 保存文件到脚本所在目录
        zip_file.write(file.content)
        print('下载成功')


def download_lase_driver(download_url, chromeVersion, chrome_main_version):
    '''更新driver'''
    versionList = get_server_chrome_versions()
    if chromeVersion in versionList:
        download_url = f"{url}{chromeVersion}/chromedriver_win32.zip"
    else:
        for version in versionList:
            if version.startswith(str(chrome_main_version)):
                download_url = f"{url}{version}/chromedriver_win32.zip"
                break
        if download_url == "":
            print("暂无法找到与chrome兼容的chromedriver版本，请在http://npm.taobao.org/mirrors/chromedriver/ 核实。")

    download_driver(download_url=download_url)
    path = get_path()
    print('当前路径为：', path)
    unzip_driver(path)
    os.remove("chromedriver.zip")
    dri_version = get_version()
    if dri_version == 0:
        return 0
    else:
        print('更新后的Chromedriver版本为：', dri_version)


def get_version():
    '''查询系统内的Chromedriver版本'''
    outstd2 = os.popen('chromedriver --version').read()
    try:
        out = outstd2.split(' ')[1]
    except:
        return 0
    return out


def unzip_driver(path):
    '''解压Chromedriver压缩包到指定目录'''
    f = zipfile.ZipFile("chromedriver.zip", 'r')
    for file in f.namelist():
        f.extract(file, path)


def check_update_chromedriver():
    try:
        chromeVersion = get_Chrome_version()
    except:
        print('未安装Chrome，请在GooGle Chrome官网：https://www.google.cn/chrome/ 下载。')
        return 0

    chrome_main_version = int(chromeVersion.split(".")[0])  # chrome主版本号

    try:
        driverVersion = get_version()
        driver_main_version = int(driverVersion.split(".")[0])  # chromedriver主版本号
    except:
        print('未安装Chromedriver，正在为您自动下载>>>')
        download_url = ""
        if download_lase_driver(download_url, chromeVersion, chrome_main_version) == 0:
            return 0
        driverVersion = get_version()
        driver_main_version = int(driverVersion.split(".")[0])  # chromedriver主版本号

    download_url = ""
    if driver_main_version != chrome_main_version:
        print("chromedriver版本与chrome浏览器不兼容，更新中>>>")
        if download_lase_driver(download_url, chromeVersion, chrome_main_version) == 0:
            return 0
    else:
        print("chromedriver版本已与chrome浏览器相兼容，无需更新chromedriver版本！")


if __name__ == '__main__':
    check_update_chromedriver()
