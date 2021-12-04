import os
import sys
import time
import tkinter as tk
import winreg
from tkinter import filedialog

import chardet
import pyperclip
import update_webdriver
from progress.bar import IncrementalBar
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait

update_webdriver.check_update_chromedriver()


class Autoahk_Upload:
    def __init__(self, username, password):
        self.options = webdriver.ChromeOptions()
        self.reg = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                  r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome")
        self.path = winreg.QueryValueEx(self.reg, 'InstallLocation')
        self.path = self.path[0] + "\\chrome.exe"
        self.options.binary_location = self.path
        self.options.add_argument('--ignore-certificate-errors')  # 关闭ssl错误提示
        self.options.add_argument('--start-maximized')  # 最大化浏览器窗口
        self.options.add_argument('--disable-gpu')  # 谷歌文档提到需要加上这个属性来规避bug
        # self.options.add_argument('blink-settings=imagesEnabled=false')  # 不加载图片, 提升速度  ，实测加入此项无法新建文章
        # self.options.add_argument('--headless')  # 不显示可视化界面
        prefs = {
            'profile.default_content_setting_values': {
                'notifications': 2
            }
        }
        self.options.add_experimental_option('prefs', prefs)  # 禁用浏览器弹窗
        self.options.add_experimental_option("detach", True)
        self.options.add_experimental_option('excludeSwitches', ['enable-logging'])  # 不打印无用日志
        self.options.add_experimental_option('excludeSwitches', ['enable-automation'])  # 关闭正在受自动化测试软件控制的提示
        self.browser = webdriver.Chrome(chrome_options=self.options)  # 打开谷歌浏览器
        self.browser.get("https://www.autoahk.com/write")
        self.browser.find_element_by_xpath("""//*[@id="login-box"]/div/div/div/form/div[2]/label[2]/input""").send_keys(
            username)
        self.browser.find_element_by_xpath("""//*[@id="login-box"]/div/div/div/form/div[2]/label[5]/input""").send_keys(
            password)
        self.browser.find_element_by_xpath("""//*[@id="login-box"]/div/div/div/form/div[2]/div[3]/button""").click()

    def upload(self, title, key):
        select = (By.XPATH, """//*[@id="write-head"]/div[2]/div[3]/select""")
        try:
            WebDriverWait(self.browser, 20).until(EC.presence_of_element_located(select))
            print('分类下拉列表加载成功')
            time.sleep(1)
            s1 = Select(self.browser.find_element_by_xpath("""//*[@id="write-head"]/div[2]/div[3]/select"""))
            if key == '1':
                s1.select_by_visible_text('AHKV1')
            elif key == '2':
                s1.select_by_visible_text('学习')
            elif key == '3':
                s1.select_by_visible_text('其他')
            elif key == '4':
                s1.select_by_visible_text('办公')
            elif key == '5':
                s1.select_by_visible_text('游戏')
            elif key == '6':
                s1.select_by_visible_text('AHKV2')
            elif key == '7':
                s1.select_by_visible_text('工具')
        except:
            print('分类下拉列表加载失败')

        tl = (By.XPATH, """/html/body/div[1]/div[2]/div[1]/div/main/div[1]/div[4]/textarea""")
        try:
            tl_element = WebDriverWait(self.browser, 20).until(EC.presence_of_element_located(tl))
            time.sleep(1)
            tl_element.click()
            tl_element.send_keys(title)
            print('标题输入成功')
        except:
            print('标题输入失败')

        # self.browser.find_element_by_xpath(
        #     """/html/body/div[1]/div[2]/div[1]/div/main/div[1]/div[4]/textarea""").click()
        # self.browser.find_element_by_xpath(
        #     """/html/body/div[1]/div[2]/div[1]/div/main/div[1]/div[4]/textarea""").send_keys(title)
        self.browser.find_element_by_xpath(
            """//*[@id="b2-editor-box"]/div[1]/div[1]/div[1]/div[1]/div/div[2]/button[3]""").click()
        self.browser.find_element_by_xpath(
            """/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[2]/textarea""").click()
        self.browser.find_element_by_xpath(
            """/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[2]/textarea""").send_keys(Keys.CONTROL, "v")
        self.browser.find_element_by_xpath("""/html/body/div[3]/div/div[2]/div[3]/div[2]/button[2]""").click()
        self.browser.find_element_by_xpath(
            """/html/body/div[1]/div[2]/div[1]/div/main/div[3]/div[5]/div[2]/button[2]""").click()

    def new_article(self):
        # time.sleep(2)
        new = (By.XPATH, """//*[@id="page"]/div[1]/div/div[1]/div/div[2]/div[2]/div[1]/div[2]/button/i""")
        try:
            new_element = WebDriverWait(self.browser, 20).until(EC.presence_of_element_located(new))
            time.sleep(1)
            new_element.click()
            print('点击新建文章成功')
        except:
            print('新建文章失败或结束')

        # self.browser.find_element_by_xpath(
        #     """//*[@id="page"]/div[1]/div/div[1]/div/div[2]/div[2]/div[1]/div[2]/button/i""").click()
        self.browser.find_element(By.XPATH,
                                  """//*[@id="page"]/div[1]/div/div[1]/div/div[2]/div[2]/div[1]/div[2]/button/i""").click()
        self.browser.find_element_by_xpath("""//*[@id="post-po-box"]/div/div/div[1]/div[1]/button""").send_keys(
            Keys.SPACE)
        self.browser.find_element_by_xpath("""//*[@id="post-po-box"]/div/div/div[1]/div[1]/button""").click()


if __name__ == '__main__':
    key = input('1.AHK V1\n2.学习\n3.其他\n4.办公\n5.游戏\n6.AHK V2\n7.工具\n请选择文章分类：')
    user_name = input('用户名：')
    pass_word = input('密码：')
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilenames()
    aa = Autoahk_Upload(user_name, pass_word)  # .upload("测试", "测试正文")
    count = 1
    for f in file_path:
        # 采样长度,最长采样长度为100，可调节
        sample_len = min(100, os.path.getsize(f))
        # 读取片段bytes
        raw = open(f, 'rb').read(sample_len)
        # 检测编码
        detect = chardet.detect(raw)
        with open(f, 'r+', encoding=detect['encoding'], errors='ignore') as file:
            # print(file.read())
            body = file.read()
            # with open(f, 'r', encoding='utf-8') as file:
            #     body = file.read()
            #     print(f)
            #     print(body)
            filename = f.split('.')[0]
            title = os.path.basename(filename)
            pyperclip.copy(body)
            aa.upload(title, key)
            print('上传成功' + ' ' + str(count) + '   ' + filename)
            aa.new_article()
        count += 1
    # aa.browser.quit()
    # aa.upload("测试", "ces")
    # aa.new_article()
