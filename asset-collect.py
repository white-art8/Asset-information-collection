# -*- coding:utf-8 -*-
# @Time :2022/3/16 9:33
# @Author ：chenjianzhi
# @File ：站长之家.py
# @Software :PyCharm

import requests  #requests库用于爬取网页
import time
import re
import random
from tqdm import tqdm
from selenium.webdriver import Chrome
from selenium import webdriver
import selenium.webdriver as wb
from selenium.webdriver.common.keys import Keys
import xlwt
from colorama import  init, Fore, Back, Style

init(autoreset=True)
class Colored(object):
    #  前景色:红色  背景色:默认
    def red(self, s):
        return Fore.RED + s + Fore.RESET
    #  前景色:绿色  背景色:默认
    def green(self, s):
        return Fore.GREEN + s + Fore.RESET
    #  前景色:黄色  背景色:默认
    def yellow(self, s):
        return Fore.YELLOW + s + Fore.RESET
    #  前景色:蓝色  背景色:默认
    def blue(self, s):
        return Fore.BLUE + s + Fore.RESET
    #  前景色:洋红色  背景色:默认
    def magenta(self, s):
        return Fore.MAGENTA + s + Fore.RESET
    #  前景色:青色  背景色:默认
    def cyan(self, s):
        return Fore.CYAN + s + Fore.RESET
    #  前景色:白色  背景色:默认
    def white(self, s):
        return Fore.WHITE + s + Fore.RESET
    #  前景色:黑色  背景色:默认
    def black(self, s):
        return Fore.BLACK
    #  前景色:白色  背景色:绿色
    def white_green(self, s):
        return Fore.WHITE + Back.GREEN + s + Fore.RESET + Back.RESET


def wash1(content):
    data=[]
    A0=re.compile(r'主办单位名称(.*?)主办单位性质', re.S)
    A1 = re.compile(r'主办单位性质(.*?)网站备案/许可证号', re.S)
    A2 = re.compile(r'网站备案/许可证号(.*?)查看截图', re.S)
    A3 = re.compile(r'网站名称(.*?)网站首页网址', re.S)
    A4 = re.compile(r'网站首页网址(.*?)安全认证', re.S)

    company_name = "公司名称：" + str(''.join(re.findall(A0, content))).strip()
    company_nature = "单位性质：" + str(''.join(re.findall(A1, content))).strip()
    license_key = "网站备案/许可证号：" + str(''.join(re.findall(A2, content))).strip()
    web_name = "网站名称：" + str(''.join(re.findall(A3, content))).strip()
    first_site = "网站首页网址：" + str(''.join(re.findall(A4, content))).strip()

    data.append(company_name)
    data.append(company_nature)
    data.append(license_key)
    data.append(web_name)
    data.append(first_site)

    return data


def wash2(content):
    data=[]
    A0=re.compile(r'数据来源：(.*?)法定代表人', re.S)
    A1 = re.compile(r'法定代表人(.*?)注册资本', re.S)
    A2 = re.compile(r'注册资本(.*?)注册时间', re.S)
    A3 = re.compile(r'注册时间(.*?)公司状态', re.S)
    A4 = re.compile(r'公司状态(.*?)公司类型', re.S)
    A5 = re.compile(r'公司类型(.*?)工商注册号 ', re.S)
    A6 = re.compile(r'工商注册号(.*?)所属行业', re.S)
    A7 = re.compile(r'所属行业(.*?)纳税人识别号', re.S)
    A8 = re.compile(r'纳税人识别号(.*?)核准日期', re.S)
    A9 = re.compile(r'核准日期(.*?)注册地址', re.S)
    A10 = re.compile(r'注册地址(.*?)经营范围', re.S)
    A11 = re.compile(r'经营范围(.*)', re.S)


    people = "法定代表人：" + str(''.join(re.findall(A1, content))).strip()
    asset = "注册资本：" + str(''.join(re.findall(A2, content))).strip()
    Registr_time = "注册时间：" + str(''.join(re.findall(A3, content))).strip()
    company_statue = "公司状态：" + str(''.join(re.findall(A4, content))).strip()
    company_sort = "公司类型：" + str(''.join(re.findall(A5, content))).strip()
    regis_number = "工商注册号：" + str(''.join(re.findall(A6, content))).strip()
    industry = "所属行业：" + str(''.join(re.findall(A7, content))).strip()
    Tax_ident_number = "纳税人识别号：" + str(''.join(re.findall(A8, content))).strip()
    Appr_date = "核准日期：" + str(''.join(re.findall(A9, content))).strip()
    Registered_Address = "注册地址：" + str(''.join(re.findall(A10, content))).strip()
    Nature_Business = "经营范围：" + str(''.join(re.findall(A11, content))).strip()
    source="数据来源：" + str(''.join(re.findall(A0, content))).strip()

    data.append(people)
    data.append(asset)
    data.append(Registr_time)
    data.append(company_statue)
    data.append(company_sort)
    data.append(regis_number)
    data.append(industry)
    data.append(Tax_ident_number)
    data.append(Appr_date)
    data.append(Registered_Address)
    data.append(Nature_Business)
    data.append(source)

    return data



def main():
    result = []  # 先命名一个list为result
    headers = {'user-agent': 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.187.119.204 Safari/537.36'}
    with open('域名输入.txt', encoding='utf-8') as f:
        for line in f:
            # result.append(line.strip('\n').split('*')[0])  # list.append函数用于提取内容，截取输出为 n个list，[0]表示提取list中的第一个string
            result.append(line.strip('\n'))                  # strip('\n')  可去除换行符
    print(result)

    data=[]

    book = xlwt.Workbook(encoding="utf-8", style_compression=0)
    sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)
    for wd in range(20):     #利用循环设置列宽
        col = sheet.col(wd)
        col.width = 256 * 20
    #xlwt创建时使用的默认宽度为2960，既11个字符0的宽度
    #width = 256 * 20     256为衡量单位，20表示20个字符宽度
    j=-1
    for i in tqdm(result):
        url = 'https://icp.chinaz.com/' + i
        time.sleep(random.random()*3)
        option = webdriver.ChromeOptions()
        option.add_argument('headless')    #不开浏览器运行
        option.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
        driver = webdriver.Chrome(chrome_options=option)
        driver.get(url)

        content1 = driver.find_element_by_xpath("/html/body/div[2]/div[3]/div[2]").text
        content2 = driver.find_element_by_xpath("/html/body/div[2]/div[3]/div[3]").text

        data1=wash1(content1)
        data2=wash2(content2)
        data=data1
        for i in data2:
            data.append(i)
        print(data)
        color = Colored()
        # for i in data:
        #     print(color.white_green(i))       #输出带颜色
        print(color.white_green("--------------------------------------------------------------"))
        j=j+1
        for i in range(len(data)):
            sheet.write(j, i, data[i])


        data = []
    book.save("域名备案信息输出.xls")



main()