 # -*- coding: UTF-8 -*-
 # 用于从武汉大学教务系统获取成绩单并导出为excel文件
 # 陈锐铭 2017.3.22-3.25
 # 版本：2.0
 # 特别感谢黄潋哲，他帮忙写了正则和输出excel的部分，对，就是比较整齐比较好看的那部分
 # 第一次写一个比较有点卵用的程序，很惭愧，就做了一点微小的工作
un = input('请输入学号并按回车确认：')
pwd = input('请输入密码并按回车确认：')
print('好了不用输验证码\n获取中，请稍候')

import requests
from bs4 import BeautifulSoup as bs
import re
import xlwt

session = requests.Session()
whuigtw = 'http://cas.whu.edu.cn/authserver/login?service=http://my.whu.edu.cn'  # 信息门户
edustm = 'http://210.42.121.241/common/caslogin.jsp'  # 信息门户里的教务系统链接


def toJson(str):
    soup = bs(str)
    tt = {}
    for inp in soup.form.find_all('input'):
        if inp.get('name') != None:
            tt[inp.get('name')] = inp.get('value')
    return tt


lt = session.get(whuigtw)
soup = toJson(lt.text)
# 获取流水号
pd = {'username': 'un',
      'password': 'pwd',
      'lt': soup["lt"],
      'dllt': 'userNamePasswordLogin',
      'execution': 'e1s1',
      '_eventId': 'submit',
      'rmShown': '1'}
# 登陆信息
hs = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
      'Accept-Encoding': 'gzip, deflate',
      'Accept-Language': 'zh-CN,zh;q=0.8',
      'Cache-Control': 'max-age=0',
      'Connection': 'keep-alive',
      'Content-Length': '170',
      'Content-Type': 'application/x-www-form-urlencoded',
      'Cookie': 'route=165d580f5e5507854f7a53b13f17f8e0; pgv_pvi=4856056832; JSESSIONID_ids1=0001fOS57t3b2FW-S3Y736RRE0b:-4602F9; route=6c1010bc2426f9d3b968d6de008806d5; JSESSIONID_ids2=0001a2HpBqt6NyGQUi37yzSwbCA:28FRNNTPRE',
      'Host': 'cas.whu.edu.cn',
      'Origin': 'http://cas.whu.edu.cn',
      'RA-Sid': '3BACB0DA-20150528-121834-484636-b8df59',
      'RA-Ver': '3.0.8',
      'Upgrade-Insecure-Requests': '1',
      'Referer': 'http://cas.whu.edu.cn/authserver/login?service=http://my.whu.edu.cn',
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36'}
# 头文件 cookie每次都一样= =

i = session.post(whuigtw, data=pd)  # 登陆信息门户
j = session.get(edustm)  # 经由信息门户进入教务系统


def csrf(str):
    soup = bs(str)
    c = {}
    for r in soup.find_all('div'):
        if r.get('onclick') != None:
            c[r.get('name')] = r.get('onclick')

    return c


ct = csrf(j.text)
Lsn = ct[None].split("'")[1]
t = Lsn.replace('/servlet/Svlt_QueryStuLsn?action=queryStuLsn&csrftoken=', '/servlet/Svlt_QueryStuScore?csrftoken=')
# 进入教务系统后获取csrftoken

print(t)  # 我也不知道为什么要先把csrftoken打出来但是不打出来后面就没东西？？？

import time

tformat = '%a%%20%b%%20%d%%20%y%%20%XGMT+0800%%20(CST)'
ti = time.strftime(tformat, time.localtime())

print(ti)  # 我也不知道为什么还要把时间打出来但是不打出来后面就没东西？？？
# 获取时间 填入url

score = 'http://210.42.121.134' + t + '&year=0&term=&learnType=&scoreFlag=0&t=' + ti  # 成绩单框架的url
s = session.get(score)
# 获取成绩单页面


def getScoreInfo(str):
    state = r'<tr null>(.*?)</tr>'  # 分课程在<tr null>标签内
    r = re.findall(state, str, re.S | re.M)
    result = []
    for lesson in r:
        stateL = r'<td>(.*?)</td>'  # 课程信息在<td>标签里
        les = re.findall(stateL, lesson, re.S | re.M)
        lesson = []
        for i in range(0, 10):
            lesson.append(les[i])
        result.append(lesson)
    return result


scorelist = getScoreInfo(s.text)
# 成绩单整理
# with open('/Users/chenruiming/Downloads/成绩单.txt', 'w') as f:
 #   f.write(str(scorelist))
def writeExcel(result):
    f = xlwt.Workbook(encoding='utf-8')
    sheet1 = f.add_sheet(u'成绩表', cell_overwrite_ok=True)
    row = 0
    for lesson in result:
        col = 0
        for info in lesson:
            sheet1.write(row, col, info, xlwt.Style.easyxf())
            col = col + 1
        row = row + 1
    f.save('score.xls')
    # 写Excel的方法，需要用到第三方库xlwt


writeExcel(scorelist)
print('完结撒花')