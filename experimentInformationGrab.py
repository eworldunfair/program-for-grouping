# grab all the class informations in the experiment platform of University
# 孟老师，实验安排，网页爬取
# 1. 2021-9-15 15:25:13
# 2. 2021-10-18 08:59:57
import xlwt
import re
import requests
import os
from bs4 import BeautifulSoup


headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36',
    'Cookie': '__utmz=16735861.1631088727.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); JSESSIONID=ACA281C62B2B2641F32347ABA75DD129; __utma=16735861.1895854436.1631088727.1631095260.1631159153.3; __utmc=16735861; __utmt=1; __utmb=16735861.7.10.1631159153'}
url0 = 'http://211.71.70.68/login'
Form_data = {'username': 'lcmeng', 'password': '123456'}
url1 = 'http://211.71.70.68/admin/weekappointment/5/'

# login
s = requests.Session()
s.post(url0, data=Form_data, headers=headers)

# Way 1：to creat a URL list with all the pages in.
URL = []
for week in range(5, 6):
    url2 = 'http://211.71.70.68/admin/weekappointment/{}/'.format(week)
    r2 = s.get(url2).text
    bs = BeautifulSoup(r2, "html.parser")
    tableList = bs.select('.p_10>.show_table')
    for tabel in tableList:
        timeList = re.findall(r'href="(.+?)"', str(tabel))
        for time in timeList:
            if 't=' in time:
                URL.append(time.replace('amp;', ''))

# # Way 2：to creat a URL list with all the pages in.
# URL = []
# for week in range(5, 1+7):
#     url2 = 'http://211.71.70.68/admin/weekappointment/{}/'.format(week)
#     r2 = s.get(url2)
#     # get the data list and eliminate the repeat data.
#     timeList = re.findall(r'\d{4}-\d{2}-\d{2}', r2.text)
#     timeList_noRepeat = []
#     for time in timeList:
#         if time not in timeList_noRepeat:
#             timeList_noRepeat.append(time)
#     # put all the url links in URL list.
#     for d in timeList_noRepeat:
#         for e in ['1', '2']:
#             for t in ['3', '4', '5', '6']:
#                 URL.append(
#                     'http://211.71.70.68/admin/weekappointment/p/?d={d1}&e={d2}&t={d3}'.format(d1=d, d2=e, d3=t))

# grab all the pages
course = '大学物理演示实验_(2021秋)'
jiaoshi_optics = {'邹晓琳': '20126288_邹晓琳_光学', '田富宇': '18121670_田富宇_光学', '罗珏婷': '20121659_罗珏婷_光学', '郭玉颖': '20126240_郭玉颖_光学', '张怀伟': '20118049_张怀伟_光学', '谢子灿': '20121665_谢子灿_光学', '徐文轩': '19120022_徐文轩_光学', '谢峥嵘': '20126286_谢峥嵘_光学', '郭哲灿': '20126241_郭哲灿_光学', '赵泽邦': '21118039_赵泽邦_光学', '贾志开': '20126245_贾志开_光学', '徐梦晨': '20126266_徐梦晨_光学', '唐陈丰': '20126257_唐陈丰_光学',
                  '薛勇': '21126261_薛勇_光学', '刘彬': '21118033_刘彬_光学', '毛炜昊': '21120382_毛炜昊_光学', '徐梦': '18118021_徐梦_光学', '高杨帆': '21121583_高杨帆_光学', '孟宁': '21121590_孟宁_光学', '庞艳兰': '21121591_庞艳兰_光学', '祝熙翔': '98930228_祝熙翔_光学', '张玉': '9611_张玉_光学', '杨亚杰': '9599_杨亚杰_光学', '郭亚光': '9635_郭亚光_光学', '孟令川': '9367_孟令川_光学', '周晓亮': '9295_周晓亮_光学'}
jjiaoshi_modern = {'邹晓琳': '20126288_邹晓琳_近代与综合', '田富宇': '18121670_田富宇_近代与综合', '罗珏婷': '20121659_罗珏婷_近代与综合', '郭玉颖': '20126240_郭玉颖_近代与综合', '张怀伟': '20118049_张怀伟_近代与综合', '谢子灿': '20121665_谢子灿_近代与综合', '徐文轩': '19120022_徐文轩_近代与综合', '谢峥嵘': '20126286_谢峥嵘_近代与综合', '郭哲灿': '20126241_郭哲灿_近代与综合', '赵泽邦': '21118039_赵泽邦_近代与综合', '贾志开': '20126245_贾志开_近代与综合', '徐梦晨': '20126266_徐梦晨_近代与综合', '唐陈丰': '20126257_唐陈丰_近代与综合',
                   '薛勇': '21126261_薛勇_近代与综合', '刘彬': '21118033_刘彬_近代与综合', '毛炜昊': '21120382_毛炜昊_近代与综合', '徐梦': '18118021_徐梦_近代与综合', '高杨帆': '21121583_高杨帆_近代与综合', '孟宁': '21121590_孟宁_近代与综合', '庞艳兰': '21121591_庞艳兰_近代与综合', '祝熙翔': '98930228_祝熙翔_近代与综合', '张玉': '9611_张玉_近代与综合', '杨亚杰': '9599_杨亚杰_近代与综合', '郭亚光': '9635_郭亚光_近代与综合', '孟令川': '9367_孟令川_近代与综合', '周晓亮': '9295_周晓亮_近代与综合'}
urlHead = 'http://211.71.70.68'

for classURL in URL:
    try:
        returnHttp = s.get(urlHead+classURL).text
        # returnHttp = BeautifulSoup(s.get(urlHead+classURL).text)
    except Exception as e:
        print(e)
        continue
    returnText = (returnHttp).replace(' ', '')
    #f = open('./classExcel.txt', 'w+')
    # f.write(returnText)
    # f.close()
    textList = re.findall(r'>(.+)</span></b>', returnText)
    if len(textList) < 8 or ('1' not in textList):
        print('A closed class')
        continue
    textList = textList[textList.index('1'):]
    # write information to excel
    if '光学' in returnText:
        for teacherName in jiaoshi_optics.keys():
            if teacherName in returnText:
                group = jiaoshi_optics[teacherName]
    else:
        for teacherName in jjiaoshi_modern.keys():
            if teacherName in returnText:
                group = jjiaoshi_modern[teacherName]
    # 创建一个worksheet
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('Worksheet')
    csvName = ['idnumber', 'username', 'lastname', 'firstname',
               'password', 'email', 'institution', 'course1', 'group1']
    # write the framework.
    for i in range(len(csvName)):
        worksheet.write(0, i, label=csvName[i])
    # write the student's information.
    studentName = textList[1::5]
    studentId = textList[2::5]
    # excetion check
    rBuiltState = 0
    for i in studentId:
        # use studentID to check if the student's info is correct
        if not re.findall(r'\d{8}', i):
            rBuiltState = 1
            break
    # excetion occurs
    if rBuiltState:
        studentName = []
        studentId = []
        startNum = 0
        while 1:
            try:
                startNum += 1
                studentName.append(textList[1+textList.index(str(startNum))])
            except:
                break
        emptyText = []
        for i in textList:
            if len(i) == 8:
                emptyText.append(i)
        studentId = re.findall(r'\d{8}', str(textList))
    # new studentId adn studentName is built
    for studentNum in range(len(studentName)):
        worksheet.write(studentNum+1, 0, label=studentId[studentNum])
        worksheet.write(studentNum+1, 1, label=studentId[studentNum])
        worksheet.write(studentNum+1, 3, label=studentName[studentNum])
        worksheet.write(studentNum+1, 4, label=(studentId[studentNum]+'Bjtu@'))
        worksheet.write(studentNum+1, 5,
                        label=(studentId[studentNum]+'Bjtu@.edu.cn'))
        worksheet.write(studentNum+1, 6, label='北京交通大学')
        worksheet.write(studentNum+1, 7, label=course)
        worksheet.write(studentNum+1, 8, label=group)
    # 保存
    # rename
    nameStr = classURL[classURL.index('?')+3:]
    if 'e=3' in nameStr:  # 把T_改为光学的下标
        nameStrNew = 'T_' + \
            nameStr[nameStr.index(
                '2021-')+5:nameStr.index('2021-')+10] + '_C'+str(int(nameStr[-1])+1)
    else:  # 把E_改为近代的下标
        nameStrNew = 'E_' + \
            nameStr[nameStr.index(
                '2021-')+5:nameStr.index('2021-')+10] + '_C'+str(int(nameStr[-1])+1)
    fileName = './{}.'.format(nameStrNew)
    workbook.save(fileName+'xlsx')
    # excel to csv
    import pandas as pd
    data = pd.read_excel(fileName+'xlsx', 'Worksheet', index_col=0)
    f = open(fileName+'csv', 'w+')
    f.close()
    data.to_csv(fileName+'csv', encoding='utf-8')
    os.remove(fileName+'xlsx')
    print('finished {} in {}'.format(
        URL.index(
            classURL), len(URL)
    )
    )