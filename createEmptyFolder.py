import os
import pandas as pd
import openpyxl
import datetime
import re

def createlog(path, name, todata):
    if os.path.exists(path + name + '.txt'):
        f = open(path + name + '.txt', 'a')
        f.write(todata + '\n')
        f.close()
    else:
        f = open(path + name + '.txt', 'w')
        f.write(todata + '\n')
        f.close()

excelfileLocation = '/Users/hongsi/Desktop/선박정보정리_2022.06.27완.xlsx'

data = pd.read_excel(excelfileLocation, header=None, usecols=[5], sheet_name='데이터원본')

#dataframe
print(type(data))
#Series
print(type(data[5]))

listFolderName = data[5].values.tolist()

#원하는것만 추출하기
listFolderName = listFolderName[5:200]
print(listFolderName)

#폴더만들 경로
resultPath = '/Users/hongsi/Desktop/선박트리/'

#기존에 있는 폴더명 먼저 받기
currentdirList = []
for root, dir, files in os.walk(resultPath, topdown=False):
    for name in dir:
        currentdirList.append(name)
print(currentdirList)

for folder in listFolderName:
    if folder in currentdirList:
        log = '{} 폴더가 이미 존재합니다. 실행날짜 : {}'.format(folder, datetime.datetime.now())
        print(log)
        createlog(resultPath, 'existFileLog', log)
        continue

    # 쓰레기값 제거
    if re.search('[a-zA-z]', folder[0]) is None:
        log = '파일 맨 앞 문자가 알파벳이 아닙니다.' + ' : ' + folder + ' time : ' + str(datetime.datetime.now())
        createlog(resultPath, 'exception', log)
        continue

    if folder[0] not in os.listdir(resultPath):
        os.mkdir(resultPath + folder[0])

    try:
        os.mkdir(resultPath + folder[0].lower() + '/' + folder)
    except Exception as e:
        print(e)
        log = str(e) + ' : ' + folder + ' time : ' + str(datetime.datetime.now())
        createlog(resultPath, 'exception', log)





