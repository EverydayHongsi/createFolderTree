import os
import pandas as pd
import openpyxl
import datetime
import re
import shutil

excelfileLocation = '/Users/hongsi/Desktop/선박정보정리_2022.06.27완.xlsx'

data = pd.read_excel(excelfileLocation, header=None, usecols=[6], sheet_name='데이터원본')

NameList = data[6].values.tolist()

toSavePath = '/Users/hongsi/Desktop/선박트리/'
fileOriginPath = '/Users/hongsi/Desktop/origin/'

#저장할 폴더 리스트
currentDirDict = {}
logDict = {}
currentFileList = []
moveCompliteFlag = 0

for root, dir, files in os.walk(toSavePath, topdown=False):
    for name in dir:
        currentFolderIncludePath = os.path.join(root, name)
        currentDirDict[name] = currentFolderIncludePath

#저장해야할 파일 리스트


for root, dir, files in os.walk(fileOriginPath, topdown=False):
    for fullname in files:
        #요상한파일 제외
        if os.path.splitext(fullname)[1].lower() in ['.png', '.pdf', '.tif', '.tiff']:
            filename = os.path.splitext(fullname)[0]
            currentFileIncludePath = os.path.join(root, fullname)
            moveCompliteFlag = 0

            for dirname in currentDirDict.keys():
                # 알파벳 폴더 제외
                if len(dirname) == 1:
                    continue
                pureDirName = dirname.split('_')[0]

                if filename == pureDirName:
                    output_path = os.path.join(currentDirDict[dirname], fullname)

                    uniq = 1
                    while os.path.exists(output_path):  # 동일한 파일명이 존재할 때
                        output_path = os.path.join(currentDirDict[dirname], filename + '(%d)' % uniq + os.path.splitext(fullname)[1].lower() )
                        uniq += 1
                    shutil.move(currentFileIncludePath, output_path)
                    #로그 남기기
                    logDict[fullname] = dirname

                    moveCompliteFlag = 1
                    break
                else:

                    nfilename = filename.replace(" ", "").lower()
                    npureDirName = pureDirName.replace(" ", "").lower()

                    if npureDirName in nfilename:
                        output_path = os.path.join(currentDirDict[dirname], fullname)
                        uniq = 1
                        while os.path.exists(output_path):  # 동일한 파일명이 존재할 때
                            output_path = os.path.join(currentDirDict[dirname],
                                                       filename + '(%d)' % uniq + os.path.splitext(fullname)[1].lower())

                            uniq += 1
                        shutil.move(currentFileIncludePath, output_path)
                        logDict[fullname] = dirname
                        moveCompliteFlag = 1
                        break

            if moveCompliteFlag == 1:
                continue


#로그 딕셔너리 엑셀로

logPD = pd.DataFrame({'Key':logDict.keys(), 'Value': logDict.values() })
logPD.to_excel('moveLog.xlsx')