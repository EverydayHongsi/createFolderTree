import os
import pandas as pd
import openpyxl
import datetime
import re
import shutil
from difflib import SequenceMatcher


excelfileLocation = 'C:\\Users\\USER\\Desktop\\선박정보정리_2022_07_08_취합완료.xlsx'

data = pd.read_excel(excelfileLocation, header=None, usecols=[5], sheet_name='데이터원본')

NameList = data[5].values.tolist()

toSavePath = 'C:\\Users\\USER\\Desktop\\폴더트리(진행중)\\선박트리\\'
fileOriginPath = 'C:\\Users\\USER\\Desktop\\폴더트리(진행중)\\stowage_origin\\'

#저장할 폴더 리스트
currentDirDict = {}
logDict = {}
currentFileList = []
moveCompliteFlag = 0

for root, dir, files in os.walk(toSavePath, topdown=False):
    for name in dir:
        currentFolderIncludePath = os.path.join(root, name)
        currentDirDict[name] = currentFolderIncludePath
#저장해야할 파일 리스트를 폴더이름(선박명_콜사인) : 폴더 주소



for root, dir, files in os.walk(fileOriginPath, topdown=False):
    for fullname in files:
        #요상한파일 제외
        if os.path.splitext(fullname)[1].lower() in ['.xlsx', '.xls', '.pdf']:
            filename = os.path.splitext(fullname)[0]
            currentFileIncludePath = os.path.join(root, fullname)
            moveCompliteFlag = 0

            #생성된 폴더명을 가져온다. 즉 넣어야할 선박명을 따올 수 있다.
            for dirname in currentDirDict.keys():
                # 알파벳 폴더 제외
                if len(dirname) == 1:
                    continue
                pureDirName = dirname.split('_')[0]

                #완전히 똑같은 이름을 가지면 넣는다.
                if filename.lower() == pureDirName.lower():
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

            if moveCompliteFlag == 1:
                continue




for root, dir, files in os.walk(fileOriginPath, topdown=False):
    for fullname in files:
        # 요상한파일 제외
        if os.path.splitext(fullname)[1].lower() in ['.xlsx', '.xls', '.pdf']:
            filename = os.path.splitext(fullname)[0]
            currentFileIncludePath = os.path.join(root, fullname)
            moveCompliteFlag = 0
            maxyusado = 0
            lastdirname = ''
            # 생성된 폴더명을 가져온다. 즉 넣어야할 선박명을 따올 수 있다.
            for dirname in currentDirDict.keys():
                # 알파벳 폴더 제외
                if len(dirname) == 1:
                    continue
                pureDirName = dirname.split('_')[0]
                nfilename = filename.replace(" ", "").lower()
                npureDirName = pureDirName.replace(" ", "").lower()
                dirNumbers = re.sub(r'[^0-9]', '', npureDirName)
                fileNumbers = re.sub(r'[^0-9]', '', nfilename[0:len(npureDirName)])
                numberoffile = 0
                splitFileName = filename.split(sep=" ")

                fileNumber = re.sub(r'[^0-9]', '', nfilename)

                if len(fileNumber) > 0:
                    for i in splitFileName:
                        try:
                            int(i)
                            numberoffile = int(splitFileName[splitFileName.index(i)])
                            break
                        except Exception as e:
                            e


                # 숫자가 다르면 스킵코드 다른 숫자들도 있어서 안될듯,, 우선 한글자, 두글자 둘다 같이 있을 때만 비교


                resultYusado = SequenceMatcher(None, nfilename[0:len(npureDirName) + 2], npureDirName).ratio()

                if maxyusado < resultYusado and resultYusado > 0.85:
                    dirNumbers = re.sub(r'[^0-9]', '', npureDirName)

                    if len(dirNumbers) < 4 and numberoffile < 4 and len(dirNumbers) > 0 and numberoffile > 0:
                        if int(dirNumbers) != int(numberoffile):
                            logDict[fullname] = dirname + ': 유사도는 높지만, 숫자가 다르니 않넣습니다.'
                            print(dirname + ":: " + dirNumbers)
                            print(fullname + ":: " + str(numberoffile))
                            continue
                    maxyusado = resultYusado
                    lastdirname = dirname

            # 한자리가 빠지는 건 오타를 중요시여기냐(더 많냐), 아니면 원래 선박들 중 한자리만 다른 것들이 많느냐에 따라 결정하면 됨.

            if maxyusado > 0.85:
                output_path = os.path.join(currentDirDict[lastdirname], fullname)
                uniq = 1
                while os.path.exists(output_path):  # 동일한 파일명이 존재할 때
                    output_path = os.path.join(currentDirDict[lastdirname],
                                               filename + '(%d)' % uniq + os.path.splitext(fullname)[1].lower())

                    uniq += 1
                shutil.move(currentFileIncludePath, output_path)
                logDict[fullname] = lastdirname + ' / 맥스유사도!!!! : ' + str(maxyusado)

#로그 딕셔너리 엑셀로

logPD = pd.DataFrame({'Key': logDict.keys(), 'Value': logDict.values() })

if os.path.exists('C:\\Users\\USER\\Desktop\\폴더트리(진행중)\\선박트리\\moveStowageLog.xlsx'):
    #있으면 불러와
    existData = pd.read_excel('C:\\Users\\USER\\Desktop\\폴더트리(진행중)\\선박트리\\moveStowageLog.xlsx', header=0, sheet_name='Sheet1', index_col=None, names = ['Key', 'Value'])
    #병합해.
    result = pd.concat([existData, logPD])
    result.to_excel('C:\\Users\\USER\\Desktop\\폴더트리(진행중)\\선박트리\\moveStowageLog.xlsx')
    print('병합했니...?')
else:
    logPD.to_excel('C:\\Users\\USER\\Desktop\\폴더트리(진행중)\\선박트리\\moveStowageLog.xlsx')
