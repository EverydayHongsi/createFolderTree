from difflib import SequenceMatcher
import re
import pandas as pd
import os
#연습용 파일

a_str = 'ASIA PEAR 5'
b_str = '(복사본) ASIA PEARL 5 (2020.4.5)'
c_str = 'ASIA PEARL 6'

a = 'CHANG HAI'
b = 'DONG CHANG HAI'

print(SequenceMatcher(None, a.lower(), b.lower()).ratio())

print(SequenceMatcher(None, a_str, c_str).ratio())

print(b[0:len(a)])


startnum = b_str.lower().find(a_str.lower()[0:2])
print(b_str[startnum:startnum + len(a_str) + 1])

print(SequenceMatcher(None, a_str, b_str[startnum:startnum+len(a_str)+1]).ratio())


string = 'aaa1234, ^&*2233pp'
numbers = re.sub(r'[^0-9]', '', a_str)
print(numbers)



splitword = c_str.split(sep=" ")

for i in splitword:
    try:
        int(i)
        print('굳')
        print(splitword.index(i))
    except Exception as e:
        print(e)

logPD = pd.DataFrame({'Key':[1,2,3], 'Value':[5,6,7]})

if os.path.exists('C:\\Users\\USER\\Desktop\\폴더트리(진행중)\\선박트리\\moveLog.xlsx'):
    #있으면 불러와
    existData = pd.read_excel('C:\\Users\\USER\\Desktop\\폴더트리(진행중)\\선박트리\\moveLog.xlsx', header=0, sheet_name='Sheet1', index_col=None, names = ['Key','Value'])
    #병합해.
    result = pd.concat([existData, logPD])
    result.to_excel('C:\\Users\\USER\\Desktop\\폴더트리(진행중)\\선박트리\\moveLog222.xlsx')