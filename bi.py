import xlwings as xw
import pandas as pd
import numpy as np
import cytoolz.curried
import os
import sys

if os.getenv('MY_PYTHON_PKG') not in sys.path:
    sys.path.append(os.getenv('MY_PYTHON_PKG'))

import syspath
from common.connection import conn_local_lite, conn_local_pg
import sqlCommand as sqlc

conn_lite = conn_local_lite('bic.sqlite3')
cur_lite = conn_lite.cursor()


def mymerge(x, y):
    m = pd.merge(x, y, on=[i for i in list(x) if i in list(y)], how='outer')
    return m


os.chdir('C:/Users/ak66h_000/OneDrive/Finance/國發會/bic/')
os.listdir()
table = '景氣指標及燈號-綜合指數'
ext = '.xls'
b = xw.Book(table + ext)
s = b.sheets['Sheet1']
if pd.isnull(s.range('B1').value) == False:
    if pd.isnull(s.range('A1').value):
        s.range('A1').value = '年月'
    if pd.isnull(s.range('A2').value):
        s.range('A2').value = '--'
    l = s.range('A1').expand().value
    col = ['年月'] + l[0][1:]
    df = pd.DataFrame(l[2:], columns=col)
df.insert(0, '年', df.年月.str.split('-').str[0])
df.insert(1, '月', df.年月.str.split('-').str[1])
df.年 = df.年.astype(int)
df.月 = df.月.astype(int)

sql = 'create table `{}` (`{}`, PRIMARY KEY ({}))'.format(table, '`,`'.join(list(df)), '`年`, `月`')
cur_lite.execute(sql)
sql = 'insert into `{}`(`{}`) values({})'.format(table, '`,`'.join(list(df)), ','.join('?'*len(list(df))))
cur_lite.executemany(sql, df.values.tolist())
conn_lite.commit()

table='景氣指標及燈號-指標構成項目'
b = xw.Book(table + ext)
s = b.sheets['Worksheet']
if pd.isnull(s.range('B1').value) == False:
    if pd.isnull(s.range('A1').value):
        s.range('A1').value = '年月'
    if pd.isnull(s.range('A2').value):
        s.range('A2').value = '--'
    l = s.range('A1').expand().value
    col = ['年月'] + l[0][1:]
    df = pd.DataFrame(l[2:], columns=col)
df.insert(0, '年', df.年月.str.split('-').str[0])
df.insert(1, '月', df.年月.str.split('-').str[1])
df.年 = df.年.astype(int)
df.月 = df.月.astype(int)
df = df.replace(',', '', regex=True)
df[[x for x in list(df) if x not in ['年', '月', '年月']]] = df[[x for x in list(df) if x not in ['年', '月', '年月']]].astype(float)

sql = 'create table `{}` (`{}`, PRIMARY KEY ({}))'.format(table, '`,`'.join(list(df)), '`年`, `月`')
cur_lite.execute(sql)
sql = 'insert into `{}`(`{}`) values({})'.format(table, '`,`'.join(list(df)), ','.join('?'*len(list(df))))
cur_lite.executemany(sql, df.values.tolist())
conn_lite.commit()

b = xw.Book('NMI-細項指數.xls')
L = []
for s in b.sheets:
    if pd.isnull(s.range('B1').value)==False:
        if pd.isnull(s.range('A1').value):
            s.range('A1').value='ym'
        if pd.isnull(s.range('A2').value):
            s.range('A2').value='--'
        l = s.range('A1').expand().value
        col = ['ym']+l[0][1:]
        df = pd.DataFrame(l[2:], columns=['ym']+l[0][1:])
        df['industry'] = str(s).split(']')[1].split('>')[0]
        L.append(df)
df = reduce(mymerge, L)
print(df)

for i in b.sheets:
    print(i.range('A3').expand().value)

b.sheets[0].range('A3').value
pd.isnull(s.range('A1').value) == False
str(b.sheets[0]).split(']')[1].split('>')[0]