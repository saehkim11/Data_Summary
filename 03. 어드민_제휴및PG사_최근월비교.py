# ---
# jupyter:
#   jupytext:
#     formats: ipynb,py:light
#     text_representation:
#       extension: .py
#       format_name: light
#       format_version: '1.5'
#       jupytext_version: 1.13.8
#   kernelspec:
#     display_name: Python 3
#     language: python
#     name: python3
# ---

# +
import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import datetime as dt
import missingno as msno
# %matplotlib inline
import matplotlib as mpl
import matplotlib.font_manager as fm

mpl.rcParams['axes.unicode_minus'] = False

path = 'C:/Windows/Fonts/malgun.ttf'
font_name = fm.FontProperties(fname=path, size=50).get_name()
plt.rc('font', family=font_name)

import time
from tqdm import tqdm
from tqdm import trange

pd.set_option('display.max_columns', 1000)
pd.set_option('display.max_rows', 1000)

import warnings
warnings.filterwarnings('ignore')
# -


# +
os.chdir('D:\\데이터\\영업관리프로그램\\영업관리프로그램(16년을 제외한 인하펀칭 전 금액 반영)')

df = pd.DataFrame()
data = pd.read_excel('2022_자사어드민_인하펀칭조정.xlsx', sheet_name=None)
data = pd.concat(data)
df = df.append(data)
df = df.reset_index()

# +
월 = df['level_0'].unique().tolist()
연월 = ['2022-01','2022-02','2022-03','2022-04','2022-05']

mon = pd.DataFrame({
    '월' : 월,
    '연월' : 연월
})

df['귀속월'] = df['level_0'].map(mon.set_index('월')['연월'])
df.drop(['level_0', 'level_1'],axis=1, inplace=True)
어드민 = df.copy()

어드민['제휴몰 상세번호'] = 어드민['제휴몰 상세번호'].astype(str).apply(lambda x:x.replace('.0', ''))
어드민.loc[어드민['판매처'].str.contains('GS'), '제휴몰 상세번호'] = 어드민.loc[어드민['판매처'].str.contains('GS'), '제휴몰 상세번호'].astype(str).str.split('0').str[0]

어드민['주문_상세'] = 어드민['제휴몰 주문번호'].astype(str) + '-' + 어드민['제휴몰 상세번호'].astype(str)
cols = ['매출일자','판매처','거래구분','품번','인하펀칭 전 금액','자사몰 주문번호','자사몰 상세번호','제휴몰 주문번호','제휴몰 상세번호','귀속월','주문_상세']
어드민 = 어드민[cols]
# -

어드민_pvt = pd.pivot_table(어드민, index='주문_상세', values='인하펀칭 전 금액', aggfunc='sum').reset_index()
어드민_pvt['판매처'] = 어드민_pvt['주문_상세'].map(어드민.drop_duplicates('주문_상세').set_index('주문_상세')['판매처'])

# # ---------------------------------------------------------------------

# # 현금입금건 가져오기

# +
현금 = pd.DataFrame()

os.chdir('D:\\매출분석\\온라인\\자사몰_제휴사별_분석')
data = pd.read_excel('온라인_현금입금.xlsx', sheet_name=None, skiprows=1)
data = pd.concat(data)
현금 = 현금.append(data)
현금 = 현금.reset_index()
현금 = 현금[현금['주문번호'].notnull()]
현금 = 현금.drop(['level_1', 'Unnamed: 1'], axis=1)
현금 = 현금.rename(columns={'level_0':'귀속월',
                  'Unnamed: 0':'날짜',
                  'Unnamed: 2':'입금자',
                  'Unnamed: 3':'지점명',
                  'Unnamed: 4':'브랜드'})
현금.loc[현금['제휴사 주문번호'].astype(str).str.contains('~'), '제휴사 주문번호'] = 현금[현금['제휴사 주문번호'].astype(str).str.contains('~')]['제휴사 주문번호'].str.split('~').str[0].str.strip()
# -

현금_pvt = pd.pivot_table(현금, index='제휴사 주문번호', values='판매금액', aggfunc='sum').reset_index()
현금_pvt['판매처'] = 현금_pvt['제휴사 주문번호'].map(현금.drop_duplicates('제휴사 주문번호').set_index('제휴사 주문번호')['판매처'])
현금_pvt = 현금_pvt.rename(columns = {'제휴사 주문번호' : '주문_상세', '판매금액':'현금입금'})

# # -----------------------------------------------------

# # 1. 인터파크 비교

어드민_인터 = 어드민[어드민['판매처'].str.contains('인터파크')]
어드민_인터_pvt = pd.pivot_table(어드민_인터, index='주문_상세', values='인하펀칭 전 금액', aggfunc='sum').reset_index()

os.chdir('D:\\데이터\\제휴몰 자료\\인터파크')
인터_PG = pd.DataFrame()
data = pd.read_excel('인터파크2022년 1~5월.xlsx',skiprows=11 ,sheet_name=None)
data = pd.concat(data, ignore_index=True)
인터_PG = 인터_PG.append(data)
인터_PG = 인터_PG[인터_PG['구분'] != '합계']
인터_PG['주문순번'] = 인터_PG['주문순번'].astype(str).apply(lambda x : x.replace('.0', ''))
인터_PG['주문_상세'] = 인터_PG['주문번호'] + '-' + 인터_PG['주문순번']
인터_PG_pvt = pd.pivot_table(인터_PG, index='주문_상세', values='자료_총판매금액', aggfunc='sum').reset_index()
인터_현금 = 현금_pvt[현금_pvt['판매처'] == '인터파크']

from functools import reduce
dfs = [어드민_인터_pvt,인터_PG_pvt, 인터_현금]
인터_mg = reduce(lambda left, right: pd.merge(left, right, on='주문_상세', how='left'),dfs)

# +
인터_mg.loc[인터_mg['자료_총판매금액'].notnull(), '차이'] = 인터_mg['인하펀칭 전 금액'] - 인터_mg['자료_총판매금액']
인터_mg.loc[인터_mg['현금입금'].notnull(), '차이'] = 인터_mg['인하펀칭 전 금액'] - 인터_mg['현금입금']

인터_mg['귀속월'] = 인터_mg['주문_상세'].str[0:6]
인터_mg.loc[인터_mg['자료_총판매금액'].notnull(), '차이'] = 인터_mg['인하펀칭 전 금액'] - 인터_mg['자료_총판매금액']
인터_mg.loc[인터_mg['현금입금'].notnull(), '차이'] = 인터_mg['인하펀칭 전 금액'] - 인터_mg['현금입금']

# +
# 현재 상태 카테고리화 하기
인터_mg.loc[(인터_mg['인하펀칭 전 금액'] == 0) & (인터_mg['자료_총판매금액'].isnull()), 'T/N'] = 'True'
인터_mg.loc[인터_mg['차이'] == 0, 'T/N'] = 'True'

인터_mg.loc[(인터_mg['인하펀칭 전 금액'] < 0) & (인터_mg['자료_총판매금액'].isnull()), 'T/N'] = '초기'
인터_mg.loc[(인터_mg['인하펀칭 전 금액'] == 0) & (인터_mg['자료_총판매금액'] > 0), 'T/N'] = '초기'

인터_mg.loc[(인터_mg['인하펀칭 전 금액'] > 0) & (인터_mg['자료_총판매금액'].isnull())& (인터_mg['현금입금'].isnull()), 'T/N'] = '미정산'

인터_mg.loc[인터_mg['T/N'].isnull() & (인터_mg['차이'] != 0), 'T/N'] = 'False'




# # 2. 이베이 비교

어드민_이베이 = 어드민[어드민['판매처'].str.contains('옥션|지마켓')]
어드민_이베이_pvt = pd.pivot_table(어드민_이베이, index='주문_상세', values='인하펀칭 전 금액', aggfunc='sum').reset_index()
print(len(어드민_이베이_pvt))

# +
os.chdir('D:\\데이터\\제휴몰 자료\\이베이\\옥션')
옥_PG = pd.DataFrame()
data = pd.read_excel('2021.09~2022.05_옥션.xlsx', sheet_name=None)
data = pd.concat(data)
옥_PG = 옥_PG.append(data)
옥_PG = 옥_PG.reset_index()
옥_PG.rename(columns = {'level_0':'기준월'}, inplace=True)

옥_PG['주문_상세'] = 옥_PG['결제번호'].astype(str) + '-' + 옥_PG['주문번호'].astype(str)
옥_PG['총_결제금액'] = 옥_PG['상품 판매가'] + 옥_PG['옵션상품 판매가']

옥_PG_pvt = pd.pivot_table(옥_PG, index='주문_상세', values= '총_결제금액',aggfunc='sum').reset_index()
옥_PG_pvt['구분'] = '옥션'

# +
os.chdir('D:\\데이터\\제휴몰 자료\\이베이\\지마켓')
G_PG = pd.DataFrame()
data = pd.read_excel('2021.09~2022.05_지마켓.xlsx', sheet_name=None, skiprows=1)
data = pd.concat(data)
G_PG = G_PG.append(data)
G_PG = G_PG.reset_index()
G_PG.rename(columns = {'level_0':'기준월'}, inplace=True)
G_PG = G_PG[G_PG['장바구니번호'].notnull()]
G_PG['주문_상세'] = G_PG['장바구니번호'].astype(str).apply(lambda x: x.replace('.0', '')) + '-' + \
G_PG['주문번호'].astype(str).apply(lambda x: x.replace('.0', ''))

G_PG['총_결제금액'] = G_PG['판매가격'] + G_PG['필수선택상품금액']

G_PG_pvt = pd.pivot_table(G_PG, index='주문_상세', values='총_결제금액', aggfunc='sum').reset_index()
G_PG_pvt['구분'] = '지마켓'
# -

이_PG_pvt = pd.concat([옥_PG_pvt, G_PG_pvt])

# # 이베이 현금 입금분

이베_현금 = 현금_pvt[(현금_pvt['판매처'] == '옥션') | (현금_pvt['판매처'] == '지마켓')]
이베_mg = pd.merge(어드민_이베이_pvt, 이_PG_pvt, on='주문_상세',how='left')
이베_mg['구분'] = 이베_mg['주문_상세'].map(어드민_이베이.drop_duplicates('주문_상세').set_index('주문_상세')['판매처'])


for i in 이베_현금['주문_상세']:
    이베_mg.loc[이베_mg['주문_상세'].str.contains(i),'현금입금'] = 이베_현금.loc[이베_현금['주문_상세'].str.contains(i),'현금입금'].values[0]
    이베_mg.loc[이베_mg['주문_상세'].str.contains(i),'구분'] = 이베_현금.loc[이베_현금['주문_상세'].str.contains(i),'판매처'].values[0]

# +
이베_mg['귀속월'] = 이베_mg['주문_상세'].map(어드민_이베이.drop_duplicates('주문_상세').set_index('주문_상세')['매출일자'].astype(str).str[0:7])

이베_mg['제휴월'] = 이베_mg['주문_상세'].map(옥_PG.drop_duplicates('주문_상세').set_index('주문_상세')['기준월'])
이베_mg['제휴월2'] = 이베_mg['주문_상세'].map(G_PG.drop_duplicates('주문_상세').set_index('주문_상세')['기준월'])
이베_mg.loc[이베_mg['제휴월'].isnull(), '제휴월'] = 이베_mg.loc[이베_mg['제휴월'].isnull(), '제휴월2'] 
이베_mg.drop('제휴월2', axis=1,inplace=True)
이베_mg.loc[이베_mg['총_결제금액'].notnull(), '차이'] = 이베_mg['인하펀칭 전 금액'] - 이베_mg['총_결제금액']
이베_mg.loc[이베_mg['현금입금'].notnull(), '차이'] = 이베_mg['인하펀칭 전 금액'] - 이베_mg['현금입금']

# +
# 현재 상태 카테고리화 하기

이베_mg.loc[(이베_mg['인하펀칭 전 금액'] == 0) & (이베_mg['총_결제금액'].isnull()), 'T/N'] = 'True'
이베_mg.loc[이베_mg['차이'] == 0, 'T/N'] = 'True'

이베_mg.loc[(이베_mg['인하펀칭 전 금액'] < 0) & (이베_mg['총_결제금액'].isnull()), 'T/N'] = '초기'
이베_mg.loc[(이베_mg['인하펀칭 전 금액'] < 0) & (이베_mg['총_결제금액'] == 0), 'T/N'] = '초기'
이베_mg.loc[(이베_mg['인하펀칭 전 금액'] == 0) & (이베_mg['총_결제금액'] > 0), 'T/N'] = '초기'

이베_mg.loc[(이베_mg['인하펀칭 전 금액'] > 0) & (이베_mg['총_결제금액'].isnull())& (이베_mg['현금입금'].isnull()), 'T/N'] = '미정산'

이베_mg.loc[이베_mg['T/N'].isnull() & (이베_mg['차이'] != 0), 'T/N'] = 'False'

# +

# # 3. 쿠팡 비교

어드민_쿠팡 = 어드민[어드민['판매처'].str.contains('쿠팡')]
어드민_쿠팡_pvt = pd.pivot_table(어드민_쿠팡, index='주문_상세', values='인하펀칭 전 금액', aggfunc='sum').reset_index()

# +
os.chdir('D:\\데이터\\제휴몰 자료\\쿠팡\\2022년')
files = os.listdir()
df = pd.DataFrame()

for f in enumerate(tqdm(files)):
    data = pd.read_excel(f[1], sheet_name=None)
    data = pd.concat(data, ignore_index=True)
    df = df.append(data)
쿠팡 = df.copy()
# -

쿠팡 = 쿠팡[쿠팡['Product ID'].notnull()]
쿠팡['주문_상세'] = 쿠팡['주문번호'].astype(str) + '-' + 쿠팡['Option ID'].astype(str)
쿠팡_pvt = pd.pivot_table(쿠팡, index='주문_상세', values='매출금액(D=A-B)', aggfunc='sum').reset_index()

# # 쿠팡 현금입금분

쿠팡_현금 = 현금_pvt[현금_pvt['판매처'] == '쿠팡']

# +
쿠팡_mg = pd.merge(어드민_쿠팡_pvt, 쿠팡_pvt, on='주문_상세', how='left')

for i in 쿠팡_현금['주문_상세']:
    쿠팡_mg.loc[쿠팡_mg['주문_상세'].str.contains(i), '현금입금'] = 쿠팡_현금.loc[쿠팡_현금['주문_상세'].str.contains(i),'현금입금'].values[0]
# -

쿠팡_mg['귀속월'] = 쿠팡_mg['주문_상세'].map(어드민.drop_duplicates('주문_상세').set_index('주문_상세')['귀속월'])

쿠팡_mg.loc[쿠팡_mg['매출금액(D=A-B)'].notnull(), '차이'] = 쿠팡_mg['인하펀칭 전 금액'] - 쿠팡_mg['매출금액(D=A-B)']
쿠팡_mg.loc[쿠팡_mg['현금입금'].notnull(), '차이'] = 쿠팡_mg['인하펀칭 전 금액'] - 쿠팡_mg['현금입금']

# +
# 현재 상태 카테고리화 하기
쿠팡_mg.loc[(쿠팡_mg['인하펀칭 전 금액'] == 0) & (쿠팡_mg['매출금액(D=A-B)'].isnull()), 'T/N'] = 'True'
쿠팡_mg.loc[쿠팡_mg['차이'] == 0, 'T/N'] = 'True'

쿠팡_mg.loc[(쿠팡_mg['인하펀칭 전 금액'] < 0) & (쿠팡_mg['매출금액(D=A-B)'].isnull()), 'T/N'] = '초기'
쿠팡_mg.loc[(쿠팡_mg['인하펀칭 전 금액'] < 0) & (쿠팡_mg['매출금액(D=A-B)'] == 0), 'T/N'] = '초기'
쿠팡_mg.loc[(쿠팡_mg['인하펀칭 전 금액'] == 0) & (쿠팡_mg['매출금액(D=A-B)'] > 0), 'T/N'] = '초기'

쿠팡_mg.loc[(쿠팡_mg['인하펀칭 전 금액'] > 0) & (쿠팡_mg['매출금액(D=A-B)'].isnull())& (쿠팡_mg['현금입금'].isnull()), 'T/N'] = '미정산'

쿠팡_mg.loc[쿠팡_mg['T/N'].isnull() & (쿠팡_mg['차이'] != 0), 'T/N'] = 'False'

# +



# # 4. 패션플러스 비교

어드민_패플 = 어드민[어드민['판매처'].str.contains('패션플러스')]
어드민_패플_pvt = pd.pivot_table(어드민_패플, index='주문_상세', values='인하펀칭 전 금액', aggfunc='sum').reset_index()
print(len(어드민_패플_pvt))

# +
os.chdir('D:\\데이터\\제휴몰 자료\\패션플러스\\2022년')
files = os.listdir()
df = pd.DataFrame()

for f in enumerate(tqdm(files)):
    data = pd.read_excel(f[1], sheet_name=None)
    data = pd.concat(data, ignore_index=True)
    df = df.append(data)
패플 = df.copy()

# +
패플 = 패플[패플['주문번호'].notnull()]
패플.loc[패플['주문유형'] == '취소반품주문', '주문번호'] = 패플.loc[패플['주문유형'] == '취소반품주문', '이전주문번호'] 

패플_pvt = pd.pivot_table(패플, index='주문번호', values='판매금액', aggfunc='sum').reset_index()
패플_pvt = 패플_pvt.rename(columns={'주문번호':'주문_상세'})
# -

# # 패플 현금입금분

패플_현금 = 현금_pvt[현금_pvt['판매처'] == '패션플러스']

dfs = [어드민_패플_pvt, 패플_pvt,패플_현금]
패플_mg = reduce(lambda left, right : pd.merge(left, right, on='주문_상세', how='left'), dfs)

패플_mg['귀속월'] = 패플_mg['주문_상세'].map(어드민_패플.drop_duplicates('주문_상세').set_index('주문_상세')['귀속월'])
패플_mg.loc[패플_mg['판매금액'].notnull(), '차이'] = 패플_mg['인하펀칭 전 금액'] - 패플_mg['판매금액']
패플_mg.loc[패플_mg['현금입금'].notnull(), '차이'] = 패플_mg['인하펀칭 전 금액'] - 패플_mg['현금입금']

# +
# 현재 상태 카테고리화 하기
패플_mg.loc[(패플_mg['인하펀칭 전 금액'] == 0) & (패플_mg['판매금액'].isnull()), 'T/N'] = 'True'
패플_mg.loc[패플_mg['차이'] == 0, 'T/N'] = 'True'

패플_mg.loc[(패플_mg['인하펀칭 전 금액'] < 0) & (패플_mg['판매금액'].isnull()), 'T/N'] = '초기'
패플_mg.loc[(패플_mg['인하펀칭 전 금액'] < 0) & (패플_mg['판매금액'] == 0), 'T/N'] = '초기'
패플_mg.loc[(패플_mg['인하펀칭 전 금액'] == 0) & (패플_mg['판매금액'] > 0), 'T/N'] = '초기'

패플_mg.loc[(패플_mg['인하펀칭 전 금액'] > 0) & (패플_mg['판매금액'].isnull())& (패플_mg['현금입금'].isnull()), 'T/N'] = '미정산'

패플_mg.loc[패플_mg['T/N'].isnull() & (패플_mg['차이'] != 0), 'T/N'] = 'False'
# -


# # ------------------------------------------------------

# # 패플 주문번호 기준

# +
어드민_패플 = 어드민[어드민['판매처'].str.contains('패션플러스')]

어드민_패플_pvt = pd.pivot_table(어드민_패플, index='제휴몰 주문번호', values='인하펀칭 전 금액', aggfunc='sum').reset_index()
# -

패플 = 패플[패플['주문번호'].notnull()]
패플.loc[패플['주문유형'] == '취소반품주문', '주문번호'] = 패플.loc[패플['주문유형'] == '취소반품주문', '이전주문번호'] 

패플['제휴몰 주문번호'] = 패플['주문번호'].str.split('-').str[0]

패플_pvt = pd.pivot_table(패플, index='제휴몰 주문번호', values='판매금액', aggfunc='sum').reset_index()

패플_현금 = 현금_pvt[현금_pvt['판매처'] == '패션플러스']
패플_현금['제휴몰 주문번호'] = 패플_현금['주문_상세'].str.split('-').str[0]

dfs = [어드민_패플_pvt, 패플_pvt,패플_현금]
패플_mg = reduce(lambda left, right : pd.merge(left, right, on='제휴몰 주문번호', how='left'), dfs)

패플_mg['귀속월'] = 패플_mg['제휴몰 주문번호'].map(어드민_패플.drop_duplicates('제휴몰 주문번호').set_index('제휴몰 주문번호')['귀속월'])
패플_mg.loc[패플_mg['판매금액'].notnull(), '차이'] = 패플_mg['인하펀칭 전 금액'] - 패플_mg['판매금액']
패플_mg.loc[패플_mg['현금입금'].notnull(), '차이'] = 패플_mg['인하펀칭 전 금액'] - 패플_mg['현금입금']

# +
# 현재 상태 카테고리화 하기
패플_mg.loc[(패플_mg['인하펀칭 전 금액'] == 0) & (패플_mg['판매금액'].isnull()), 'T/N'] = 'True'
패플_mg.loc[패플_mg['차이'] == 0, 'T/N'] = 'True'

패플_mg.loc[(패플_mg['인하펀칭 전 금액'] < 0) & (패플_mg['판매금액'].isnull()), 'T/N'] = '초기'
패플_mg.loc[(패플_mg['인하펀칭 전 금액'] < 0) & (패플_mg['판매금액'] == 0), 'T/N'] = '초기'
패플_mg.loc[(패플_mg['인하펀칭 전 금액'] == 0) & (패플_mg['판매금액'] > 0), 'T/N'] = '초기'

패플_mg.loc[(패플_mg['인하펀칭 전 금액'] > 0) & (패플_mg['판매금액'].isnull())& (패플_mg['현금입금'].isnull()), 'T/N'] = '미정산'

패플_mg.loc[패플_mg['T/N'].isnull() & (패플_mg['차이'] != 0), 'T/N'] = 'False'
# -





# # 5. 트라이씨클, 하프클럽 비교

어드민_하프 = 어드민[어드민['판매처'].str.contains('하프클럽')]
어드민_하프_pvt = pd.pivot_table(어드민_하프, index='주문_상세', values='인하펀칭 전 금액', aggfunc='sum').reset_index()
print(len(어드민_하프_pvt))

# +
os.chdir('D:\\데이터\\제휴몰 자료\\하프클럽\\2022년')
files = os.listdir()
df = pd.DataFrame()

for f in enumerate(tqdm(files)):
    data = pd.read_excel(f[1], sheet_name=None)
    data = pd.concat(data, ignore_index=True)
    df = df.append(data)
하프 = df.copy()
# -

하프['주문_상세'] = 하프['주문번호'].astype(str) +'-'+ 하프['순번'].astype(str)
하프_pvt = pd.pivot_table(하프, index='주문_상세', values='실출고금액', aggfunc='sum').reset_index()

# # 하프클럽 현금입급분

하프_현금 = 현금_pvt[현금_pvt['판매처'] == '하프클럽']
하프_현금.head(1)

dfs = [어드민_하프_pvt, 하프_pvt, 하프_현금]
하프_mg = reduce(lambda left, right : pd.merge(left, right, on='주문_상세', how='left'), dfs)

하프_mg['귀속월'] = 하프_mg['주문_상세'].map(어드민_하프.drop_duplicates('주문_상세').set_index('주문_상세')['귀속월'])
# 하프_mg['차이'] = 하프_mg['인하펀칭 전 금액'] - 하프_mg['실출고금액']
하프_mg.loc[하프_mg['실출고금액'].notnull(), '차이'] = 하프_mg['인하펀칭 전 금액'] - 하프_mg['실출고금액']
하프_mg.loc[하프_mg['현금입금'].notnull(), '차이'] = 하프_mg['인하펀칭 전 금액'] - 하프_mg['현금입금']

# +
# 현재 상태 카테고리화 하기
하프_mg.loc[(하프_mg['인하펀칭 전 금액'] == 0) & (하프_mg['실출고금액'].isnull()), 'T/N'] = 'True'
하프_mg.loc[하프_mg['차이'] == 0, 'T/N'] = 'True'

하프_mg.loc[(하프_mg['인하펀칭 전 금액'] < 0) & (하프_mg['실출고금액'].isnull()), 'T/N'] = '초기'
하프_mg.loc[(하프_mg['인하펀칭 전 금액'] < 0) & (하프_mg['실출고금액'] == 0), 'T/N'] = '초기'
하프_mg.loc[(하프_mg['인하펀칭 전 금액'] == 0) & (하프_mg['실출고금액'] > 0), 'T/N'] = '초기'

하프_mg.loc[(하프_mg['인하펀칭 전 금액'] > 0) & (하프_mg['실출고금액'].isnull())& (하프_mg['현금입금'].isnull()), 'T/N'] = '미정산'

하프_mg.loc[하프_mg['T/N'].isnull() & (하프_mg['차이'] != 0), 'T/N'] = 'False'
# -




# # 6. GS홈쇼핑 비교

어드민_GS = 어드민[어드민['판매처'].str.contains('GS')]
어드민_GS_pvt = pd.pivot_table(어드민_GS, index='주문_상세', values='인하펀칭 전 금액', aggfunc='sum').reset_index()
어드민_GS_pvt
print(len(어드민_GS_pvt))

# +
os.chdir('D:\\데이터\\제휴몰 자료\\GS\\2022년')
files = os.listdir()
df = pd.DataFrame()

for f in enumerate(tqdm(files)):
    data = pd.read_excel(f[1], sheet_name=None)
    data = pd.concat(data, ignore_index=True)
    df = df.append(data)
GS = df.copy()
# -

GS.loc[GS['주문유형'] != '정상주문', '주문번호'] = GS.loc[GS['주문유형'] != '정상주문', '원주문번호']
GS['주문번호'] = GS['주문번호'].astype(str).apply(lambda x : x.replace('.0', ''))
GS['주문_상세'] = GS['주문번호'].astype(str) + '-' + GS['품목'].astype(str)

GS_pvt = pd.pivot_table(GS, index='주문_상세', values='판매가격', aggfunc='sum').reset_index()

# # GS 현금입금분

GS_현금 = 현금_pvt[현금_pvt['판매처'] == 'GS']
GS_현금.head(1)

# +

GS_mg = pd.merge(어드민_GS_pvt, GS_pvt,on='주문_상세', how='left')
GS_mg['귀속월'] = GS_mg['주문_상세'].map(어드민_GS.drop_duplicates('주문_상세').set_index('주문_상세')['귀속월'])
for i in GS_현금['주문_상세']:
    GS_mg.loc[GS_mg['주문_상세'].str.contains(i), '현금입금'] = GS_현금.loc[GS_현금['주문_상세'].str.contains(i), '현금입금'].values[0]
# -

GS_mg['귀속월'] = GS_mg['주문_상세'].map(어드민_GS.drop_duplicates('주문_상세').set_index('주문_상세')['귀속월'])
GS_mg.loc[GS_mg['판매가격'].notnull(), '차이']= GS_mg['인하펀칭 전 금액'] - GS_mg['판매가격']
GS_mg.loc[GS_mg['현금입금'].notnull(), '차이']= GS_mg['인하펀칭 전 금액'] - GS_mg['현금입금']

# +
# 현재 상태 카테고리화 하기
GS_mg.loc[(GS_mg['인하펀칭 전 금액'] == 0) & (GS_mg['판매가격'].isnull()), 'T/N'] = 'True'
GS_mg.loc[GS_mg['차이'] == 0, 'T/N'] = 'True'

GS_mg.loc[(GS_mg['인하펀칭 전 금액'] < 0) & (GS_mg['판매가격'].isnull()), 'T/N'] = '초기'
GS_mg.loc[(GS_mg['인하펀칭 전 금액'] < 0) & (GS_mg['판매가격'] == 0), 'T/N'] = '초기'
GS_mg.loc[(GS_mg['인하펀칭 전 금액'] == 0) & (GS_mg['판매가격'] > 0), 'T/N'] = '초기'

GS_mg.loc[(GS_mg['인하펀칭 전 금액'] > 0) & (GS_mg['판매가격'].isnull())& (GS_mg['현금입금'].isnull()), 'T/N'] = '미정산'

GS_mg.loc[GS_mg['T/N'].isnull() & (GS_mg['차이'] != 0), 'T/N'] = 'False'
# -

c1 = GS_mg['귀속월'] != '2022-05'
c2 = GS_mg['T/N'] == '미정산'

# 조회대상 = 어드민_GS[어드민_GS['주문_상세'].isin(GS_mg[c1 & c2]['주문_상세'])]
조회대상 = 어드민_GS[어드민_GS['제휴몰 주문번호'].isin(GS_mg[c1 & c2]['주문번호'])]
os.chdir('D:\\김세환 백업\\Python_Files\\매출분석\\온라인\\자사몰_자동화')
조회대상.to_excel('GS_조회대상.xlsx',index=False)


# ### GS 변경 후 번호 가져오기

os.chdir('D:\\온라인\\자사몰_자동화')
bf = pd.read_excel('GS_조회결과.xlsx')
bf['변경후'] = bf['변경후'].astype(str).str.split('.').str[0]
bf['변경후'] = bf['변경후'].astype(str)
bf['제휴몰_주문번호'] = bf['제휴몰_주문번호'].astype(str)

GS_mg['변경후'] = GS_mg['주문번호'].map(bf.drop_duplicates('제휴몰_주문번호').set_index('제휴몰_주문번호')['변경후'])
GS_mg['변경후_금액'] = GS_mg['변경후'].map(GS_pvt.set_index('주문번호')['판매가격'])

GS_mg.loc[GS_mg['판매가격'].notnull(), '차이']= GS_mg['인하펀칭 전 금액'] - GS_mg['판매가격']
GS_mg.loc[GS_mg['현금입금'].notnull(), '차이']= GS_mg['인하펀칭 전 금액'] - GS_mg['현금입금']
GS_mg.loc[GS_mg['변경후_금액'].notnull(), '차이'] = GS_mg['인하펀칭 전 금액'] - GS_mg['변경후_금액']

# +
# 현재 상태 카테고리화 하기
GS_mg.loc[(GS_mg['인하펀칭 전 금액'] == 0) & (GS_mg['판매가격'].isnull()), 'T/N'] = 'True'
GS_mg.loc[GS_mg['차이'] == 0, 'T/N'] = 'True'

GS_mg.loc[(GS_mg['인하펀칭 전 금액'] < 0) & (GS_mg['판매가격'].isnull()), 'T/N'] = '초기'
GS_mg.loc[(GS_mg['인하펀칭 전 금액'] < 0) & (GS_mg['판매가격'] == 0), 'T/N'] = '초기'
GS_mg.loc[(GS_mg['인하펀칭 전 금액'] == 0) & (GS_mg['판매가격'] > 0), 'T/N'] = '초기'

GS_mg.loc[(GS_mg['인하펀칭 전 금액'] > 0) & (GS_mg['판매가격'].isnull()) \
          & (GS_mg['현금입금'].isnull()) & (GS_mg['변경후_금액'].isnull()), 'T/N'] = '미정산'

GS_mg.loc[GS_mg['T/N'].isnull() & (GS_mg['차이'] != 0), 'T/N'] = 'False'
# -

GS_mg[(GS_mg['T/N'] == '미정산') & (GS_mg['귀속월'] != '2022-05')]



# ## 내용요약, 하나의 파일로 작성하기

# +
writer = pd.ExcelWriter(f'D:\\매출분석\\온라인\\자사몰_제휴사별_분석\\제휴몰_미정산_리스트.xlsx', engine='xlsxwriter')

인터_mg.to_excel(writer, sheet_name='인터파크')
이베_mg.to_excel(writer, sheet_name='이베이')
쿠팡_mg.to_excel(writer, sheet_name='쿠팡')
패플_mg.to_excel(writer, sheet_name='패션플러스')
하프_mg.to_excel(writer, sheet_name='하프클럽')
GS_mg.to_excel(writer, sheet_name='GS')

writer.save()
# -

