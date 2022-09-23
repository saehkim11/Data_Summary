import pandas as pd
import os
import numpy as np
import seaborn as sns
import datetime
import matplotlib as mpl
%matplotlib inline
import matplotlib as plt
import matplotlib.font_manager as fm

mpl.rcParams['axes.unicode_minus'] = False

path = 'C:/Windows/Fonts/malgun.ttf'
font_name = fm.FontProperties(fname=path, size=50).get_name()
plt.rc('font', family=font_name)


pd.set_option('display.max_columns', 10000)
pd.set_option('display.max_rows', 10000)

import warnings
warnings.filterwarnings(action='ignore')

# 법인카드 사용내역 정리
1. 매월 6일 데이터 다운로드 ( 외환, 국민 : 승인내역으로 대체)
    - 외환카드 승인기간 : 전월 6일 ~ 당월 5일
    - KB카드 승인기간 : 전월 9일 ~ 당월 8일

# 청구내역 사용시
os.chdir('./\\Data')
df1 = pd.read_csv('청구내역상세조회.CSV', encoding='cp949')

# 승인내역 카드번화 사용자 매칭

# 승인내역 사용하는 카드사 : 국민카드 , 외환카드
# kb = pd.read_csv('승인내역상세조회_국민.CSV', encoding='cp949')
ex = pd.read_csv('승인내역상세조회_외환.CSV', encoding='cp949')

user = pd.DataFrame({
    '카드번호' : [********],
    '사용자' : [********]
})
user['카드번호'] = pd.to_numeric(user['카드번호'])

# kb['사용자'] = kb['카드번호'].map(user.set_index('카드번호')['사용자'])
ex['사용자'] = ex['카드번호'].map(user.set_index('카드번호')['사용자'])

# 청구내역과 승인내역 리스트를 합치기
df = pd.concat([df1, ex])

# 승인내역에 없는 컬럼 이용일자, 청구금액 으로 정보 가져오기
df.loc[df['이용일자'].isnull(), '이용일자'] = df.loc[df['이용일자'].isnull(), '승인일자']
df.loc[df['청구금액'].isnull(), '청구금액'] = df.loc[df['청구금액'].isnull(), '승인금액']

df.rename(columns={'메모':'지점'}, inplace=True)
df['사용자'].fillna('기타', inplace=True)
df.drop(['카드별칭', '사용구분', 'ERP전송자','ERP전송일시'], axis=1, inplace=True)
df['카드번호'] = df['카드번호'].astype(str)
df['카드번호'] = df['카드번호'].str[0:4]+'-'+df['카드번호'].str[4:8]+'-'+df['카드번호'].str[8:12]+'-'+df['카드번호'].str[12:16]


# 카드별 사용자별 계정과목 구분
class card_cat:
    
    def __init__(self, df, account, store, indust):
        self.df = df
        self.account = account
        self.store = store
        self.indust = indust
        
    def get_store(self):
        self.df.loc[self.df['가맹점'].astype(str).str.contains(self.store), '계정과목'] = self.account
        
    def toptier(self):   # 파라미터 포함되어 있으므로 클래스 실행할 때 df만 넣고 나머지는 공란처리
        self.df.loc[(self.df['가맹점'].str.contains('*****')) & (self.df['청구금액'] == *****), '계정과목'] = '광고선전비_기타'
        self.df.loc[(self.df['가맹점'].str.contains('*****')) & (self.df['청구금액'] == *****), '지점'] = '광고선전비_배부'
        self.df.loc[self.df['가맹점'].str.contains('페이스북|당근마켓|FACEBK'), '계정과목'] = '광고선전비_배부'
        
    def supply_expense(self):   # 2순위 분류
        con1 = self.df['가맹점'].str.contains(self.store)
        con2 = self.df['업종'].notnull() & self.df['업종'].str.contains(self.indust)
        con3 = (self.df['업종'].notnull()) & (self.df['사용자'] == '****') & (self.df['업종'].str.contains('****'))
        con4 = self.df['계정과목'].isnull()
        
        self.df.loc[con1 | con2 | con3 & con4,'계정과목'] = self.account
        
    def welfare_benefit(self):
        con1 = self.df['가맹점'].str.contains(self.store)
        con2 = self.df['업종'].notnull() & self.df['업종'].str.contains(self.indust)
        con3 = (self.df['계정과목'].isnull()) & (self.df['가맹점'].str.contains('****'))
        con4 = (self.df['사용자'].str.contains('****')) & (self.df['업종'].str.contains('****')) & (self.df['업종'].notnull())
        con5 = self.df['계정과목'].isnull()
        
        self.df.loc[con1 | con2 | con3 | con4 & con5, '계정과목'] = self.account
        
    def commission_fee(self):
        con1 = self.df['가맹점'].str.contains(self.store)
        con2 = self.df['업종'].notnull() & self.df['업종'].str.contains(self.indust)
        con3 = self.df['계정과목'].isnull()
        
        self.df.loc[con1 | con2 & con3, '계정과목'] = self.account
        
    def ent_exp(self):
        con1 = self.df['가맹점'].str.contains(self.store)
        con2 = self.df['업종'].notnull() & self.df['업종'].str.contains(self.indust)
        con3 = (self.df['사용자'].str.contains('****')) & (self.df['청구금액'] > ****)
        con4 = (self.df['사용자'].str.contains('****')) & (self.df['청구금액'] > ****)
        
        self.df.loc[con1 | con2 | con3 | con4, '계정과목'] = 'self.account'
        
    def account_exp(self):
        con1 = self.df['가맹점'].str.contains(self.store)
        
        self.df.loc[con1, '계정과목'] = self.account
        
    def law_exp(self):
        con1 = self.df['가맹점'].str.contains(self.store)
        con2 = self.df['업종'].notnull() & self.df['업종'].str.contains(self.indust)
        con3 = self.df['계정과목'].isnull() & self.df['사용자'].str.contains('****')
        
        self.df.loc[con1 | con2 | con3, '계정과목'] = self.account
        
    def veh_cost(self):
        con1 = self.df['가맹점'].str.contains(self.store)
        con2 = self.df['업종'].notnull() & self.df['업종'].str.contains(self.indust)
        
        self.df.loc[con1 | con2, '계정과목'] = self.account
        
    def trv_exp(self):
        con1 = self.df['가맹점'].str.contains(self.store)
        con2 = self.df['업종'].notnull() & self.df['업종'].str.contains(self.indust)
        con3 = (self.df['가맹점'].str.contains('****')) & (self.df['사용자'] == '****')
        
        self.df.loc[con1 | con2 | con3, '계정과목'] = self.account
        
    def gas_exp(self):
        con1 = self.df['가맹점'].str.contains(self.store)
        con2 = self.df['업종'].notnull() & self.df['업종'].str.contains(self.indust)
        
        self.df.loc[con1 | con2 , '계정과목'] = self.account
        
    def comm_exp(self):
        con1 = self.df['가맹점'].str.contains(self.store)
        
        self.df.loc[con1, '계정과목'] = self.account
        
    def insur_exp(self):
        con1 = self.df['가맹점'].str.contains('****')
        con2 = self.df['청구금액'] < ****
        
        self.df.loc[con1 & con2, '계정과목'] = self.account
        
    def promo_exp(self):
        con1 = self.df['가맹점'].str.contains(self.store)
        
        self.df.loc[con1, '계정과목'] = self.account
        
    def tax_exp(self):
        con2 = self.df['업종'].notnull() & self.df['업종'].str.contains(self.indust)
        
        self.df.loc[con2, '계정과목'] = self.account
      
    def get_last(self):
        ***
        self.df.loc[con1 | con2 & con3, '계정과목'] = self.account

# 클래스 메서드 실행하기
a = card_cat(card_summary, _, _, _)
a.toptier()

store_list = '***'
indust_list = '***'
a = card_cat(card_summary, '소모품비_기타', store_list, indust_list)
a.supply_expense()

store_list = '***'
indust_list = '***'
a = card_cat(card_summary, '복리후생비_기타', store_list, indust_list)
a.welfare_benefit()

store_list = '***'
indust_list = '***'
a = card_cat(card_summary, '지급수수료_기타', store_list, indust_list)
a.welfare_benefit()

store_list = '***'
indust_list = '***'
a = card_cat(card_summary, '접대비', store_list, indust_list)
a.welfare_benefit()

store_list = '***'
# indust_list = 
a = card_cat(card_summary, '지급수수료_회계', store_list, indust_list)
a.account_exp()

store_list = '***'
indust_list = '***'
a = card_cat(card_summary, '지급수수료_법률', store_list, indust_list)
a.law_exp()

store_list = '***'
indust_list = '***'
a = card_cat(card_summary, '차량유지비_기타', store_list, indust_list)
a.veh_cost()

store_list = '***'
indust_list = '***'
a = card_cat(card_summary, '여비교통비_기타', store_list, indust_list)
a.trv_exp()

store_list = '***'
indust_list = '***'
a = card_cat(card_summary, '차량유지비_유류대', store_list, indust_list)
a.gas_exp()

store_list = '***'
indust_list = '***'
a = card_cat(card_summary, '통신비_우편대', store_list, indust_list)
a.comm_exp()

store_list = '***'
indust_list = '***'
a = card_cat(card_summary, '보험료_기타', store_list, indust_list)
a.insur_exp()

store_list = '***'
indust_list = '***'
a = card_cat(card_summary, '판촉비_기타', store_list, indust_list)
a.promo_exp()

store_list = '***'
indust_list = '***'
a = card_cat(card_summary, '세금과공과_기타', store_list, indust_list)
a.tax_exp()

store_list = '***'
indust_list = '***'
a = card_cat(card_summary, _, _, _)
a.get_last()


# 카드별, 사용자별 지점 구분
class Comcard:
    def __init__(self, df, col, num, br):
        self.df = df
        self.col = col     # 특정컬럼에서
        self.num = num     # 특정값을 보유할 때
        self.br = br       # 지점을 구분
    
    def card_num(self):
        self.df.loc[self.df[self.col].astype(str).str.contains(self.num), '지점'] = self.br

cdl = {   # 딕셔너리 정보입력
**** }

cdl2 = pd.DataFrame(list(cdl.items()), columns = ['카드번호', '지점'])

# 카드번호로 사용지점 구분
for i, j in cdl2.iterrows():
    a1 = Comcard(card_summary, '카드번호', j[0], j[1])
    a1.card_num()
    
# 사용자명으로 사용지점 구분
user_list = [***]

for i in user_list:
    a2 = Comcard(card_summary, '사용자', i, ***)
    a2.card_num()
    
가맹점별 = pd.DataFrame(card_summary.groupby(['계정과목','가맹점'])['청구금액'].sum())
지점별 = pd.DataFrame(card_summary.groupby(['지점','계정과목'])['청구금액'].sum())
사용자별 = pd.DataFrame(card_summary.groupby(['사용자','계정과목','가맹점'])['청구금액'].sum())
카드별 = pd.DataFrame(card_summary.groupby(['카드사','카드번호','사용자'])['청구금액'].sum())


# 요약내역 엑셀파일로 
os.chdir('./\\Result')
writer = pd.ExcelWriter('법인카드_정산내역_V1.xlsx', engine='xlsxwriter')

card_summary.to_excel(writer, sheet_name='전체')
지점별.to_excel(writer, sheet_name='지점별')
가맹점별.to_excel(writer, sheet_name='가명점별')
사용자별.to_excel(writer, sheet_name='사용자별')
카드별.to_excel(writer, sheet_name='카드별')

writer.save()
print(os.getcwd())
