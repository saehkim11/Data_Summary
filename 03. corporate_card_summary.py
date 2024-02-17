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


### 법인카드 사용자 정보 가져오기

user = pd.read_excel(glob.glob('./Data/01.법인카드_관리현황.xlsx')[0], sheet_name='사용자')
user['카드번호'] = user['카드번호'].str.replace('-','').str[:16]

# 판촉비 안분비율 가져오기
br_ratio = pd.read_excel(glob.glob('./Data/01.법인카드_관리현황.xlsx')[0], sheet_name='안분비율')

# 계정구분 기준 가져오기
acc_std = pd.read_excel(glob.glob('./Data/02.계정구분기준.xlsx')[0], sheet_name=None)

# 청구내역만 있을 때 승인내역 있을 때 구분해서 데이터프레임을 합치기

dfs = pd.DataFrame()

if len(glob.glob('./Data/*.CSV')) == 1:
    file = glob.glob('./Data/03.청구내역상세조회.CSV')
    df1 = pd.read_csv(file[0], encoding='cp949')
    dfs = dfs.append(df1)
    
else:
    file = glob.glob('./Data/03.청구내역상세조회.CSV')
    df1 = pd.read_csv(file[0], encoding='cp949')
    

    dfs = dfs.append(df1)
    
    # 승인내역 사용하는 카드사 : 국민카드 , 외환카드
    if len(glob.glob('./Data/*국민*')) != 0:
        kb = pd.read_csv(glob.glob('./Data/*국민*')[0], encoding='cp949')
        kb['카드번호'] = kb['카드번호'].astype(str)
        kb['사용자'] = kb['카드번호'].map(user.set_index('카드번호')['사용자'])
        dfs = dfs.append(kb)
    else : 
        pass
    
    if len(glob.glob('./Data/*외환*')) != 0:
        ex = pd.read_csv(glob.glob('./Data/*외환*')[0], encoding='cp949')
        ex['카드번호'] = ex['카드번호'].astype(str)
        ex['사용자'] = ex['카드번호'].map(user.set_index('카드번호')['사용자'])
        dfs = dfs.append(ex)
    else :
        pass

    
    # 승인내역에 없는 컬럼 이용일자, 청구금액 으로 정보 가져오기
    dfs.loc[dfs['이용일자'].isnull(), '이용일자'] = dfs.loc[dfs['이용일자'].isnull(), '승인일자']
    dfs.loc[dfs['청구금액'].isnull(), '청구금액'] = dfs.loc[dfs['청구금액'].isnull(), '승인금액']  

df = dfs.copy()

df.rename(columns={'메모':'지점'}, inplace=True)
df['사용자'].fillna('기타', inplace=True)
df.drop(['카드별칭', '사용구분', 'ERP전송자','ERP전송일시'], axis=1, inplace=True)
df['카드번호'] = df['카드번호'].astype(str)
# df['카드번호'] = df['카드번호'].str[0:4]+'-'+df['카드번호'].str[4:8]+'-'+df['카드번호'].str[8:12]+'-'+df['카드번호'].str[12:16]

card_summary = df.copy()


### 카드별 사용자별 계정과목 구분
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

subset = acc_std['소모품비_기타']
a = card_cat(card_summary, '소모품비_기타', subset['거래처'], subset['업종'])   # card_cat 함수가 str.contains 가 아니라 isin() 으로 바뀌어야 함
a.supply_expense()

subset = acc_std['복리후생비_기타']
a = card_cat(card_summary, '복리후생비_기타', subset['거래처'], subset['업종'])
a.welfare_benefit()

subset = acc_std['지급수수료_기타']
a = card_cat(card_summary, '지급수수료_기타', subset['거래처'], subset['업종'])
a.welfare_benefit()

subset = acc_std['접대비']
a = card_cat(card_summary, '접대비', subset['거래처'], subset['업종'])
a.welfare_benefit()

subset = acc_std['지급수수료_회계']
a = card_cat(card_summary, '지급수수료_회계', subset['거래처'], subset['업종'])
a.account_exp()

subset = acc_std['지급수수료_법률']
a = card_cat(card_summary, '지급수수료_법률', subset['거래처'], subset['업종'])
a.law_exp()

subset = acc_std['차량유지비_기타']
a = card_cat(card_summary, '차량유지비_기타', subset['거래처'], subset['업종'])
a.veh_cost()

subset = acc_std['여비교통비_기타']
a = card_cat(card_summary, '여비교통비_기타', subset['거래처'], subset['업종'])
a.trv_exp()

subset = acc_std['차량유지비_유류대']
a = card_cat(card_summary, '차량유지비_유류대', subset['거래처'], subset['업종'])
a.gas_exp()

subset = acc_std['통신비_우편대']
a = card_cat(card_summary, '통신비_우편대', subset['거래처'], subset['업종'])
a.comm_exp()

subset = acc_std['보험료_기타']
a = card_cat(card_summary, '보험료_기타', subset['거래처'], subset['업종'])
a.insur_exp()

subset = acc_std['판촉비_기타']
a = card_cat(card_summary, '판촉비_기타', subset['거래처'], subset['업종'])
a.promo_exp()

subset = acc_std['세금과공과_기타']
a = card_cat(card_summary, '세금과공과_기타', subset['거래처'], subset['업종'])
a.tax_exp()

a = card_cat(card_summary)
a.get_last()


# 계정과목이 구분되지 않은 리스트들만 엑셀파일로 추출
if len(card_summary[card_summary['계정과목'].isnull()]) > 0 :
    card_summary[card_summary['계정과목'].isnull()].to_excel('./Data/계정누락분.xlsx')
else :
    pass

# 법인카드 관리내역의 카드번호로 지점을 구분하는 과정

class Comcard:
    def __init__(self, df, col, num, br):
        self.df = df
        self.col = col     # 특정컬럼에서
        self.num = num     # 특정값을 보유할 때
        self.br = br       # 지점을 구분
    
    def card_num(self):
        self.df.loc[self.df[self.col].astype(str).str.contains(self.num), '지점'] = self.br

for i, j in user.iterrows():
    a1 = Comcard(card_summary, '카드번호', j[0], j[2])   # 카드서머리 df 고정데이터와 법카 관리현황 데이터
    a1.card_num()                                      # for 반복문으로 법카 관리현황 하나씩 매칭해서 일치하는 것만 지점 기입


cols = ['카드사','카드번호','이용일자','관리자','사용자','가맹점','업종','청구금액','이용수수료','지점','계정과목']
card_summary = card_summary[cols]

가맹점별 = pd.DataFrame(card_summary.groupby(['계정과목','가맹점'])['청구금액'].sum())
지점별 = pd.DataFrame(card_summary.groupby(['지점','계정과목'])['청구금액'].sum())
사용자별 = pd.DataFrame(card_summary.groupby(['사용자','계정과목','가맹점'])['청구금액'].sum())
카드별 = pd.DataFrame(card_summary.groupby(['카드사','카드번호','사용자'])['청구금액'].sum())
삼십만원 = card_summary[card_summary['청구금액'] > xxxx].sort_values('청구금액', ascending=False)
유류대_데이터 = card_summary[card_summary['계정과목'] == '차량유지비_유류대'].sort_values('청구금액', ascending=False)
유류대_피벗 = card_summary[card_summary['계정과목'] == '차량유지비_유류대'].groupby(['사용자','카드번호'])['청구금액'].sum().to_frame().sort_values('청구금액', ascending=False)


판촉비_리스트 = card_summary[(card_summary['계정과목'].str.contains('광고|판촉')) & (card_summary['지점'] == '본사')]

div_exp = pd.DataFrame()

for i in 판촉비_리스트['청구금액']:
     div_exp = div_exp.append(i * br_ratio['비율'])
        
div_exp.columns = br_ratio['지점']
div_exp.index = 판촉비_리스트['청구금액']

판촉비_리스트_2 = 판촉비_리스트.merge(div_exp.stack().to_frame().reset_index().drop_duplicates(),left_on='청구금액', right_on='청구금액', how='left')
판촉비_리스트_2['청구금액'] = 판촉비_리스트_2[0]
판촉비_리스트_2 = 판촉비_리스트_2.drop(['지점_x',0], axis=1)
판촉비_리스트_2.rename(columns={'지점_y':'지점'})

# 기존 판촉비 리스트 제외 후 안분된 판촉비로 업데이트
card_summary_2 = card_summary[~((card_summary['계정과목'].str.contains('광고|판촉')) & (card_summary['지점'] == '본사'))]
card_summary_2 = card_summary_2.append(판촉비_리스트_2)
card_summary_2.loc[card_summary_2['지점'].isnull(),'지점'] = card_summary_2.loc[card_summary_2['지점'].isnull(),'지점_y']
card_summary_2 = card_summary_2.drop('지점_y', axis=1)

가맹점별_2 = pd.DataFrame(card_summary_2.groupby(['계정과목','가맹점'])['청구금액'].sum())
지점별_2 = pd.DataFrame(card_summary_2.groupby(['지점','계정과목'])['청구금액'].sum())
사용자별_2 = pd.DataFrame(card_summary_2.groupby(['사용자','계정과목','가맹점'])['청구금액'].sum())
카드별_2 = pd.DataFrame(card_summary_2.groupby(['카드사','카드번호','사용자'])['청구금액'].sum())

# ----------

# 가맹점별['청구금액'].sum(), 지점별['청구금액'].sum(), 사용자별['청구금액'].sum(), 카드별['청구금액'].sum()

# os.chdir('../')   # 상위폴더로 빠져나오기
# os.chdir('./\\Result')  # 하위폴더로 들어가기

정산월 = card_summary['이용일자'].dropna().astype(str).str[0:6].value_counts().index[0]
writer = pd.ExcelWriter('{}_법인카드_계정분류.xlsx'.format(정산월), engine='xlsxwriter')

#df.to_excel('당월_법인카드_구분_수정중.xlsx')
card_summary.to_excel(writer, sheet_name='원본')
card_summary_2.to_excel(writer, sheet_name='전체')
지점별_2.to_excel(writer, sheet_name='지점별')
가맹점별_2.to_excel(writer, sheet_name='가맹점별')
사용자별_2.to_excel(writer, sheet_name='사용자별')
카드별_2.to_excel(writer, sheet_name='카드별')
삼십만원.to_excel(writer, sheet_name='30만원이상')
유류대_데이터.to_excel(writer, sheet_name='유류대')
유류대_피벗.to_excel(writer, sheet_name='유류대_피벗')                                                                                                                       

writer.save()
