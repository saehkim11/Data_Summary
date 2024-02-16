import pandas as pd
import os
import re

import warnings
warnings.filterwarnings(action='ignore')

pd.set_option('display.max_columns', 500)
pd.set_option('display.max_rows', 500)

date = 'yyyymmdd' # 회계일자 입력

def sales_closing(target):
    
    df = pd.read_excel('온라인매출 정산.xlsx', sheet_name='매출정산', skiprows=3)
    if target == '상품매출':
        desc = '상품매출 적요' ; value = '공급금액' ; tax = '세액' ; total = '순매출' ; sales_code = '410103'
    else:
        desc = '기타매출 적요' ; value = '공급금액.1' ; tax = '세액.1' ; total = '합계금액' ; sales_code = '410300'
    
    df['거래처.1'] = df['거래처.1'].astype(str).apply(lambda x: x.replace('.0',''))
    df['사업자등록번호'] = df['사업자등록번호'].astype(str).apply(lambda x: x.replace('.0',''))
    df['사업장'] = '0' + df['사업장'].astype(str).apply(lambda x: x.replace('.0',''))
    df['코스트센터'] = df['코스트센터'].astype(str).apply(lambda x: x.replace('.0',''))
    
    # 적요 desc를 인덱스로 피벗 합계
    df_2 = pd.pivot_table(df, index= desc, values= [value, tax, total], aggfunc='sum')
    df_3 = df_2[df_2[value] != 0]   # 실적 0원인 row 제거

    # 컬럼을 stack 함수 사용해 -> 인덱스로 재배열
    df_4 = pd.DataFrame(df_3.stack()).reset_index()
    cols = ['적요', '계정', '금액']
    df_4.columns = cols
    
    # ACHECK 필요정보를 적요 desc에서 추출
    df_4['귀속사업장'] = df_4['적요'].map(df.drop_duplicates(desc).set_index(desc)['사업장'])
    df_4['코스트센터'] = df_4['적요'].map(df.drop_duplicates(desc).set_index(desc)['코스트센터'])
    df_4['거래처코드'] = df_4['적요'].map(df.drop_duplicates(desc).set_index(desc)['거래처.1'])

    df_4.loc[df_4['계정'] == total, '차대구분'] = 1
    df_4.loc[df_4['계정'] == value, '차대구분'] = 2
    df_4.loc[df_4['계정'] == tax, '차대구분'] = 2

    df_4.loc[df_4['계정'] == value, '계정코드'] = sales_code
    df_4.loc[df_4['계정'] == tax, '계정코드'] = '210600'
    df_4.loc[df_4['계정'] == total, '계정코드'] = '110301'

    df_4['회계일자'] = date
    df_4.loc[df_4['계정코드'] == '210600', '발행일'] = date
    df_4.loc[df_4['계정코드'] == '210600', '세무구분'] = '18'
    df_4.loc[df_4['계정'] == total, '코스트센터'] = '112000'

    df_4['과세표준액'] = df_4[df_4['계정코드'] == '210600']['적요'].map(df.drop_duplicates(desc).set_index(desc)[value])
    df_4[tax] = df_4[df_4['계정코드'] == '210600']['적요'].map(df.drop_duplicates(desc).set_index(desc)[tax])

    df_4.loc[df_4['계정'] == value, '귀속사업장'] = ''
    df_4['자금예정일'] = ''
    df_4['사업자등록번호'] = ''
    df_4['부서'] = ''

    cols = ['차대구분', '계정코드', '금액', '거래처코드', '발행일','자금예정일', '과세표준액', tax,'세무구분','사업자등록번호','적요', '귀속사업장','부서','코스트센터']
    
    if target == '상품매출':
        df_4[cols].to_excel('상품매출_회계전표_양식.xlsx')
    else:
        df_4[cols].to_excel('기타매출_회계전표_양식.xlsx')

# 함수 실행 -> 자동전표 양식으로 엑셀파일 추출
sales_closing('상품매출')
sales_closing('기타매출')
