import pandas as pd
#import xlrd
import os
import re
import datetime

def commaParse(num):
   return re.sub('(?<=\d)(?=(\d{3})+(?!\d))',',',str(num))

DATA_PATH = "C:\\Temp\\NsmData\\"
today = datetime.datetime.today() # datetime.datetime.now
yesterday = today + datetime.timedelta(days=-1) # 오늘에서 1일을 빼서 어제를 구한다
file_name = "order" + yesterday.strftime('%Y%m%d') + ".xlsx"
#file_name ='order20190401to0419.xlsx'

excel_file = DATA_PATH + file_name
#excel_file = "C:\\Temp\\NsmData\order20190419.xlsx"

excel_file = "C:\\Dev\\Data\Sheet1.xlsx"

try:
    if not os.path.exists(excel_file):
        print('실패! ' + excel_file + ' file not found!! ')
        exists_files = os.listdir(DATA_PATH)
        print('Exists files :' + exists_files)
        exit()
except FileNotFoundError as e:
    print(e)
    exit()

# 프로그램 시작 시간 표시
start_dt = datetime.datetime.now()
ts = start_dt.timetuple()
print('지금 시간은  {}년 {}월 {}일 {}시 {}분 {}초'.format(ts.tm_year,ts.tm_mon, ts.tm_mday, ts.tm_hour, ts.tm_min, ts.tm_sec))

print('시작시간:' + start_dt.strftime('%Y-%m-%d, %H:%M:%S')) #'2019-04-19, 13:31:11'
print('=-='*30) # 줄긋기

df = pd.read_excel(excel_file, Sheet_name='Sheet1')

read_columns = ['주담당채널', '영업채널명', '영업담당자명', '고객명', '사업자/주민', '주문일자', '주문상태',
              '주문현황', '납입구분', '상품명', '단위상품명', '공급가', '금액(VAT별도)', '판매유형', '기회시작일' ]

df1 = df[read_columns]

product_category = {'ICUBE 클라우드서버', 'GW 클라우드서버', 'ICUBE/GW 클라우드서버', 'ICUBE/IU 클라우드서버',
                   'IU/GW 클라우드서버', 'Private Cloud', 'iU Private', 'iCUBE Private', '클라우드백업서버'} #집합(중복값 불허)

df2 = df1[df1['상품명'].isin(product_category)]

groupby_columns_list = ['영업채널명', '영업담당자명', '사업자/주민', '고객명', '상품명', '판매유형', '기회시작일']
grouped = df2.groupby(groupby_columns_list)
df3 = grouped.sum().reset_index()

# 업체명/솔루션/년금액/센터/담당자/판매유형
regex = re.compile(r'.*[^\s?IT코디센터$]') # space와 IT코디센터가 아닌것을 선택
price_sum = 0
for row in df3.values:
   text = row[0]
   match_obj = regex.search(text)
   if match_obj:
      it_coodi = match_obj.group().replace('/경북','')
   else: it_coodi = text
   price = row[8] # 년금액
   price_sum += int(price) # 합계금액 

   print("{0}/{1}/년{2}원/{3}/{4}/{5}".format(row[3], row[4].split()[0], commaParse(row[8]), it_coodi, row[1], row[5]))

comma_sum = commaParse(price_sum)
print("총 금액 : {}원".format(comma_sum))
# 결과 갯수 표시
#count_row = df3.shape[0]  # gives number of row count   참고) len(df3)는 결과값이 같으나 느림
#count_col = df3.shape[1]  # gives number of col count
print('{0}개가 검색되었습니다.'.format(df3.shape[0]))

print('-'*50)
# 업체명 (사업자번호)
for row in df3.values:
   print("- {0} ({1})".format(row[3], row[2]))

# 프로그램 종료 시간 표시
print('=-='*30) # 줄긋기
end_dt = datetime.datetime.now()
print("소요시간: {0}".format(end_dt - start_dt))
print('종료시간: ' + end_dt.strftime('%Y-%m-%d, %H:%M:%S')) #'2019-04-19, 13:31:11'
