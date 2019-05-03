import pandas as pd
#import xlrd
import os
import re
import datetime


def input_file_path():
    #cur_dir = os.getcwd()
    #print('current directory : ' + cur_dir)
    print('program exit : \'x\' press key or \'Enter\' key')
    while True:
        input_data = input('input source file path : ')
        if input_data == 'x' or input_data == '':
            break
        else:
            if os.path.exists(input_data):
                #print(input_data)
                return input_data
            else:
                print('file not found : ' + input_data)
    return ''

def get_filter_dataframe(df):
    read_columns = ['주담당채널', '영업채널명', '영업담당자명', '고객명', '사업자/주민', '주문일자', '주문상태',
                    '주문현황', '납입구분', '상품명', '단위상품명', '공급가', '금액(VAT별도)', '판매유형', '기회시작일']

    df1 = df[read_columns]

    product_category = {'ICUBE 클라우드서버', 'GW 클라우드서버', 'ICUBE/GW 클라우드서버', 'ICUBE/IU 클라우드서버',
                        'IU/GW 클라우드서버', 'Private Cloud', 'iU Private', 'iCUBE Private', '클라우드백업서버'}  # 집합(중복값 불허)

    df2 = df1[df1['상품명'].isin(product_category)]

    groupby_columns_list = ['영업채널명', '영업담당자명', '사업자/주민', '고객명', '상품명', '판매유형', '기회시작일']
    grouped = df2.groupby(groupby_columns_list)
    df3 = grouped.sum().reset_index()

    df4 = df3.sort_values(by=['기회시작일', '상품명', '영업채널명', '영업담당자명', '고객명'], ascending=[True, True, True, True, True])



    # 업체명/솔루션/년금액/센터/담당자/판매유형
    """
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
    """
    return df4

def save_excel_file(data_file_path, save_df):
    #df_dic = {'영업채널명': [], '영업담당자명':[], '사업자/주민':[], '고객명':[], '상품명':[], '판매유형':[], '공급가':[], '금액(VAT별도)':[]}
    start, ext = os.path.splitext(data_file_path)  # ('C:\\Users\\김진태\\Jupyter_Folder\\Data\\order20190401to0419.xlsx', '.xlsx')
    tail = '' + "_{:%Y%m%d%H%M%S}".format(datetime.datetime.now())  # '_20190503135732'
    save_file_name = start + tail + ext
    save_df.to_excel(save_file_name)


def save_text_file(data_file_path, save_text):
    start, ext = os.path.splitext(
        data_file_path)  # ('C:\\Users\\김진태\\Jupyter_Folder\\Data\\order20190401to0419', '.xlsx')
    tail = '' + "_{:%Y%m%d%H%M%S}".format(datetime.datetime.now())  # '_20190503135732'
    save_file_name = start + tail + '.txt'

    with open(save_file_name, 'w') as f:
        f.write(save_text)


def main():
    print('{0}[{1:^21}]{2}'.format('=' * 20, '더존NSM 발주정보 확인(클라우드사업부 전용)', '=' * 20))
    print('{:>72}'.format('2019.05.03 K.J.T'))
    print('-'*76)

    data_file_path = input_file_path()

    if 0 >= len(data_file_path):
        print('program exit... good-bye!')
        exit()


    # 프로그램 시작 시간 표시
    start_dt = datetime.datetime.now()
    ts = start_dt.timetuple()
    print('지금 시간은  {}년 {}월 {}일 {}시 {}분 {}초'.format(ts.tm_year,ts.tm_mon, ts.tm_mday, ts.tm_hour, ts.tm_min, ts.tm_sec))

    print('시작시간:' + start_dt.strftime('%Y-%m-%d, %H:%M:%S')) #'2019-04-19, 13:31:11'
    print('=-='*25) # 줄긋기

    df = pd.read_excel(data_file_path, Sheet_name='Sheet1')

    df2 = get_filter_dataframe(df)

    save_excel_file(data_file_path, df2)

    # 업체명/솔루션/년금액/센터/담당자/판매유형
    save_text = ''
    price_sum = int()
    for row in df2.values:
        org_price = int(row[8]) # 년금액
        price_sum += org_price  # 합계금액
        comma_price = "{:,}원".format(org_price)
        dept_name = row[0].replace('IT코디센터', '').rstrip().replace('/경북', '')
        print_text = "- {0}/{1}/년{2}/{3}/{4}/{5}".format(row[3], row[4].split()[0], comma_price, dept_name, row[1], row[5])
        print(print_text)
        save_text += print_text + '\n'


    # 결과 갯수 표시
    #count_row = df2.shape[0]  # gives number of row count   참고) len(df2)는 결과값이 같으나 느림
    #count_col = df2.shape[1]  # gives number of col count
    #print('{0}개가 검색되었습니다.'.format(df2.shape[0]))

    comma_sum = "{:,}원".format(price_sum)
    print_text = "총금액: {}({}건)".format(comma_sum, df2.shape[0])
    print(print_text)
    save_text += print_text + '\n'

    print('-'*50)
    # 업체명 (사업자번호)
    for row in df2.values:
        print_text = "- {0} ({1})".format(row[3], row[2])
        print(print_text)
        save_text += print_text + '\n'

    save_text_file(data_file_path, save_text)

    # 프로그램 종료 시간 표시
    print('=-='*30) # 줄긋기
    end_dt = datetime.datetime.now()
    print("소요시간: {0}".format(end_dt - start_dt))
    print('종료시간: ' + end_dt.strftime('%Y-%m-%d, %H:%M:%S')) #'2019-04-19, 13:31:11'

    print('-' * 76)
    print('create success! good-bye!')


if __name__ == "__main__":
    main()