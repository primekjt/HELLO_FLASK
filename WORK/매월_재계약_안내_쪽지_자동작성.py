import pandas
import openpyxl
#import re #regex
from datetime import datetime

def get_header_001(center_name, month):
    header = """
[재계약 리스트] {0}IT코디센터 2019년 {1} 재계약 리스트 전달 및 수주 등록 요청드립니다.
수신 : {0}IT코디센터 센터장님, 영업담당자
참조 : CLOUD사업팀
발신 : 김진태 부장
-----------------------------------------------
안녕하십니까? 클라우드사업팀 김진태입니다

클라우드서버 및 클라우드ERP 상품 관련하여 아래와 같이 재계약 리스트를 전달하오니
관련 상품 재계약 여부 확인하시어 수주 등록 부탁드립니다.(재계약 NSM 발주)

※ 재계약 등록 시 솔루션에 맞는 상품으로 진행 부탁드립니다.
※ 계산서 발행월 기준 리스트로 누락 없이 진행 부탁드립니다.
※ 당월 반품.해지등의 사유로 재계약 불가시 당월 이용요금이 청구될 수 있습니다.

※ 해지 등의 사유로 미진행시 반드시 회신 부탁드립니다.
   (쪽지 수신자 : 남원길 부장, 김진태 부장, 홍민순 과장, 김태은 대리 포함 쪽지 회신 요망)

감사합니다.

-------------------- 아   래 ------------------
[ 2019년 {1} 재계약 리스트 ]

센터 | 상품 | 납부 | 고객명 | 사업자번호 | 월금액 | 재계약금액 | 재계약월
    """.format(center_name, month)
    return header


def text_print_001(ws_rows):
    # pattern = re.compile('(?<=\d)(?=(\d{3})+(?!\d))')
    pre_center_name = 'none'
    result_text = ''
    for row in ws_rows:
        center_name = row[1].value

        # 첫줄은 pass
        if center_name == '센터':
            pre_center_name = center_name
            continue

            # month_price = re.sub(pattern,',',str(row[6].value))
        if 's' == row[6].data_type:  # 칼럼 헤더 구분
            month_price = row[6].value
            # print('{0}:{1}: row index = {2}'.format(row[6].value, type(row[6].value), row[6].row))
        else:
            month_price = "{:,}원".format(int(row[6].value))

        # year_price = re.sub(pattern,',',str(row[7].value))
        if 's' == row[7].data_type:  # 칼럼 헤더 구분
            year_price = str(row[7].value)
        else:
            year_price = "{:,}원".format(int(row[7].value))

        month = row[8].value  # 재계약월

        # print('pre_center_name:{0} , center_name:{1}'.format(pre_center_name, center_name)) #debug

        if pre_center_name != center_name:
            if pre_center_name != '센터':
                print_text = '\n{0}   끝   {1}\n'.format('-' * 20, '-' * 20)
                result_text += print_text + '\n'
                # print(print_text)
            if pre_center_name != 'none':
                print_text = get_header_001(center_name, month)
                # print_text = '[ header ]' + '\n'  #debug
                result_text += print_text + '\n'
                # print(print_text)

            pre_center_name = center_name

        print_text = '{0} | {1} | {2} | {3} | {4} | {5} | {6} | {7}'.format(
            row[1].value, row[2].value, row[3].value, row[4].value, row[5].value, month_price, year_price, row[8].value)
        result_text += print_text + '\n'
        # print(print_text)

    # 마지막 라인에 '끝' 추가
    print_text = '\n{0}   끝   {1}\n'.format('-' * 20, '-' * 20)
    result_text += print_text + '\n'
    # print(print_text)

    return result_text

if __name__ == "__main__":
    data_file_path = 'C:\\Users\\김진태\\Jupyter_Folder\\Data\\5월 재계약 업체 리스트_20190502.xlsx'
    wb = openpyxl.load_workbook(data_file_path)

    try:
        ws = wb["Sheet1"]

        # 첫줄을 읽어 칼럼명 리스트를 만든다
        column_name_list = [cell.value for cell in ws[1]]

        save_text = text_print_001(ws.rows)
        # print(save_text)
        save_file_name = "C:\\Users\\김진태\\Jupyter_Folder\\Data\\재계약_쪽지_" + datetime.now().strftime('%Y%m%d') + ".txt"
        with open(save_file_name, 'w') as f:
            f.write(save_text)

    except Exception as e:
        print('Exception : %s' % e)
    finally:
        wb.close()
