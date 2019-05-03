import re
import os
from datetime import datetime

def commaParse(num):
    """입력 숫자에 3자리 마다 콤마를 넣어 리턴한다 """
   return re.sub('(?<=\d)(?=(\d{3})+(?!\d))',',',str(num))

def input_file_path():
    cur_dir = os.getcwd()
    print('current directory : ' + cur_dir)
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

def next_file_name():
    today = datetime.datetime.today()  # datetime.datetime.now()
    yesterday = today + datetime.timedelta(days=-1)  # 오늘에서 1일을 빼서 어제를 구한다
    file_name = "order" + yesterday.strftime('%Y%m%d') + ".xlsx"
    # file_name ='order20190401to0419.xlsx'
    return file_name

if __name__ == '__main__':
    pass