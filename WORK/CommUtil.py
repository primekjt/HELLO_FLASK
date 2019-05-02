import re

def commaParse(num):
    """입력 숫자에 3자리 마다 콤마를 넣어 리턴한다 """
   return re.sub('(?<=\d)(?=(\d{3})+(?!\d))',',',str(num))


if __name__ == '__main__':
    pass