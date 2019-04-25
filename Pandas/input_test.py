import os

def test01():
    cur_dir = os.getcwd()
    print('current directory : ' + cur_dir)
    print('exit : x') 
    print('source file : Douzone_2019____')
    while True:
        tail = input('source file tail(mmdd) : ')
        if tail == 'x':
            exit()
        else:
            excel_file = cur_dir + '\\' + "Douzone_2019" + tail + ".xls"
            if os.path.exists(excel_file):
                print(excel_file)
                break
            else:
                print('file not found : ' + excel_file)

def get_file_path():
    cur_dir = os.getcwd()
    print('current directory : ' + cur_dir)
    print('exit : x') 
    while True:
        input_data = input('source file path : ')
        if input_data == 'x':
            break
        else:
            if os.path.exists(input_data):
                #print(input_data)
                return input_data
            else:
                print('file not found : ' + input_data)
    return ''

def dept_info(file_path):
    import pandas as pd
    df = pd.read_excel(file_path, Sheet_name="Sheet1")

    df1 = df[['회사명', '부서명','사원명(ID)','직급']]#.head(30)

    df2 = df1[df1['부서명'].str.match(r'.*IT코디센터>(센터장|팀원|팀장)$')] #IT코디센터 팀명 발췌


file_path = get_file_path()

if len(file_path) > 0:
    print("OK : " + file_path)
else:
    print('exit')
