import pandas as pd
import xlrd
import os

path_to_file = os.path.join(os.path.expanduser("~"), "\\Data\\Douzone_20190416.xls")
if not os.path.exists(path_to_file):
    print("File not found!!")
else : 
    print(path_to_file)

df = pd.read_excel(path_to_file, Sheet_name="Sheet1")

df1 = df[['회사명', '부서명','사원명(ID)','직급']]#.head(30)

#df1[df1['부서명'].str.endswith('CLOUD사업팀>팀원')] #
#df1[df1['직급'].isin(['부장'])] #직급이 부장인 ROW 선택
#idx = df1["사원명(ID)"].apply(lambda x: x.startswith('김'))
#df1.loc[idx] # 사원의 성이 '김' 인 ROW 선택
#df2 = df1[df1['부서명'].apply(lambda x: 'CLOUD사업본부' in x)]

'''
# 입력한 전체이름에 정의한 팀에 속하는지 확인하여 결과를 리턴한다.
DIC_TEAM = {'서울1IT코디센터':'서울1', '서울2IT코디센터':'서울2','서울3IT코디센터':'서울3','서울4IT코디센터':'서울4'}

def isItCoordinator(full_name):
    for k in DIC_TEAM.values():
        #print(DIC_TEAM[k])
        if k in full_name:
            return True
    return False

df2 = df1[df1['부서명'].apply(lambda x: isItCoordinator(x))]
'''

#df2 = df1[df1['부서명'].apply(lambda x: '더존비즈온>TS본부>서울1' in x)]
#df2 = df1[df1['부서명'].str.match(r'.*>CLOUD사업본부>CLOUD사업부>*.*>(팀원|팀장|부서장)$')] #IT코디센터 팀명 발췌

df2 = df1[df1['부서명'].str.match(r'.*>.*?IT코디센터>(센터장|팀원|팀장)$')] #IT코디센터 팀명 발췌

df2.loc[:, 'NEW_COL0'] = df2.loc[:, '부서명'].str.replace(r'.*?>TS본부>|.*?>IT코디센터>|(IT코디센터>(센터장|팀원|팀장)$)', '') # 가로와 가로안의 글자 제거 (이름)

'''for row in df2['사원명(ID)']:
    #print(df2[row])
    print(row)
'''    

'''
change_idx = df2['사원명(ID)'].apply(lambda x: x.startswith('김'))
df3 = df2.loc[change_idx]
if 'NEW_COL' in df3.columns:
    del df3['NEW_COL']
else:
    print('not found!')
df3.loc[:, 'NEW_COL'] = df3.loc[:, '사원명(ID)'].str.replace(r'\(.*?\)', '***')
'''

#df2.loc[:, 'NEW_COL1'] = df2.loc[:, '사원명(ID)'].str.replace(r'\(.*?\)$', '') # 가로와 가로안의 글자 제거 (이름)
df2.loc[:, 'NEW_COL2'] = df2.loc[:, '사원명(ID)'].str.replace(r'[^가-힣]', '') # 한글 이름 이외 글자 제거 (한글이름, 영문제외)
#df2.loc[:, 'NEW_COL3'] = df2['사원명(ID)'].str.replace(r'([가-힣]|([a-zA-Z]\()|\(|\))', '') # 가로안의 ID이외 글자 제거(ID)
df2.loc[:, 'NEW_COL3'] = df2.loc[:, '사원명(ID)'].str.replace(r'([가-힣]|([a-zA-Z]\()|\(|\))', '') # 가로안의 ID이외 글자 제거(ID)
#df2.loc[:,"사원명(ID)"].replace(r'[^가-힣]','', inplace=True, regex=True) #칼럼 추가 없이 바꾸기

LEVEL_TO_CODE = {'사원':'', '대리':'D', '과장':'K', '차장':'C', '부장':'B', '이사':'E', '상무':'S', '전무':'J'}
df2.loc[:,'LEVEL_CODE'] = df2['직급'].apply(lambda x : LEVEL_TO_CODE[x])

row_count = df2.count(0)[0] # data row count
print(row_count)

df3 = df2.copy()

df3.to_csv('./data/sample.csv', index = False, header = True)

df4 = pd.read_csv('./data/sample.csv')

df4.nunique(axis=0)

