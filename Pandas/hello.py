import re
#import pandas as pd

#df = pd.DataFrame({'a':'one', 'b': 1, 'c':10}, {'a':'two', 'b':2, 'c':20}, {'a':'three', 'b':3, 'c':30})
print('hello world')

text = '서울1 IT코디센터'
text_list = ['서울1 IT코디센터', '부산IT코디센터', '대구/경북IT코디센터']

regex = re.compile(r'.*[^\s?IT코디센터$]')
for text in text_list:
    match_obj = regex.search(text)
    print(match_obj)
    if match_obj:
        a = match_obj.group().replace('/경북','')
        print("["+a+"]")