import json
import pandas as pd


def first(xl):
    data = [{}]
    k = 1
    
    for x in xl.sheet_names:
        print(x[0])
        if x == 'Лист1':
            df1 = xl.parse(x)
            data = [{} for _ in range(df1.count()['type'])]
            for y in df1.keys():
                for i, u in enumerate(df1.get(y)):
                    if (y=='level'):
                        k=u
                    if not pd.isna(u):
                        data[i][y] = u
                    else:
                        data[i][y] = ""
            with open(name + ".json", 'w', encoding='utf-8') as w:
                json.dump(data, w, ensure_ascii=False, indent=4)
            k += 1

name = 'cup_128'
xl = pd.ExcelFile(name + '.xlsx')
first(xl)

name = 'cup_64'
xl = pd.ExcelFile(name + '.xlsx')
first(xl)

name = 'cup_32'
xl = pd.ExcelFile(name + '.xlsx')
first(xl)

name = 'cup_16'
xl = pd.ExcelFile(name + '.xlsx')
first(xl)

name = 'cup_8'
xl = pd.ExcelFile(name + '.xlsx')
first(xl)

name = 'cup_4'
xl = pd.ExcelFile(name + '.xlsx')
first(xl)

name = 'cup_2'
xl = pd.ExcelFile(name + '.xlsx')
first(xl)

name = 'cup_f'
xl = pd.ExcelFile(name + '.xlsx')
first(xl)

name = 'cup_s'
xl = pd.ExcelFile(name + '.xlsx')
first(xl)

