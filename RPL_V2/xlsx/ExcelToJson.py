import json
import pandas as pd


def ExcelToJson(xl):

    for x in xl.sheet_names:
        df1 = xl.parse(x)
        data = [{} for _ in range(df1.count()['type'])]
        for y in df1.keys():
            for i, u in enumerate(df1.get(y)):
                if not pd.isna(u):
                    data[i][y] = u
                else:
                    data[i][y] = ""
        with open(x + ".json", 'w', encoding='utf-8') as w:
            json.dump(data, w, ensure_ascii=False, indent=4)


name = 'Матч Тура 25-26'
xl = pd.ExcelFile(name + '.xlsx')
ExcelToJson(xl)