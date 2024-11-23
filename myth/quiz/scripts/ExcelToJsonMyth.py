import json
import pandas as pd


def ExcelToJson(xl):

    for x in xl.sheet_names:
        try:
            df1 = xl.parse(x)
            data = [{} for _ in range(df1.count()['textQue'])]
            for y in df1.keys():
                for i, u in enumerate(df1.get(y)):
                    if not pd.isna(str(u)):
                        if 'wrAns' not in y:
                            data[i][y] = str(u).strip()
                        else:
                            if data[i].get('wrAns'):
                                data[i]['wrAns'].append(str(u).strip())
                            else:
                                data[i]['wrAns'] = [str(u).strip()]
                    elif 'comment' in y:
                        data[i][y] = ''
            with open(x + ".json", 'w', encoding='utf-8') as w:
                json.dump(data, w, ensure_ascii=False, indent=4)
        except Exception as e:
            print(e)


name = 'ancient_greece'
xl = pd.ExcelFile(name + '.xlsx')
ExcelToJson(xl)
