import json
import os
import sys
import xlsxwriter
from string import uppercase
from collections import OrderedDict

def save_to_excel(j, name):
    path = os.path.join(os.path.abspath('.'), name)
    workbook = xlsxwriter.Workbook(path)
    title = workbook.add_format({"align": "center", "font_size": 10,
                                 "font_name": "Times New Roman"})
    worksheet = workbook.add_worksheet('CWE')

    # add titles of each column at the top of excel worksheet
    titles= ['binary', 'CWE', 'description']
    for t, i in zip(titles, range(len(titles))):
        worksheet.write(uppercase[i]+'1', t, title)

    r=1
    c=0
    try:
        for binary in j:
            if binary is not None:
                for item in j[binary]:
                    worksheet.write(r, c, binary, title)
                    c += 1
                    for k , v in item.iteritems():
                        worksheet.write(r, c, v, title)
                        c += 1
                    r += 1
                    c=0
        workbook.close()
    except IndexError:
        pass


if __name__ == '__main__':
    # open json file
    path = os.path.join(os.path.abspath('.'),'data.txt')
    with open(path, 'r') as f:
        j = json.load(f, encoding='utf-8')
        f.close()
    j_ = OrderedDict(sorted(j.items(), key=lambda j: j[0]))
    save_to_excel(j_, 'result.xlsx')
