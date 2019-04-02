import xlrd, xlwt, os, json


def get_all_name():
    path = 'data/'
    files = os.listdir(path)
    list = []
    for file in files:
        if file[0] == '~':
            pass
        else:
            list.append(file)
    return list


def read_data(filename):
    data = xlrd.open_workbook(f'data/{filename}')
    table1 = data.sheets()[0]
    table2 = data.sheets()[1]
    with open('setting.json','r') as f:
        setting = json.load(f)

    slice = []
    for key in list(setting.keys()):
        if setting[key][0] == 0:
            slice.append(table1.cell_value(setting[key][1],setting[key][2]))
        elif setting[key][0] == 1:
            slice.append(table2.cell_value(setting[key][1],setting[key][2]))
        else:
            pass

    return slice


def write_data():
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('worksheet')
    n = 0

    for file in get_all_name():
        slice = read_data(file)

        for i in range(0,len(slice)):
            worksheet.write(n,i,slice[i])

        n = n + 1

    workbook.save('all_data.xls')


write_data()
