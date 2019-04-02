import xlrd, xlwt, os

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
    realname = f'data/{filename}'
    data = xlrd.open_workbook(realname)
    table1 = data.sheets()[0]
    table2 = data.sheets()[1]
    table3 =
    first_name = table1.cell_value(2,1)
    last_name = table1.cell_value(1,1)
    name = first_name + last_name
    date = table1.cell_value(0,4)
    height = table1.cell_value(5,1)
    weight = table1.cell_value(6,1)
    age = table1.cell_value(4,1)
    gender = table1.cell_value(3,1)
    bmi = table1.cell_value(5,4)
    hrmax = table1.cell_value(6,4)
    mets = table2.cell_value(13,7)
    vo2max = table2.cell_value(11,7)
    return name, date, height, weight, age, gender, bmi, hrmax, mets, vo2max


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
