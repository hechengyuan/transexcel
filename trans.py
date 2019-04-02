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
    realname = 'data/{}'.format(filename)
    data = xlrd.open_workbook(realname)
    table1 = data.sheets()[0]
    table2 = data.sheets()[1]
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

def write_worksheet(row,line,slice):
    worksheet.write(row,line,slice[line])


def write_data():
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('worksheet')
    n = 0
    for file in get_all_name():
        slice = read_data(file)
        worksheet.write(n,0,slice[0])
        worksheet.write(n,1,slice[1])
        worksheet.write(n,2,slice[2])
        worksheet.write(n,3,slice[3])
        worksheet.write(n,4,slice[4])
        worksheet.write(n,5,slice[5])
        worksheet.write(n,6,slice[6])
        worksheet.write(n,7,slice[7])
        worksheet.write(n,8,slice[8])
        worksheet.write(n,9,slice[9])
        n = n + 1



    workbook.save('all_data.xls')

write_data()
