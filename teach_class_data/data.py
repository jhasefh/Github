import xlrd
import xlwt
import random

excel_file = xlrd.open_workbook(r'/Users/mengqingdong/Github/teach_class_data/教学班.xlsx')

sheet = excel_file.sheet_by_index(0)

# import pdb; pdb.set_trace()
fri_col = sheet.col(1)
sec_col = sheet.col(2)
thr_col = sheet.col(3)
fou_col = sheet.col(4)
fiv_col = sheet.col(5)

del fri_col[0]
del sec_col[0]
del thr_col[0]
del fou_col[0]
del fiv_col[0]

# 生层随机数
def num_random():
    
    random_list = []
    for index in range(int(3206/7)):
        
        # a = random.randint(0,6)
        s=[]
        while len(s)<3:
            a=(random.randint(0,6))+index * 7
            if a not in s:
                s.append(a)

        random_list = random_list + s
    return random_list


# 删除列表的第数组下标的元素
def del_list(num_list,nun_random):
    return([num_list[i] for i in range(len(num_list)) if (i not in nun_random)])

# 将列数据单元格格式转换为str
def turn_cel(col):
    s = []
    import pdb; pdb.set_trace()
    for x in range(len(col)):
        s.append(col[x].value)
    return s
# 写入数据
def write_exl(col1,col2,col3,col4,col5):

    # 创建一个workbook 设置编码
    workbook = xlwt.Workbook(encoding = 'utf-8')
    # 创建一个worksheet
    worksheet = workbook.add_sheet('My Worksheet')
    for row in range(int(len(col1))):
        worksheet.write(row,0,col1[row])
    for row in range(int(len(col1))):
        worksheet.write(row,1,col2[row])
    for row in range(int(len(col1))):
        worksheet.write(row,2,col3[row])
    for row in range(int(len(col1))):
        worksheet.write(row,3,col4[row])
    for row in range(int(len(col1))):
        worksheet.write(row,4,col5[row])
    workbook.save('Excel_test.xls')
# 主函数
def main():
    random_list_tol = num_random()

    fri_col_final = del_list(fri_col,random_list_tol)
    sec_col_final = del_list(sec_col,random_list_tol)
    thr_col_final = del_list(thr_col,random_list_tol)
    fou_col_final = del_list(fou_col,random_list_tol)
    fiv_col_final = del_list(fiv_col,random_list_tol)

    a = turn_cel(fri_col_final)
    b = turn_cel(sec_col_final)
    c = turn_cel(thr_col_final)
    d = turn_cel(fou_col_final)
    e = turn_cel(fiv_col_final)
    # import pdb; pdb.set_trace()

    write_exl(a,b,c,d,e)

if __name__ == "__main__":
    main()