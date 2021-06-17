import xlrd,os


file_name='999.xlsx'
excel_file=os.getcwd()+os.sep+file_name
rdata=xlrd.open_workbook(excel_file)

#获取所有的工作表
names=rdata.sheet_names()
# print(names)

# 获取到table对象
table = rdata.sheet_by_index(0)
# print(table)

#获取行数
nrow = table.nrows
print(nrow)

print(table.row(9))