import xlrd
import xlwt






class ExcelData():
    def __init__(self, data_path, sheetname):
        self.data_path = data_path
        self.sheetname = sheetname
        self.data = xlrd.open_workbook(self.data_path,formatting_info=False)
        self.table = self.data.sheet_by_name(self.sheetname)


        self.rowNum = self.table.nrows
        self.colNum = self.table.ncols
        self.merged = self.table.merged_cells
        print(self.merged)



    def get_call(self):
        list=[]
        for (rlow, rhigh, clow, chigh) in self.merged:
            dict={}
            data=self.table.cell_value(rlow, clow)
            dict['data']=data
            dict['orgin']=(rlow, rhigh, clow, chigh)
            list+=[dict,]
        return list


    def read(self):
        l=[]
        for i in self.get_call():
            if i.get('data').find('合计金额(大写)')>=0:
                row=i.get('orgin')[0]
                break
        for i in self.get_call():
            if i.get('data').find('客户名称')>=0:
                for j in range(1,self.rowNum):
                    c_cell =self.table.cell_value(i.get('orgin')[0],j)
                    if c_cell !='':
                        l+=[{'orgin':(i.get('orgin')[0],j),'data':c_cell}]
                        break
        for i in range(row, self.rowNum):
            for j in range(self.colNum):
                c_cell = self.table.cell_value(i, j)
                if c_cell !='':
                    dict={}
                    dict['data']=c_cell
                    dict['orgin']=(i,j)
                    l+=[dict,]
        return l

    def write_file(self):
        worksheet = xlwt.Workbook()
        table = worksheet.add_sheet(self.sheetname,cell_overwrite_ok=True)

        font = xlwt.Font()
        font.height=20*15
        font2 = xlwt.Font()
        font2.height=20*11

        alignment=xlwt.Alignment()
        alignment.horz = 0x02
        alignment.vert = 0x01
        style=xlwt.XFStyle()
        style.alignment=alignment
        style.font=font

        style2=xlwt.XFStyle()
        style2.font=font2

        #写入已经合并的单元格
        for i in self.get_call():
            a,b,c,d=i.get('orgin')
            table.write_merge(a,b-1,c,d-1,i.get('data'),style)

        for i in self.read():
            a,b = i.get('orgin')
            table.write(a,b,i.get('data'),style2)

        worksheet.save('new.xls')



if __name__ == "__main__":
    data_path = "999.xlsx"
    sheetname = "对账单"
    get_data = ExcelData(data_path, sheetname)
    datas = get_data.write_file()


