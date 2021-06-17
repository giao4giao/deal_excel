import xlrd
import xlwt
from xlrd import xldate_as_tuple
import datetime
import math

'''
xlrd中单元格的数据类型
数字一律按浮点型输出，日期输出成一串小数，布尔型输出0或1，所以我们必须在程序中做判断处理转换
成我们想要的数据类型
0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
'''


class ExcelData():
	# 初始化方法
	def __init__(self, data_path, sheetname):
		# 定义一个属性接收文件路径
		self.data_path = data_path
		# 定义一个属性接收工作表名称
		self.sheetname = sheetname
		# 使用xlrd模块打开excel表读取数据
		try:
			self.data = xlrd.open_workbook(self.data_path)
		except FileNotFoundError:
			print('未发现源文件')
			exit()
		# 根据工作表的名称获取工作表中的内容（方式①）
		self.table = self.data.sheet_by_name(self.sheetname)
		# 根据工作表的索引获取工作表的内容（方式②）
		# self.table = self.data.sheet_by_name(0)
		# 获取第一行所有内容,如果括号中1就是第二行，这点跟列表索引类似
		self.keys = self.table.row_values(6)
		# 获取工作表的有效行数
		self.rowNum = self.table.nrows
		# 获取工作表的有效列数
		self.colNum = self.table.ncols

		self.merged = self.table.merged_cells

		self.width = 256 * 20
		self.worksheet = xlwt.Workbook()
		self.write_table = self.worksheet.add_sheet(self.sheetname, cell_overwrite_ok=True)

	# 定义一个读取excel表的方法
	def readExcel(self):
		# 定义一个空列表
		datas = []
		for i in range(7, self.rowNum - 7):
			# 定义一个空字典
			sheet_data = {}
			sheet_data['行数'] = i
			for j in range(self.colNum):
				# 获取单元格数据类型
				c_type = self.table.cell(i, j).ctype
				# 获取单元格数据
				c_cell = self.table.cell_value(i, j)
				if c_type == 2 and c_cell % 1 == 0:  # 如果是整形
					c_cell = int(c_cell)
				elif c_type == 3:
					# 转成datetime对象
					date = datetime.datetime(*xldate_as_tuple(c_cell, 0))
					# c_cell = date.strftime('%Y/%d/%m %H:%M:%S')
					c_cell = date.strftime('%Y/%#m/%#d')
				elif c_type == 4:
					c_cell = True if c_cell == 1 else False
				sheet_data[self.keys[j]] = c_cell
				# 循环每一个有效的单元格，将字段与值对应存储到字典中
				# 字典的key就是excel表中每列第一行的字段
				# sheet_data[self.keys[j]] = self.table.row_values(i)[j]
			# 再将字典追加到列表中
			datas.append(sheet_data)
		# 返回从excel中获取到的数据：以列表存字典的形式返回
		return datas

	def solve_data(self, names):
		self.write_file(names)
		data = self.readExcel()
		if not names:
			return None
		list = []
		for name in names:
			l = []
			# for i in data[0:10]:
			for i in data:
				# print(i.get('产品名称'))
				# print(i.get('产品名称')=='、' and now)
				if i.get('产品名称') in ['、', '挂版'] and now:
					l += [i, ]
				else:
					now = False
				if i.get('产品名称').find(name) != -1:
					now = True
					l += [i, ]
			# print(l)
			for i in l:
				n = [i.get('送货日期'), i.get('单号'), i.get('产品名称'), i.get('规格'), i.get('单位'), i.get('数量')]
				ns = [i.get('单价'), i.get('金额'), i.get('备注')]
				list += [(n + ns), ]
			black = ('', '', '', '', '', '', '', '', '')
			list += [(('合计', '', '', '', '', '', '', sum([float(i.get('金额')) for i in l]), '')), black]
		datas = []
		for i in data:
			n = [i.get('送货日期'), i.get('单号'), i.get('产品名称'), i.get('规格'), i.get('单位'), i.get('数量')]
			ns = [i.get('单价'), i.get('金额'), i.get('备注')]
			datas += [(n + ns), ]
		l = []
		for i in datas:
			if i not in list:
				l += [i, ]
				list += [i, ]
		black = ('', '', '', '', '', '', '', '', '')
		list += [(('合计', '', '', '', '', '', '', sum([float(i[-2]) for i in l]), '')), black]

		# print(list)
		font2 = xlwt.Font()
		font2.height = 20 * 11
		font2.name = '宋体'
		alignment = xlwt.Alignment()
		alignment.horz = 0x02
		alignment.vert = 0x01
		borders = xlwt.Borders()
		borders.left = xlwt.Borders.MEDIUM
		borders.right = xlwt.Borders.MEDIUM
		borders.top = xlwt.Borders.MEDIUM
		borders.bottom = xlwt.Borders.MEDIUM
		pattern = xlwt.Pattern()
		pattern.pattern = xlwt.Pattern.SOLID_PATTERN
		pattern.pattern_fore_colour = 5

		style = xlwt.XFStyle()
		style.font = font2
		style.alignment = alignment
		style.borders = borders

		style2 = xlwt.XFStyle()
		style2.font = font2
		style2.alignment = alignment
		style2.borders = borders
		style2.pattern = pattern
		n = 7
		for j in range(self.colNum):
			num = max([len(tuple((str(i[j]).replace('.','').replace('/','')))) * 450 if type(list[0][j]) == type('') else len(tuple((str(i[j]).replace('.','').replace('/',''))))* 200 for i in list])
			# num = max([len((str(i[j]).replace('.','')) * 500 for i in list])
			num = num if num > 1500 else 1500
			self.write_table.col(j).width = num
			self.write_table.col(j).height = self.width
		for i in range(self.colNum):
			self.write_table.write(6, i, self.keys[i], style)
			self.write_table.col(j).height = self.width
		for i in list:
			for j in range(self.colNum):
				if i[0] == '合计':
					self.write_table.write(n, j, i[j], style2)
				else:
					self.write_table.write(n, j, i[j], style)
			n += 1
		try:
			self.worksheet.save('new.xls')
		except PermissionError:
			print('转换失败')
			print('请关闭文件再试')

	def get_call(self):
		list = []
		for (rlow, rhigh, clow, chigh) in self.merged:
			dict = {}
			data = self.table.cell_value(rlow, clow)
			dict['data'] = data
			dict['orgin'] = (rlow, rhigh, clow, chigh)
			list += [dict, ]
		return list

	def read(self):
		l = []
		for i in self.get_call():
			if i.get('data').find('合计金额(大写)') >= 0:
				row = i.get('orgin')[0]
				break
		for i in self.get_call():
			if i.get('data').find('客户名称') >= 0:
				for j in range(1, self.rowNum):
					c_cell = self.table.cell_value(i.get('orgin')[0], j)
					if c_cell != '':
						l += [{'orgin': (i.get('orgin')[0], j), 'data': c_cell}]
						break
		for i in range(row, self.rowNum):
			for j in range(self.colNum):
				c_cell = self.table.cell_value(i, j)
				if c_cell != '':
					dict = {}
					dict['data'] = c_cell
					dict['orgin'] = (i, j)
					l += [dict, ]
		return l

	def write_file(self, names):
		font = xlwt.Font()
		font.height = 20 * 15
		font.name = '宋体'
		borders = xlwt.Borders()
		borders.left = xlwt.Borders.MEDIUM
		borders.right = xlwt.Borders.MEDIUM
		borders.top = xlwt.Borders.MEDIUM
		borders.bottom = xlwt.Borders.MEDIUM
		font2 = xlwt.Font()
		font2.height = 20 * 11
		font2.name = '宋体'
		font3 = xlwt.Font()
		font3.height = 20 * 11
		font3.bold = True
		font3.name = '宋体'
		font4 = xlwt.Font()
		font4.height = 20 * 20
		font4.name = '宋体'
		font4.bold = True
		font5 = xlwt.Font()
		font5.height = 20 * 16
		font5.name = '宋体'
		font5.bold = True

		font6 = xlwt.Font()
		font6.height = 20 * 11
		font6.name = '宋体'
		font6.bold = True

		alignment2 = xlwt.Alignment()
		# 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
		alignment2.horz = 0x01
		# 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
		alignment2.vert = 0x01
		font7 = xlwt.Font()
		font7.height = 20 * 14
		font7.name = '宋体'
		font7.bold = True

		alignment = xlwt.Alignment()
		alignment.horz = 0x02
		alignment.vert = 0x01
		style = xlwt.XFStyle()
		style.alignment = alignment
		style.font = font

		style4 = xlwt.XFStyle()
		style4.alignment = alignment
		style4.font = font4

		style5 = xlwt.XFStyle()
		style5.alignment = alignment
		style5.font = font5

		style6 = xlwt.XFStyle()
		style6.alignment = alignment
		style6.font = font6

		style7 = xlwt.XFStyle()
		style7.alignment = alignment
		style7.font = font7
		style7.alignment=alignment2

		style3 = xlwt.XFStyle()
		style3.alignment = alignment
		style3.font = font
		style3.borders = borders

		style2 = xlwt.XFStyle()
		style2.font = font2
		# 写入已经合并的单元格
		for i in self.get_call():
			a, b, c, d = i.get('orgin')
			if i.get('data').find('合计金额(大写') != -1 or i.get('data').find('小写金额') != -1:
				a += ((len(names) + 1) * 2)
				b += ((len(names) + 1) * 2)
				self.write_table.write_merge(a, b - 1, c, d - 1, i.get('data'), style3)
			else:
				if a==0:
					style_=style4
				elif a==3:
					style_ = style5
				elif a==1 or a==2:
					style_ = style6
				elif a==5:
					style_ = style7
				else:
					style_=style
				self.write_table.write_merge(a, b - 1, c, d - 1, i.get('data'), style_)
			self.write_table.col(a).height = self.width

		for i in self.read():
			a, b = i.get('orgin')
			ls = ['注：请核对无误后', '（签字盖章）', '日期：']
			state = False
			for l in ls:
				if i.get('data').find(l) != -1:
					state = True
					break
			if state:
				a += ((len(names) + 1) * 2)
			self.write_table.write(a, b, i.get('data'), style2)
			self.write_table.col(a).height = self.width


if __name__ == "__main__":
	data_path = "999.xlsx"
	sheetname = "对账单"

	names = ['婴宝', '乐居', '海康', '耀锋','嘉宜','崴光']

	get_data = ExcelData(data_path, sheetname)
	# datas = get_data.readExcel()
	# datas = get_data.write_file(names)
	datas = get_data.solve_data(names)
