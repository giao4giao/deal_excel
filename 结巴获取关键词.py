import xlsxwriter
import xlrd
import time
import os
from xlrd import xldate_as_tuple
import datetime
import math
import jieba
import re
from collections import Counter

'''
xlrd中单元格的数据类型
数字一律按浮点型输出，日期输出成一串小数，布尔型输出0或1，所以我们必须在程序中做判断处理转换
成我们想要的数据类型
0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
'''


class ExcelData():
	# 初始化方法
	def __init__(self, data_path, sheetname,newname):
		# 定义一个属性接收文件路径
		self.data_path = data_path
		# 定义一个属性接收工作表名称
		self.sheetname = sheetname
		# 使用xlrd模块打开excel表读取数据
		try:
			self.data = xlrd.open_workbook(self.data_path)
		except FileNotFoundError:
			print('未发现源文件\n')
			print('十秒后自动关闭。')
			time.sleep(10)
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

		self.width = 20
		self.worksheet = xlsxwriter.Workbook(newname)
		# 设置默认的字的大小
		self.worksheet.formats[0].set_font_size(11)
		self.write_table = self.worksheet.add_worksheet(self.sheetname)


	def check(self,names):
		datas=self.readExcel()
		_names=[]
		for data in datas:
			for name in names:
				if data.get('产品名称').find(name)!=-1:
					_names+=[name,]
		return list(set(_names))

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



	# 数据处理
	def solve_data(self, names,newname):
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
		# 这里有个大坑，就是有可能两次数据一样导致被去掉
		for i in datas:
			if i not in list :
				l += [i, ]
				list += [i, ]
				n=i
			elif n==i:
				# print(i)
				print('\n在',i[0],'的时候有两条数据是一样的，这不是为难我吗\n')
				list += [i, ]

		black = ('', '', '', '', '', '', '', '', '')
		list += [(('合计', '', '', '', '', '', '', sum([float(i[-2]) for i in l]), '')), black]

		# print(len(list))
		# print(len(datas))
		# for i in (list):
		# 	if i not in datas:
		# 		print(i)

		style2 = self.worksheet.add_format({
		'border': 1,  # 边框
		'align': 'center',  # 水平居中
		'valign': 'vcenter',  # 垂直居中

		'bold': False,  # 加粗（默认False）
		'font': u'宋体',  # 字体
		'fg_color': '#FFFF00',  # 背景色
		'color': 'red'  # 字体颜色
		})
		style = self.worksheet.add_format({
		'border': 1,  # 边框
		'align': 'center',  # 水平居中
		'valign': 'vcenter',  # 垂直居中

		'bold': False,  # 加粗（默认False）
		'font': u'宋体',  # 字体
		# 'fg_color': '#FFFF00',  # 背景色
		# 'color': 'green'  # 字体颜色
		})
		style6 = self.worksheet.add_format({
		'border': 1,  # 边框
		'align': 'center',  # 水平居中
		'valign': 'vcenter',  # 垂直居中

		'bold': True,  # 加粗（默认False）
		'font': u'宋体',  # 字体
		# 'fg_color': '#FFFF00',  # 背景色
		# 'color': 'green'  # 字体颜色
		})
		n = 7
		for j in range(self.colNum):
			num = max([len(tuple((str(i[j]).replace('.',''))))  if type(list[0][j]) == type('') else len(tuple((str(i[j]).replace('.','')))) for i in list])
			num = (num if num > 5 else 5) * 1.5
			self.write_table.set_row(j, self.width)
			self.write_table.set_column(j,j,num)
		# 写入前信息 7行
		for i in range(self.colNum):
			self.write_table.write(6, i, self.keys[i], style6)
			self.write_table.set_row(i, self.width)
		for i in list:
			for j in range(self.colNum):
				if i[0] == '合计':
					self.write_table.write(n, j, i[j], style2)
				else:
					self.write_table.write(n, j, i[j], style)
			n += 1
		try:
			self.worksheet.close()
			return True
		except PermissionError:
			print('转换失败')
			print('请关闭文件再试')


	#获取到合并单元格对象
	def get_call(self):
		list = []
		for (rlow, rhigh, clow, chigh) in self.merged:
			dict = {}
			data = self.table.cell_value(rlow, clow)
			dict['data'] = data
			dict['orgin'] = (rlow, rhigh, clow, chigh)
			list += [dict, ]
		return list


	# 读出主要信息
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


	# 先行写入一些的内容
	def write_file(self, names):
		style3 = self.worksheet.add_format({
		'border': 1,  # 边框
		'align': 'center',  # 水平居中
		'valign': 'vcenter',  # 垂直居中

		# 'bold': True,  # 加粗（默认False）
		'font': u'宋体',  # 字体
		# 'fg_color': '#FFFF00',  # 背景色
		# 'color': 'green'  # 字体颜色
		})
		style3.set_font_size(15)
		style4 = self.worksheet.add_format({
		# 'border': 1,  # 边框
		'align': 'center',  # 水平居中
		'valign': 'vcenter',  # 垂直居中

		'bold': True,  # 加粗（默认False）
		'font': u'宋体',  # 字体
		# 'fg_color': '#FFFF00',  # 背景色
		# 'color': 'green'  # 字体颜色
		})
		style4.set_font_size(20)
		style5 = self.worksheet.add_format({
		# 'border': 1,  # 边框
		'align': 'center',  # 水平居中
		'valign': 'vcenter',  # 垂直居中

		'bold': True,  # 加粗（默认False）
		'font': u'宋体',  # 字体
		# 'fg_color': '#FFFF00',  # 背景色
		# 'color': 'green'  # 字体颜色
		})
		style5.set_font_size(16)
		style6 = self.worksheet.add_format({
		# 'border': 1,  # 边框
		'align': 'center',  # 水平居中
		'valign': 'vcenter',  # 垂直居中

		'bold': True,  # 加粗（默认False）
		'font': u'宋体',  # 字体
		# 'fg_color': '#FFFF00',  # 背景色
		# 'color': 'green'  # 字体颜色
		})
		style7 = self.worksheet.add_format({
		# 'border': 1,  # 边框
		'align': 'left',  # 水平居中
		'valign': 'vcenter',  # 垂直居中

		'bold': True,  # 加粗（默认False）
		'font': u'宋体',  # 字体
		# 'fg_color': '#FFFF00',  # 背景色
		# 'color': 'green'  # 字体颜色
		})
		style7.set_font_size(14)

		# 写入已经合并的单元格
		for i in self.get_call():
			a, b, c, d = i.get('orgin')
			# print(a, c, b-1, d-1)
			# self.write_table.merge_range(a, c, b-1, d-1 , i.get('data'),style7)
			
			if i.get('data').find('合计金额(大写') != -1 or i.get('data').find('小写金额') != -1:
				a += ((len(names) + 1) * 2)
				b += ((len(names) + 1) * 2)
				# print(a, c, b-1, d-1)
				self.write_table.merge_range(a, c, b-1, d-1 , i.get('data'),style3)
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
				self.write_table.merge_range(a, c, b-1, d-1 , i.get('data'),style_)
			self.write_table.set_row(a, self.width)

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
			self.write_table.write(a, b, i.get('data'))
			self.write_table.set_row(a, self.width)

	# 获取到词语
	def get_pariciple(self):
		chinese = '[\u4e00-\u9fa5]+'
		datas=self.readExcel()
		l=[]
		for data in datas:
			l+=[data.get('产品名称'),]
		ls=[]
		l =[' '.join(re.findall(chinese,i)) for i in l]
		for i in l:
			# ls+=[jieba.lcut(i,cut_all=True),]
			ls+=[jieba.lcut(i),]
		# print(ls)
		l=[]
		for s in ls:
			for i in s:
				if i not in (' ','挂版','沙发') and len(i)>=2:
					l+=[i,]
		# l=list(set(l))
		# print(Counter(l))
		dic=Counter(l)
		dic = sorted(dic.items(),key= lambda x:x[1],reverse=True)[:4]
		key=[x[0] for x in dic]
		print(key)



# 读取txt文档内容
def read_txt():
	name='deal.txt'
	if not os.path.isfile(name):
		with open(name,'w',encoding='gbk')as f:
			f.write('name=999.xlsx\nname_new=new.xlsx\n在下面写区分的关键字，下面一行一个(请不要改这行字)\n')
		print('\n第一次打开，未发现',name,'文件,修改后重新打开\n')
		print('十秒后自动关闭。')
		time.sleep(10)
		exit()
	with open(name, 'r', encoding='gbk')as f:
		datas=[i.replace('\n','').replace(' ','') for i in f if i.replace('\n','').replace(' ','') not in ('','在下面写区分的关键字，下面一行一个(请不要改这行字)')]
	# print(datas)
	dic,l={},[]
	n=0
	for data in datas:
		if data.find('name=') !=-1:
			dic['name']=data.replace('name=','',1)
			n+=1
		elif data.find('name_new=') !=-1:
			dic['name_new'] = data.replace('name_new=', '', 1)
			n+=1
		else:
			l+=[data,]
	if n != 2:
		with open(name,'w',encoding='gbk')as f:
			f.write('name=999.xlsx\nname_new=new.xlsx\n在下面写区分的关键字，下面一行一个(请不要改这行字)\n')
		print('\n发现',name,'文件格式有错误，已经初始化文件\n')
		print('十秒后自动关闭。')
		time.sleep(10)
		exit()
	dic['data']=l
	return dic






if __name__ == "__main__":
	# data_path = "999.xlsx"
	sheetname = "对账单"

	dic = read_txt()
	# print(dic)
	# names = ['婴宝', '乐居', '海康', '耀锋','嘉宜','崴光']
	data_path=dic.get('name')
	newname=dic.get('name_new')
	names=dic.get('data')


	get_data = ExcelData(data_path, sheetname,newname)
	names = get_data.check(names)

	if not names:
		print('\n请输入存在的关键字\n')
		print('十秒后自动关闭。')
	else:
		print('读取到关键字是：')
		for name in names:
			print('               ',name)
	# datas = get_data.readExcel()
	# datas = get_data.write_file(names)
	# get_data.worksheet.close()
	get_data.get_pariciple()
	'''
	datas = get_data.solve_data(names,newname)
	if datas:
		print('\n    完成\n')
	else:
		print('\n请输入筛选条件\n')
	print('十秒后自动关闭。')
	time.sleep(10)
	'''