from PyQt5.QtCore import Qt,QTimer,QThread,pyqtSignal
from PyQt5.QtWidgets import QMessageBox, QWidget, QApplication, QDialog
from new import Ui_Form
import os,xlrd,xlsxwriter
from xlrd import xldate_as_tuple
from xlrd.biffh import XLRDError
import datetime
from xlsxwriter.exceptions import FileCreateError
import re
from collections import Counter

import jieba
jieba.set_dictionary('dict/dict.txt')
jieba.initialize()


class ExcelData(QThread):
	data_signal = pyqtSignal(dict)
	# 初始化方法
	def __init__(self, data_path, sheetname,newname):
		super(ExcelData, self).__init__()
		# 定义一个属性接收文件路径
		self.data_path = data_path
		# 定义一个属性接收工作表名称
		self.sheetname = sheetname
		self.newname = newname
		# 使用xlrd模块打开excel表读取数据
		try:
			self.data = xlrd.open_workbook(self.data_path)
			self.state = True
		except XLRDError:
			self.state=False
			return
		except FileNotFoundError:
			self.state=False
			return
		# 根据工作表的名称获取工作表中的内容（方式①）
		try:
			self.table = self.data.sheet_by_name(self.sheetname)
			self.state = True
		except XLRDError:
			self.state=False
			return
		# 根据工作表的索引获取工作表的内容（方式②）
		# self.table = self.data.sheet_by_name(0)
		# 获取第一行所有内容,如果括号中1就是第二行，这点跟列表索引类似
		try:
			self.keys = self.table.row_values(6)
			self.state = True
		except IndexError:
			self.state=False
			return
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
		print('初始化完成')

	def check(self,names):
		try:
			if not self.state:
				return
			datas=self.readExcel()
			_names=[]
			for data in datas:
				for name in names:
					if data.get('产品名称').find(name)!=-1:
						_names+=[name,]
			self.state = True
			return list(set(_names))
		except AttributeError :
			self.state = False

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
	def solve_data(self):
		self.write_file()
		data = self.readExcel()
		if not self.names:
			return None
		list = []
		for name in self.names:
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
		self.worksheet.close()


	#获取到合并单元格对象
	def get_call(self):
		list = []
		for (rlow, rhigh, clow, chigh) in self.merged:
			dict = {}
			data = self.table.cell_value(rlow, clow)
			if data.find('小写金额')!=-1:
				data='小写金额：'+str(round(float(data.split('小写金额：')[-1]),2))
			dict['data'] = data
			dict['orgin'] = (rlow, rhigh, clow, chigh)
			list += [dict, ]
		# print(list)
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
	def write_file(self):
		style = self.worksheet.add_format({
		'align': 'center',  # 水平居中
		'valign': 'vcenter',  # 垂直居中
		'font': u'宋体',  # 字体
		})
		style.set_font_size(15)
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
				a += ((len(self.names) + 1) * 2)
				b += ((len(self.names) + 1) * 2)
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
					style_= style
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
				a += ((len(self.names) + 1) * 2)
			self.write_table.write(a, b, i.get('data'))
			self.write_table.set_row(a, self.width)

	def run(self):
		try:
			try:
				self.solve_data()
				self.data_signal.emit({'state':1,'data':'转换完成'})
			except FileCreateError:
				self.data_signal.emit({'state':0,'data':'未关闭xlsx文件，关闭后再试'})
		except BaseException as e:
			self.data_signal.emit({'state':0,'data':str(e)+'\n           选择的文件内部格式错误\n         请重新选择'})

	# 获取到词语
	def get_pariciple(self,max_num=4):
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
		dic = sorted(dic.items(),key= lambda x:x[1],reverse=True)[:max_num]
		key=[x[0] for x in dic]
		# print(key)
		return '\n'.join(key)






class new_pane(QWidget, Ui_Form):
	def __init__(self, parent=None, *args, **kwargs):
		super().__init__(parent, *args, **kwargs)
		self.setAttribute(Qt.WA_StyledBackground, True)
		self.setupUi(self)


		self.timer=QTimer()
		self.timer.timeout.connect(lambda : self.check())
		self.timer.timeout.connect(lambda : self.timer.stop())
		self.timer.start(500)

		self.sheetname = "对账单"
		self.way = True



	def start(self):
		print('start')
		if self.lineEdit.text()=='':
			self.lineEdit.setText('new_'+os.path.splitext(self.comboBox.currentText())[0])

		if not self.way:
			# self.change_('')
			QMessageBox.about(self, '失败', '导入的文件有问题，无法识别')
			return

		data = self.textEdit.toPlainText()
		datas = [i for i in data.split('\n') if i != '']
		if datas == []:
			QMessageBox.about(self, '提示', '请输入关键词')
			return
		names = self.Excel.check(datas)
		if names == []:
			QMessageBox.about(self, '提示', '没有在文件中发现输入的关键词\n       请重新输入')
			return
		if not names:
			QMessageBox.about(self, '失败', '导入的文件有问题，无法识别')
			return
		if len(names) != len(datas):
			l=[data for data in datas if data not in names]
			if l != []:
				string='\n'.join(l)
			QMessageBox.about(self, '提示', '发现文件中不存在的关键词:\n'+string)
		# print(names)
		self.Excel.names =names
		self.Excel.start()

	def change_(self,string):
		# print('change_')
		self.Excel=ExcelData(self.comboBox.currentText(),self.sheetname,self.lineEdit.text()+'.xlsx')

		if not self.Excel.state:
			self.way=False
			return
		else:
			self.way=True
		self.Excel.data_signal.connect(self.deal)

	def deal(self,dic):
		self.change_('')
		if dic.get('state'):
			QMessageBox.about(self, '成功',dic.get('data') )
		else:
			QMessageBox.about(self, '失败',dic.get('data') )

	def change(self,string):
		print(string)
		if string=='没有读取到文件':
			self.pushButton.setEnabled(False)
			self.pushButton_3.setEnabled(False)
			return
		self.pushButton.setEnabled(True)
		self.pushButton_3.setEnabled(True)
		self.lineEdit.setText('new_'+os.path.splitext(string)[0])



	def check(self):
		print('check')
		self.comboBox.clear()
		self.names=[i for i in os.listdir() if os.path.splitext(i)[-1] in ('.xlsx','.xls')]

		if self.names == []:
			self.names = None
			self.comboBox.addItem('')
			self.comboBox.setItemText(0, '没有读取到文件')
			self.lineEdit.setText('None')
			self.pushButton_3.setEnabled(False)
			self.pushButton.setEnabled(False)
			QMessageBox.about(self, '警告', '没有读取到文件')
			return
		n=0
		for name in self.names:
			self.comboBox.addItem('')
			self.comboBox.setItemText(n, name)
			n+=1
		os.path.splitext(self.names[0])[0]
		self.lineEdit.setText('new_'+os.path.splitext(self.names[0])[0])
		self.pushButton.setEnabled(True)
		self.pushButton_3.setEnabled(True)
		QMessageBox.about(self, '提示', '获取成功')
		# self.change(self.lineEdit.text())

	def check_jieba(self):
		print('check_jieba')

		text=self.Excel.get_pariciple(int(self.spinBox.text()))
		# print(text)
		self.textEdit.setText(text)


if __name__ == "__main__":
	import sys
	try:
		app = QApplication(sys.argv)
		window = new_pane()
		window.show()

	except BaseException as e:
		QMessageBox.about(QDialog(), '错误', e)

	except AttributeError as e:
		QMessageBox.about(QDialog(), '错误', e)
	sys.exit(app.exec_())
