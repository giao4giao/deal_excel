import os




names=[i for i in os.listdir() if os.path.splitext(i)[-1] in ('.xlsx','.xls')]
print(names)
if nemas == []:
	names = None
print(names)