from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfdevice import PDFDevice
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LAParams
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed
#正则模块，主要用到的函数包括：
#	1、re.sub(a, b, s):
#		替换函数，把s中的a替换成b，返回新字符串s'
#		a: 需要替换的对象
#		b: 替换成
#		s: 需要进行替换的字符串
#	2、re.split(a, s):
#		分割函数，把s按照a进行拆分，得到删掉a之后的列表，返回列表
#		a: 拆分依据，拆分之后a会消失
#		s: 需要进行拆分的字符串
import re

#pandas: 将字典转化为DataFrame格式，并以excel形式保存
import pandas as pd

#defaultdict: 创建默认字典
from collections import defaultdict

#os
#	1、os.remove(file):
#		删除file文件
#	2、os.mkdir(dir_path):
#		在dir_path创建文件夹，文件夹名字包含在dir_path中
#		例如'./a'，即在当前目录创建名为a的文件夹
#	3、os.listdir(path):
#		获取path路径下的所有文件名称，用于获取当前路径下所有的PDF名称并打开进行处理
import os


#parse: 解析PDF文件，将PDF文件转换为纯文本的txt，但是包含一些页码等不规则文本
#	DataIO: PDF文件路径，例如'./问询函.pdf'，即当前目录下的【问询函】PDF文件
#	save_path: 将解析后的txt文本储存在save_path中
#		例如'./a.txt'，即将解析后的txt文件储存在当前目录下的a.txt文件中
#	代码细节不需要了解
def parse(DataIO, save_path):
	parser = PDFParser(DataIO)
	doc = PDFDocument()
	parser.set_document(doc)
	doc.set_parser(parser)
	doc.initialize()

	if not doc.is_extractable:
		raise PDFTextExtractionNotAllowed
	else:
		rsrcmagr = PDFResourceManager()
		laparams = LAParams()
		device = PDFPageAggregator(rsrcmagr, laparams=laparams)
		interpreter = PDFPageInterpreter(rsrcmagr, device)

		for page in doc.get_pages():
			interpreter.process_page(page)
			layout = device.get_result()
			for x in layout:
				try:
					if(isinstance(x, LTTextBoxHorizontal)):
						with open('%s'%(save_path), 'a') as f:
							result = x.get_text()
							#print(result)
							f.write(result)
				except:
					print('failed')


#pdf2txt: 将PDF解析为txt，清洗txt文本并储存在r.txt中
#	pdf_path: PDF文件的路径
def pdf2txt(pdf_path):

	#打印提示信息，PDF文件正在解析
	print(pdf_path, ' converted to txt')

	#以二进制格式打开PDF文件并利用parse函数将PDF文件解析为初始txt并保存在当前路径下的a.txt文件中
	with open(pdf_path, 'rb') as f:
		parse(f, r'./a.txt')

	#处理刚刚保存的初始txt文件a.txt，里面有一些页码信息需要删除
	with open('./a.txt', 'r') as f:
		t = f.read()

	#根据初始txt文件中的格式信息，可以发现
	#页码信息以【回车 + 数字 + 空格 + 数字】的形式出现
	#利用re.sub可以将初始txt文件中的【页码信息】替换为【''】，即空值，相当于删除
	r = re.sub('\n[0-9] \n', '', t)
	#对于具有两个或三个数字的页码，同样利用re进行删除处理
	r = re.sub('\n[0-9][0-9] \n', '', r)
	r = re.sub('\n[0-9][0-9][0-9] \n', '', r)

	#将处理后的文本r储存在当前目录下的r.txt文件中
	with open('./r.txt', 'w') as f:
		f.write(r)
	
	#将初始txt文件【a.txt】利用os.remove函数删除
	os.remove(r'./a.txt')

	#打印提示信息，PDF文件已经处理完毕
	print(pdf_path, ' is converted!')


#questions_extraction
#将每一个大类下的小问题进行拆分并提取
#	question_txt: 每一大类下的小问题文本，类型为字符串
def questions_extraction(question_txt):

	#通过分析question_txt文本可以发现
	#大类下小问题文本里均为【问题1. XXXX】/【问题2. XXXX】类似的形式出现
	#因此可以通过对【'问题' + 数字】这一类格式将文本进行拆分，从而获取大类下的每一个小问题
	#q_list即为小问题列表
	#\n代表回车
	#	例如[XXXX（此处为无用信息）, '关于收入确认\nXXXXX', 关于营业收入\nXXXX', ...]
	q_list = re.split('问题[0-9][0-9].|问题[0-9].', question_txt)

	#构建默认字典，字典的key为每个小问题的标题，value为每个小问题的具体内容
	#即将q_list=[XXXX（此处为无用信息）, '关于收入确认\nXXXXX', 关于营业收入\nXXXX', ...]
	#转变为c = {'关于收入确认': 'XXXXX', '关于营业收入': 'XXXX'}
	c = defaultdict(str)

	#根据q_list的内容，可以发现我们需要的问题信息出现在第二个位置及其之后
	#在列表中对应的index即为从【1】到【len(q_list)】（列表长度）
	#选取位置为1到最后的q_list信息
	#	即q_list[i]，例如q_list[i] = '关于收入确认\nXXXXX'
	for i in range(1, len(q_list)):
		#按照回车进行拆分，选取第一个元素（index为0），即为当前小问题的标题
		#	例如w = '关于收入确认'
		w = q_list[i].split()[0]
		#将q_list[i]中的【标题】与【回车或空格】替换为空，即可获得小问题的具体内容
		#例如t = 'XXXXX'
		t = re.sub(w, '', q_list[i])
		t = re.sub(r'\s', '', t)
		#构建字典c[w] = t
		#	例如{'关于收入确认': 'XXXXX'}
		c[w] = t

	#返回字典c
	return c

#txt2excel: 获得到的文本信息处理为excel并保存
#	file_path:文本信息即处理过的r.txt
def txt2excel(file_path):

	#打开r.txt文件，并按照utf-8格式进行读取
	with open(file_path, 'rb') as f:
		txt = f.read().decode('utf-8')

	#将txt中的大类标题进行拆分，获得txt_list列表
	#txt_list中包括五个元素，分别为【上市公司信息】，【规范性问题】，【信息披露问题】，【财务会计问题】，【其他问题】
	#需要注意的是某些文件可能会出现没有【其他问题】这一部分的情况，即拆分后的列表仅有四个元素
	#利用re.sub进行拆分
	txt_list = re.split('一、规范性问题|二、信息披露问题|三、与财务会计资料相关的问题|四、其他问题|一、 规范性问题|二、 信息披露问题|三、 与财务会计资料相关的问题|四、 其他问题', txt)

	#如果拆分出的列表的元素个数不等于5，即出现了非正常格式的文件，打印相关信息以便于后续处理
	if len(txt_list) != 5:
		print('*'*20)
		print(file_path)
		print('*'*20)

	#获取公司名称和保荐券商并作为文件夹名称
	#将txt_list的第一部分，即txt_list[0]，按照【英文冒号】或【中文冒号】进行拆分
	#拆分后的第一个元素即为【上市公司名称 + '并' + 保荐券商名称】
	#	例如: gs_qs = '中航富士达科技股份有限公司并招商证券股份有限公司、中航证券有限公司'
	gs_qs = re.split(':|：', txt_list[0])[0]
	#将gs_qs依据【并】进行拆分
	#第一个元素即为公司名称，第二个元素即为保荐券商名称
	#此时可能由于格式问题，名称中包含【空格】或【回车】，需要利用re模块把【空格】或【回车】删除
	gs = gs_qs.split('并')[0]
	gs = re.sub('\s', '', gs)
	qs = gs_qs.split('并')[1]
	qs = re.sub('\s', '', qs)
	#得到公司和券商名称，并打印信息
	#	例如：中航富士达科技股份有限公司 by 招商证券股份有限公司、中航证券有限公司 is processing
	print(gs, ' by ', qs, ' is processing')

	#利用questions_extraction模块获取【规范性问题】大类下的小问题，返回字典c_gf
	c_gf = questions_extraction(txt_list[1])

	#信息披露问题
	c_xx = questions_extraction(txt_list[2])

	#财务问题
	c_cw = questions_extraction(txt_list[3])

	#其他问题
	#由于某些文档中可能出现没有【其他问题】这一部分，所以利用try-except结构进行试错
	#	如果可以获得txt_list[4]，即存在【其他问题】这一部分，那么进行处理，但需要删除最后的无用信息
	#	无用信息为：'除上述问题外， XXXXX'，可利用re模块进行拆分并删除，获取c_qt这一【其他问题】部分的字典
	#	如果不能获得txt_list[4]，即不存在【其他问题】这一部分，那么c_qt为空字典
	try:
		qt_list = re.split('问题[0-9][0-9].|问题[0-9].', txt_list[4])
		c_qt = defaultdict(str)
		for i in range(1, len(qt_list)):
			w = qt_list[i].split()[0]
			t = re.sub(w, '', qt_list[i])
			t = re.sub(r'\s', '', t)
			if i == len(qt_list) - 1:
				t = re.split('除上述问题外', t)[0]
			c_qt[w] = t
		#print(c_qt)
	except:
		c_qt = defaultdict(str)
	
	#当前PDF文件夹名称为dir_path
	#	例如'中航富士达科技股份有限公司_招商证券股份有限公司、中航证券有限公司'
	dir_path = './' + gs + '_' + qs
	#如果当前目录下不存在此文件夹，那么利用os.mkdir进行创建
	if not os.path.exists(dir_path):
		os.mkdir(dir_path)

	#在已创建的文件夹中储存各类小问题
	df_gf = pd.DataFrame(dict(c_gf), index=[0])
	df_gf.T.to_excel(dir_path + '/规范性问题.xlsx')

	df_xx = pd.DataFrame(dict(c_xx), index=[0])
	df_xx.T.to_excel(dir_path + '/信息披露问题.xlsx')

	df_cw = pd.DataFrame(dict(c_cw), index=[0])
	df_cw.T.to_excel(dir_path + '/财务问题.xlsx')

	df_qt = pd.DataFrame(dict(c_qt), index=[0])
	df_qt.T.to_excel(dir_path + '/其他问题.xlsx')

	#打印提示信息，已完成当前PDF分析
	print('done')
	print('='*10)

if __name__ == '__main__':

	#pdf_list: 当前路径下的所有需要处理的PDF文件名称
	pdf_list = []

	#file_list: 当前路径下的所有文件名称
	file_list = os.listdir('./')

	#如果文件名称的后缀为.pdf，那么放入pdf_list中
	for i in file_list:
		if i.endswith('.pdf'):
			pdf_list.append(i)

	#打印提示信息，即需要处理的所有PDF文件名称
	print(pdf_list)

	#利用pdf2txt, txt2excel函数进行处理
	#	i: PDF文件名称，例如【问询函.pdf】
	for i in pdf_list:
		pdf2txt(i)
		txt2excel(r'r.txt')











