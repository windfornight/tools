# -*- coding: utf-8 -*-


import xlrd
import xlwt
from xlutils.copy import copy

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

from langconv import *  
  
def simple2tradition(line):  
    #将简体转换成繁体  
    line = Converter('zh-hant').convert(line.decode('utf-8'))  
    #line = line.encode('utf-8')  
    return line  


conSheetConf = [{'name':'1', 'txt':1, 'ver':3}, {'name':'2', 'txt':1, 'ver':2}, {'name':'3', 'txt':1, 'ver':2}, {'name':'4', 'txt':1, 'ver':2}, {'name':'5', 'txt':1, 'ver':2}, 
	{'name':'6', 'txt':1, 'ver':2}, {'name':'7', 'txt':1, 'ver':2}, {'name':'8', 'txt':1, 'ver':3}, {'name':'9', 'txt':1, 'ver':2}, {'name':'10', 'txt':1, 'ver':2}, 
	{'name':'11', 'txt':1, 'ver':3}, {'name':'12', 'txt':1, 'ver':2}, {'name':'13', 'txt':1, 'ver':2}, {'name':'14', 'txt':1, 'ver':3}]

fileRec = 'convRecord.txt'

def convXmls(readFile, wrFile, conVersion):
	book = xlrd.open_workbook(readFile)
	file = xlrd.open_workbook(wrFile) # formatting_info=True
	wrBook = copy(file)
	wrBook.encoding = 'utf-8'
	fileRecord = open(fileRec, 'w')

	for info in conSheetConf:
		sheetName = info['name']
		sheet = book.sheet_by_name(sheetName)
		wrSheet = wrBook.get_sheet(sheetName)
		for rx in range(sheet.nrows):
			if (sheet.ncols > info['ver']) and (sheet.cell_value(rx, info['ver']) == conVersion):
				for cy in range(sheet.ncols):
					if cy == info['txt']: #转换繁体
						str = simple2tradition(sheet.cell_value(rx, cy))
						wrSheet.write(rx, cy, str)
					else:
						wrSheet.write(rx, cy, sheet.cell_value(rx, cy))
				fileRecord.write('{0} : {1} \n'.format(sheetName, rx+1))
	wrBook.save(wrFile)
	fileRecord.close()
	
def main():
	fileName = sys.argv[1]
	outName = sys.argv[2]
	conVer = sys.argv[3]
	convXmls(fileName, outName, conVer)
	
if __name__=="__main__":
	main()