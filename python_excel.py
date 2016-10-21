
#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: lfl
# @Date:   2016-09-13 20:01:38
# @Last Modified by: lfl
# @Last Modified time: 2016-10-21 15:51:24
from collections import OrderedDict
import xlrd,sys
import xlwt


class PyExcel:
    def __init__(self, OldFile, NewFile):
        self.DataDict = OrderedDict()
        self.OldFile = OldFile
        self.NewFile = NewFile
        try:
            self.data = xlrd.open_workbook(self.OldFile)
        except Exception,e:
            print e.message

    def ReadTableExcel(self):
        for name in self.data.sheet_names():
            table = self.data.sheet_by_name(name)
            nrows = table.nrows
            for i in range(nrows):
                self.DataDict.setdefault(name,[]).append(table.row_values(i))
        return self.DataDict

    def WriteExcel(self):
        self.DataDict = self.ReadTableExcel()
        sheet_names = self.DataDict.keys()
        table = xlwt.Workbook()
        sheet_list = []
        for name in sheet_names:
            sheet_list.append(table.add_sheet(name))
        for sheet in sheet_list:
            for i in range(len(self.DataDict[sheet.name])):
                for j in range(len(self.DataDict[sheet.name][i])):
                    sheet.write(i,j,self.DataDict[sheet.name][i][j])
        table.save(self.NewFile)


def open_excel(file):
	try:
		data = xlrd.open_workbook(file)
		return data
	except Exception,e:
		print e.message


def excel_table_buindex(file, colindex=0, by_index=0):
    data = open_excel(file)
    data_dict = OrderedDict()
    for name in data.sheet_names():
        table = data.sheet_by_name(name)
        nrows = table.nrows
        for i in range(nrows):
            data_dict.setdefault(name,[]).append(table.row_values(i))
    return data_dict

def ExceWrite(data_dict):
    sheet_names = data_dict.keys()
    table = xlwt.Workbook()
    sheet_list = []
    for name in sheet_names:
        sheet_list.append(table.add_sheet(name))
    
    for sheet in sheet_list:
        for i in range(len(data_dict[sheet.name])):
            for j in range(len(data_dict[sheet.name][i])):
                sheet.write(i,j,data_dict[sheet.name][i][j])
    table.save("config-test.xls")
   

if __name__ == "__main__":
    old = "config.xlsx"
    new = "config-test.xls"
    e = PyExcel(old,new)
    e.WriteExcel()

