'''
author:songjianhao
time:2020-04-03
'''

import xlrd

class read_excel():

    '''读取Excel数据，并返回字典，用于测试框架'''

    def __init__(self, file_name, sheet_name = 'sheet1'):
        self.data = xlrd.open_workbook(file_name)
        self.table = self.data.sheet_by_name(sheet_name)

        self.nrows = self.table.nrows
        self.ncols = self.table.ncols

    def read_data(self):
        if self.nrows > 1:
            pass

