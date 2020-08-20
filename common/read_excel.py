'''
author:songjianhao
time:2020-04-03
'''

import xlrd

class read_excel():

    '''读取Excel数据，并返回列表，用于测试框架'''

    def __init__(self, file_name, sheet_name = 'sheet1'):
        self.data = xlrd.open_workbook(file_name)
        self.table = self.data.sheet_by_name(sheet_name)

        self.nrows = self.table.nrows
        self.ncols = self.table.ncols

    def read_data(self):
        if self.nrows > 1:
            '''获取第一行数据作为key，返回列表格式'''
            keys = self.table.row_values(0)
            list_apidata = []
            for col in range(1, self.nrows):
                values = self.table.row_values(col)
                api_dict = dict(zip(keys , values))
                list_apidata.append(api_dict)
            return list_apidata
        else:
            print('该文件内容为空，请检查文件是否正确！')
            return None
