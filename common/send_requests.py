'''
author:songjianhao
date:2020.04.05
'''

import json


class send_requests():
    '''发送请求数据'''

    def send_requests(self, s, apidata):
        try:
            '''读取的Excel文件内容作为传参'''
            method = apidata["method"]
            url = apidata["url"]
            if apidata["params"] == "":
                params = None
            else:
                params = eval(apidata["params"])

            if apidata["headers"] == "":
                headers = None
            else:
                headers = eval(apidata["headers"])

            if apidata["data"] == "":
                data = None
            else:
                data = eval(apidata["data"])

            type = apidata["type"]
            v = False
            if type == "data":
                body = data
            elif type == "json":
                body = json.dumps(data)
            else:
                body = data

            re = s.request(method=method, url=url, headers=headers, params=params, data=body, verify=v)
            return re

        except Exception as e:
            print(e)
