import requests, json, jsonpath, traceback, ast
from urllib.parse import urlencode


class HTTP:
    def __init__(self, w, s, l):
        self.url = ''
        self.session = requests.session()
        self.result = self.response = None
        self.param = {}
        self.jsonres = {}
        self.code = 0
        self.js = jsonpath
        self.w = w
        self.s = s
        self.l = l

    def add_header(self, key, value):
        """添加请求头"""
        try:
            self.session.headers[key] = self.__get_values(value)
        except Exception as e:
            self.l.warning("没有变量：%s" % e)
            self.session.headers[key] = value
        if self.session.headers.get(key) is not None:
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
        else:
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
        try:
            self.w.sheet_write(self.w.row, self.w.col + 1, self.session.headers, self.s.res_result())
        except Exception as e:
            self.l.warning("类型错误：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col + 1, str(self.session.headers), self.s.res_result())

    def assert_equals(self, key, value):
        """测试结果断言"""
        if isinstance(key, int):
            if 200 >= self.code > 400 and key == int(value):
                self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            else:
                self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
        elif isinstance(key, str):
            if 200 >= self.code > 400 and key == str(value):
                self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            else:
                self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
        elif 200 >= self.code > 400 and key == value:
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
        else:
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
        try:
            self.w.sheet_write(self.w.row, self.w.col + 1, self.jsonres[key], self.s.res_result())
        except Exception as e:
            self.l.warning("匹配失败：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col + 1, self.extractor(key), self.s.res_result())

    def __find_values(self, s):
        """查询字典中的变量"""
        for value in self.param.values():
            s = s.find('{' + value + '}', self.param[value])
        # self.l.info("本次查找结果：%s" % s)
        return s

    def extractor(self, data):
        """提取接口的返回值"""
        value = None
        key = self.js.jsonpath(json.loads(self.response), data)
        for i in key:
            if isinstance(i, str):
                value = "".join(i)
            else:
                value = i
        if value != '' or value is not None:
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, value, self.s.res_result())
            return value
        else:
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, str(value), self.s.res_result())

    def get(self, path, payload=None, ):
        """Get请求总入口"""
        if payload == "" or payload is None:
            self.__get_url(path)
        else:
            self.__get_payload(path, self.__get_params(payload))

    @staticmethod
    def __get_params(param):
        """Get入参处理"""
        if param.startswith('{') and param.endswith('}'):
            try:
                param = ast.literal_eval(param)
            finally:
                return urlencode(param)
        else:
            return str(param)

    def __get_payload(self, path, payload=None):
        """有入参的Get请求"""
        try:
            self.result = self.session.get(self.url + path + '?' + payload, timeout=8)
            self.response = self.result.content.decode('utf-8')
            self.jsonres = json.loads(self.response)
            self.code = self.result.status_code
            self.l.info("Get请求成功：%s" % self.response)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.__write_content(self.response), self.s.res_result())
            return self.response
        except Exception as e:
            self.l.error("Get请求失败：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            try:
                self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
            except Exception as e:
                self.l.warning("类型错误：%s" % e)
                self.w.sheet_write(self.w.row, self.w.col + 1, str(traceback.format_exc()), self.s.res_result())
        finally:
            self.l.info("Get请求头：%s" % self.session.headers)
            self.l.info("Get请求参数：%s" % self.url + path + '?' + payload)

    def __get_url(self, path):
        """无入参的Get请求"""
        try:
            self.result = self.session.get(self.url + path, timeout=8)
            self.response = self.result.content.decode('utf-8')
            self.jsonres = json.loads(self.response)
            self.code = self.result.status_code
            self.l.info("Get请求成功：%s" % self.response)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.__write_content(self.response), self.s.res_result())
            return self.response
        except Exception as e:
            self.l.error("Get请求失败：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            try:
                self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
            except Exception as e:
                self.l.warning("类型错误：%s" % e)
                self.w.sheet_write(self.w.row, self.w.col + 1, str(traceback.format_exc()), self.s.res_result())
        finally:
            self.l.info("Get请求头：%s" % self.session.headers)
            self.l.info("Get请求地址：%s" % self.url + path)

    def __get_values(self, s):
        """获取变量的值"""
        for key in self.param.keys():
            s = s.replace('{' + key + '}', self.param[key])
        # self.l.info("本次替换结果：%s" % s)
        return s

    def post(self, path, payload=None, ):
        """Post请求总入口"""
        if payload == '' or payload is None:
            self.__post_url(path)
        else:
            self.__post_payload(path, payload)

    @staticmethod
    def __post_params(param):
        """Post入参处理"""
        if param.startswith('{') and param.endswith('}'):
            return ast.literal_eval(param)
        else:
            params = {}
            if len(param) >= 1:
                p = param.split('&')
                for pp in p:
                    ppp = pp.split('=')
                    params[ppp[0]] = ppp[1]
                return params

    def __post_payload(self, path, payload):
        """处理有入参的Post请求"""
        try:
            if self.session.headers.get('Content-Type') == 'application/x-www-form-urlencoded':
                self.result = self.session.post(self.url + path, data=self.__get_params(payload), timeout=26)
            elif self.session.headers.get('Content-Type') == 'application/json':
                payload = self.__post_params(payload)
                for value in payload.values():
                    if not isinstance(value, str):
                        pass
                    elif not value.startswith('{'):
                        pass
                    else:
                        for key in payload.keys():
                            payload[key] = self.__get_values(payload[key])
                self.result = self.session.post(self.url + path, data=json.dumps(payload), timeout=26)
            self.response = self.result.content.decode('utf-8')
            self.jsonres = json.loads(self.response)
            self.code = self.result.status_code
            self.l.info("Post请求成功：%s" % self.response)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.__write_content(self.response), self.s.res_result())
            return self.response
        except Exception as e:
            self.l.error("Post请求失败：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            try:
                self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
            except Exception as e:
                self.l.warning("类型错误：%s" % e)
                self.w.sheet_write(self.w.row, self.w.col + 1, str(traceback.format_exc()), self.s.res_result())
        finally:
            self.l.info("Post请求头：%s" % self.session.headers)
            self.l.info("Post请求地址：%s" % self.url + path)
            self.l.info("Post请求入参：%s" % payload)

    def __post_url(self, path):
        """处理无入参的Post请求"""
        try:
            self.result = self.session.post(self.url + path, timeout=26)
            self.response = self.result.content.decode('utf-8')
            self.jsonres = json.loads(self.response)
            self.code = self.result.status_code
            self.l.info("Post请求成功：%s" % self.response)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.__write_content(self.response), self.s.res_result())
            return self.response
        except Exception as e:
            self.l.error("Post请求失败：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            try:
                self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
            except Exception as e:
                self.l.warning("类型错误：%s" % e)
                self.w.sheet_write(self.w.row, self.w.col + 1, str(traceback.format_exc()), self.s.res_result())
        finally:
            self.l.info("Post请求头：%s" % self.session.headers)
            self.l.info("Post请求地址：%s" % self.url + path)

    def remove_header(self, key):
        """删除请求头"""
        try:
            self.session.headers.pop(key)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
        except Exception as e:
            self.l.error("删除出错：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
        try:
            self.w.sheet_write(self.w.row, self.w.col + 1, self.session.headers, self.s.res_result())
        except Exception as e:
            self.l.warning("类型错误：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col + 1, str(self.session.headers), self.s.res_result())

    def save_json(self, key, value):
        """保存上个接口的返回值"""
        try:
            self.param[key] = self.jsonres[value]
        except Exception as e:
            self.l.warning("操作失败：%s" % e)
            try:
                self.param[key] = self.extractor(value)
            except Exception as e:
                self.l.warning("操作失败：%s" % e)
                self.param[key] = value
        if self.param != {} or self.param is not None:
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
        else:
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
        try:
            self.w.sheet_write(self.w.row, self.w.col + 1, self.param, self.s.res_result())
        except Exception as e:
            self.l.warning("类型错误：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col + 1, json.dumps(self.param), self.s.res_result())

    def set_url(self, host, url=None):
        """设置主机地址"""
        if host.startswith('h') or host.startswith('H'):
            if url == '' or url is None:
                self.url = host
            else:
                self.url = host + url
        else:
            if url == '' or url is None:
                self.url = 'http://' + host
            else:
                self.url = 'http://' + host + url
        self.l.info("主机地址：%s" % self.url)
        self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
        self.w.sheet_write(self.w.row, self.w.col + 1, self.url, self.s.res_result())
        return self.url

    @staticmethod
    def __write_content(s):
        """控制写入的长度"""
        length = len(s)
        if length > 32767:
            return s[:32767]
        else:
            pass
