import requests, json, jsonpath, traceback, string, random, time
from urllib.parse import urlencode


class HTTP:
    def __init__(self, w, s, log):
        self.session = requests.session()
        self.url = self.result = self.response = None
        self.param = self.dict = {}
        self.code = 0
        self.js = jsonpath
        self.w = w
        self.s = s
        self.log = log

    def add_header(self, key, value):
        """添加请求头"""
        try:
            self.session.headers[key] = self.__get_values(value)
        except Exception as w:
            self.log.warning("没有变量：%s" % w)
            self.session.headers[key] = value
        if self.session.headers.get(key) is not None:
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
        else:
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
        try:
            self.w.sheet_write(self.w.row, self.w.col + 1, self.session.headers, self.s.res_result())
        except Exception as w:
            self.log.warning("类型错误：%s" % w)
            self.w.sheet_write(self.w.row, self.w.col + 1, str(self.session.headers), self.s.res_result())

    def assert_equals(self, key, value):
        """测试结果断言"""
        param = json.loads(self.response)
        if param.get(key) is None and param.get(value) is None:
            try:
                if isinstance(self.extractor(key), str) and 200 >= self.code < 400:
                    if self.extractor(key) == value:
                        self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
                    else:
                        self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.pass_result())
                elif isinstance(self.extractor(key), int) and 200 >= self.code < 400:
                    if self.extractor(key) == int(value):
                        self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
                    else:
                        self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.pass_result())
            finally:
                self.w.sheet_write(self.w.row, self.w.col + 1, self.extractor(key), self.s.res_result())
        else:
            try:
                if isinstance(param.get(key), str) and 200 >= self.code < 400:
                    if param.get(key) == value:
                        self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
                    else:
                        self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.pass_result())
                elif isinstance(param.get(key), int) and 200 >= self.code < 400:
                    if param.get(key) == int(value):
                        self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
                    else:
                        self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.pass_result())
            finally:
                self.w.sheet_write(self.w.row, self.w.col + 1, param[key], self.s.res_result())

    def assert_json(self, key, value):
        param = json.loads(self.response)
        if param.get(key) is None and param.get(value) is not None:
            try:
                if self.extractor(key) == self.extractor(value) and 200 >= self.code < 400:
                    self.log.info("预期结果：%s" % self.extractor(key))
                    self.log.info("实际结果：%s" % self.extractor(value))
                    self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
                else:
                    self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.pass_result())
            finally:
                self.w.sheet_write(self.w.row, self.w.col + 1, self.extractor(key), self.s.res_result())
        else:
            try:
                if param.get(key) == param.get(value) and 200 >= self.code < 400:
                    self.log.info("预期结果：%s" % param.get(key))
                    self.log.info("实际结果：%s" % param.get(value))
                    self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
                else:
                    self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.pass_result())
            finally:
                self.w.sheet_write(self.w.row, self.w.col + 1, param[key], self.s.res_result())

    def delete(self, path, payload):
        """Delete请求总入口"""
        if '{' in path or path.endswith('}'):
            tmp_path = self.__url_var(path)
            tmp = self.__get_values(tmp_path)
            path_list = path.split('/')[1:]
            index = path_list.index(tmp_path)
            path = self.__url_join(tmp, index, path_list)
        try:
            self.result = self.session.delete(self.url + path + payload, timeout=16)
            self.response = self.result.content.decode('utf-8')
            self.dict = json.loads(self.response)
            self.code = self.result.status_code
            self.log.info("Delete请求成功：%s" % self.response)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.__write_content(self.response), self.s.res_result())
            return self.response
        except Exception as e:
            self.log.error("Delete请求失败：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
        finally:
            self.log.info("Delete请求头：%s" % self.session.headers)
            self.log.info("Delete请求参数：%s" % self.url + path + payload)

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

    def file_post(self, file, path, payload=None, ):
        self.__test_params(payload)
        file = {'file': open(file, 'rb')}
        try:
            self.result = self.session.post(self.url + path, json=payload, files=file, timeout=16)
        except Exception as e:
            self.log.error("Post请求失败：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())

    def get(self, path, payload=None, ):
        """Get请求总入口"""
        # if '{' in path or path.endswith('}'):
        #     tmp_path = self.__url_var(path)
        #     tmp = self.__get_values(tmp_path)
        #     path_list = path.split('/')[1:]
        #     index = path_list.index(tmp_path)
        #     path = self.__url_join(tmp, index, path_list)
        if payload == "" or payload is None:
            self.__get_url(path)
        else:
            self.__get_payload(path, self.__test_params(payload))

    def __get_payload(self, path, payload):
        """带入参的Get请求"""
        try:
            self.result = self.session.get(self.url + path + '?' + urlencode(payload), timeout=8)
            self.response = self.result.content.decode('utf-8')
            self.dict = json.loads(self.response)
            self.code = self.result.status_code
            self.log.info("Get请求成功：%s" % self.response)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.__write_content(self.response), self.s.res_result())
            return self.response
        except Exception as e:
            self.log.error("Get请求失败：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
        finally:
            self.log.info("Get请求头：%s" % self.session.headers)
            try:
                self.log.info("Get请求参数：%s" % self.url + path + '?' + payload)
            except Exception as w:
                self.log.warning("类型错误：%s" % w)
                self.log.info("Get请求参数：%s" % self.url + path + '?' + urlencode(payload))

    def __get_url(self, path):
        """无入参的Get请求"""
        try:
            self.result = self.session.get(self.url + path, timeout=8)
            self.response = self.result.content.decode('utf-8')
            self.dict = json.loads(self.response)
            self.code = self.result.status_code
            self.log.info("Get请求成功：%s" % self.response)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.__write_content(self.response), self.s.res_result())
            return self.response
        except Exception as e:
            self.log.error("Get请求失败：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
        finally:
            self.log.info("Get请求头：%s" % self.session.headers)
            self.log.info("Get请求地址：%s" % self.url + path)

    def __get_values(self, s):
        """获取变量的值"""
        if s == '' or s is None:
            return ''
        else:
            for key in self.param.keys():
                s = s.replace('{' + key + '}', self.param[key])
            return s

    def post(self, path, payload=None):
        """Post请求总入口"""
        # if '{' in path:
        #     tmp_path = self.__url_var(path)
        #     tmp = self.__get_values(tmp_path)
        #     path_list = path.split('/')[1:]
        #     index = path_list.index(tmp_path)
        #     path = self.__url_join(tmp, index, path_list)
        if payload == '' or payload is None:
            self.__post_url(path)
        else:
            # self.__post_payload(path, self.__req_params(self, payload))
            self.__post_payload(path, self.__test_params(payload))

    def __post_payload(self, path, payload):
        """处理有入参的Post请求"""
        try:
            if self.session.headers.get('Content-Type') == 'application/x-www-form-urlencoded':
                self.result = self.session.post(self.url + path, data=urlencode(payload), timeout=16)
            elif self.session.headers.get('Content-Type') == 'application/json':
                if not isinstance(payload, dict):
                    self.result = self.session.post(self.url + path, data=payload, timeout=16)
                self.result = self.session.post(self.url + path, json=payload, timeout=16)
            else:
                self.result = self.session.post(self.url + path + '?' + urlencode(payload), timeout=16)
            self.response = self.result.content.decode('utf-8')
            self.dict = json.loads(self.response)
            self.code = self.result.status_code
            self.log.info("Post请求成功：%s" % self.response)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.__write_content(self.response), self.s.res_result())
            # self.w.sheet_write(self.w.row, self.w.col + 2, str(payload), self.s.res_result())
            return self.response
        except Exception as e:
            self.log.error("Post请求失败：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
        finally:
            self.log.info("Post请求头：%s" % self.session.headers)
            self.log.info("Post请求地址：%s" % self.url + path)
            self.log.info("Post请求入参：%s" % payload)

    def __post_url(self, path):
        """处理无入参的Post请求"""
        try:
            self.result = self.session.post(self.url + path, timeout=16)
            self.response = self.result.content.decode('utf-8')
            self.dict = json.loads(self.response)
            self.code = self.result.status_code
            self.log.info("Post请求成功：%s" % self.response)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.__write_content(self.response), self.s.res_result())
            return self.response
        except Exception as e:
            self.log.error("Post请求失败：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
        finally:
            self.log.info("Post请求头：%s" % self.session.headers)
            self.log.info("Post请求地址：%s" % self.url + path)

    def put(self, path, payload):
        """Put请求总入口"""
        if '{' in path:
            tmp_path = self.__url_var(path)
            tmp = self.__get_values(tmp_path)
            path_list = path.split('/')[1:]
            index = path_list.index(tmp_path)
            path = self.__url_join(tmp, index, path_list)
        try:
            payload = self.__test_params(payload)
            self.result = self.session.put(self.url + path, data=json.dumps(payload), timeout=16)
            self.response = self.result.content.decode('utf-8')
            self.dict = json.loads(self.response)
            self.code = self.result.status_code
            self.log.info("Put请求成功：%s" % self.response)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.__write_content(self.response), self.s.res_result())
            return self.response
        except Exception as e:
            self.log.error("Put请求失败：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
        finally:
            self.log.info("Put请求头：%s" % self.session.headers)
            self.log.info("Put请求地址：%s" % self.url + path)
            try:
                self.log.info("Put请求入参：%s" % payload)
            except Exception as w:
                self.log.warning("类型错误：%s" % w)
                self.log.info("Put请求入参：%s" % str(payload))

    def remove_header(self, key):
        """删除请求头"""
        try:
            self.session.headers.pop(key)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
        except Exception as e:
            self.log.error("删除出错：%s" % e)
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
        try:
            self.w.sheet_write(self.w.row, self.w.col + 1, self.session.headers, self.s.res_result())
        except Exception as w:
            self.log.warning("类型错误：%s" % w)
            self.w.sheet_write(self.w.row, self.w.col + 1, str(self.session.headers), self.s.res_result())

    def __test_params(self, param):
        if param.startswith('{') and param.endswith('}'):
            if '{' or '}' not in param[2:-2]:
                return json.loads(param)
            else:
                p = json.loads(param)
                for k, v in p.items():
                    if v.startswith('{') and isinstance(v, str):
                        p[k] = self.__get_values(v)
                    else:
                        continue
                return p
        else:
            params = {}
            for pp in param.split('&'):
                if '{' or '}' in pp[pp.find('=') + 1:]:
                    params[pp[0:pp.find('=')]] = self.__get_values(pp[pp.find('=') + 1:])
                else:
                    params[pp[0:pp.find('=')]] = pp[pp.find('=') + 1:]
            return params

    def save_json(self, key, value):
        """保存上个接口的返回值"""
        try:
            self.param[key] = self.dict[value]
        except Exception as e:
            self.log.warning("操作失败：%s" % e)
            if '.' in value:
                self.param[key] = self.extractor(value)
            else:
                if value == '随机':
                    var = random.sample(string.ascii_letters + string.digits, random.randint(4, 8))
                    self.param[key] = value.replace(value, ''.join(var))
                elif value == '时间':
                    var = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
                    self.param[key] = value.replace(value, var)
                else:
                    self.param[key] = value
        if self.param.get(key) is not None:
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
        else:
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
        try:
            self.w.sheet_write(self.w.row, self.w.col + 1, self.param, self.s.res_result())
        except Exception as w:
            self.log.warning("类型错误：%s" % w)
            self.w.sheet_write(self.w.row, self.w.col + 1, str(self.param), self.s.res_result())

    def set_url(self, host, url=None):
        """设置主机地址"""
        try:
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
            self.log.info("主机地址：%s" % self.url)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.url, self.s.res_result())
        except Exception as e:
            self.log.error("没有地址：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
        return self.url

    @staticmethod
    def __url_join(text, index, path_list):
        """替换地址中的变量，并重新完成地址的拼接"""
        tmp_path = []
        path_list.pop(index)
        path_list.insert(index, text)
        for i in path_list:
            p = '/' + i
            tmp_path.append(p)
        return ''.join(tmp_path)

    @staticmethod
    def __url_var(path):
        """查找URL中的变量"""
        li = path.split('/')
        var = None
        for i in li:
            if i.startswith('{'):
                var = i
            else:
                pass
        return var

    @staticmethod
    def __write_content(s):
        """控制写入的长度"""
        if len(s) <= 32767:
            return s
        else:
            return s[:32767]


class KeyWords:
    @staticmethod
    def __find_values(s):
        """查询字典中的变量"""
        for value in s.values():
            if isinstance(value, str) and value.startswith('{'):
                for key in s.keys():
                    s[key] = s[key]
            else:
                continue
        else:
            return s

    @staticmethod
    def __get_params(param):
        """Get入参处理"""
        '''
        try:
            if isinstance(param[key], str):
                if 200 >= self.code < 400 and param[key] == value:
                    self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
                else:
                    self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            elif isinstance(param[key], int):
                if 200 >= self.code < 400 and param[key] == int(value):
                    self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
                else:
                    self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            else:
                if 200 >= self.code < 400 and param[key] == param[value]:
                    self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
                else:
                    self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, param[key], self.s.res_result())
        except Exception as w:
            self.log.warning("匹配失败：%s" % w)
            if isinstance(self.extractor(key), str):
                if 200 >= self.code < 400 and self.extractor(key) == value:
                    self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
                else:
                    self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            elif isinstance(self.param[key], int):
                if 200 >= self.code < 400 and self.extractor(key) == int(value):
                    self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
                else:
                    self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.extractor(key), self.s.res_result())
        '''
        if isinstance(param, dict):
            return urlencode(param)
        else:
            return str(param)

    def __url_params(self, param):
        params = {}
        if len(param) >= 1:
            p = param.split('&')
            for pp in p:
                ppp = pp.split('=')
                if ppp[1].startswith('{') or ppp[1].endswith('}'):
                    params[ppp[0]] = self.__find_values(ppp[1])
                else:
                    params[ppp[0]] = ppp[1]
            return params

    def __dict_params(self, param):
        if '{' in param[1:-2] and '}' in param[1:-2]:
            p = json.loads(param)
            for k, v in p.items():
                if not isinstance(v, str):
                    continue
                elif v.startswith('{'):
                    p[k] = self.__find_values(v)
                else:
                    continue
            return p
        else:
            return json.loads(param)

    @staticmethod
    def __req_params(self, param):
        """Post入参处理"""
        # if '{' or '}' in quote(pp[pp.find('=') + 1:]):
        #     params[quote(pp[0:pp.find('=')])] = self.__get_values(quote(pp[pp.find('=') + 1:]))
        # else:
        #     params[quote(pp[0:pp.find('=')])] = quote(pp[pp.find('=') + 1:])
        if param.startswith('{') and param.endswith('}'):
            p = json.loads(param)
            for k, v in p.items():
                if not isinstance(v, str):
                    continue
                elif v.startswith('{'):
                    p[k] = self.__get_values(v)
                else:
                    continue
            else:
                return p
        else:
            params = {}
            if len(param) >= 1:
                p = param.split('&')
                for pp in p:
                    ppp = pp.split('=')
                    if ppp[1].startswith('{') or ppp[1].endswith('}'):
                        params[ppp[0]] = self.__get_values(ppp[1])
                    else:
                        params[ppp[0]] = ppp[1]
                return params
