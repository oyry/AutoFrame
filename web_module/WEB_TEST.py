import os, sys, time, random, traceback, datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.select import Select


class Extend:
    def __init__(self, w, s, log):
        self.dr = self.ele = None
        self.op = Options()
        self.url = self.u = ''
        self.w = w
        self.s = s
        self.log = log
        self.to = random.randint(10, 20)  # timeout简写
        self.pf = random.uniform(0.6, 1.6)  # poll_frequency简写

    def browser_option(self, text=None):
        time.sleep(random.randint(1, 3))
        try:
            if text == '' or text is None:
                self.dr.close()
                self.dr.quit()
            elif text == 'tab' or 'Tab' or 'TAB':
                self.dr.close()
            elif text == 'back' or 'Back' or 'BACK':
                self.dr.back()
            elif text == 'refresh' or 'Refresh' or 'REFRESH':
                self.dr.refresh()
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "浏览器操作成功", self.s.res_result())
        except Exception as e:
            self.log.error("浏览器操作失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            try:
                self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
            except Exception as w:
                self.log.warning("类型错误：%s" % w)
                self.w.sheet_write(self.w.row, self.w.col + 1, str(traceback.format_exc()), self.s.res_result())

    def default_chrome(self, text=None):
        if text == '' or text is None:
            sys.exit()
        else:
            try:
                self.op.add_argument('--disable-infobars')
                self.dr = webdriver.Chrome(options=self.op)
                self.dr.maximize_window()
                self.dr.implicitly_wait(random.randint(10, 20))
                self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
                self.w.sheet_write(self.w.row, self.w.col + 1, "打开谷歌成功", self.s.res_result())
            except Exception as e:
                self.log.error("打开谷歌失败：{}；{}".format(e, traceback.format_exc()))
                self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
                try:
                    self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
                except Exception as w:
                    self.log.warning("类型错误：%s" % w)
                    self.w.sheet_write(self.w.row, self.w.col + 1, str(traceback.format_exc()), self.s.res_result())
            finally:
                return self.dr.get(self.url)

    def get_url(self, u):
        try:
            if u.startswith('h') or u.startswith('H'):
                self.url = u
            else:
                self.url = 'http://' + u
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.url, self.s.res_result())
            return self.url
        except Exception as e:
            self.log.error("没有地址：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            try:
                self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
            except Exception as w:
                self.log.warning("类型错误：%s" % w)
                self.w.sheet_write(self.w.row, self.w.col + 1, str(traceback.format_exc()), self.s.res_result())

    def open_chrome(self, path):
        try:
            self.op.add_argument('--disable-infobars')
            self.op.add_experimental_option('prefs', {'profile.defalut_content_setting_values': {'notifications': 2}})
            self.dr = webdriver.Chrome(executable_path=path, options=self.op)
            self.dr.maximize_window()
            self.dr.implicitly_wait(random.randint(10, 20))
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "打开谷歌浏览器成功", self.s.res_result())
        except Exception as e:
            self.log.error("打开谷歌浏览器失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            try:
                self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
            except Exception as w:
                self.log.warning("类型错误：%s" % w)
                self.w.sheet_write(self.w.row, self.w.col + 1, str(traceback.format_exc()), self.s.res_result())
        finally:
            return self.dr.get(self.url)

    def open_firefox(self, path):
        try:
            self.dr = webdriver.Firefox(executable_path=path)
            self.dr.maximize_window()
            self.dr.implicitly_wait(random.randint(16, 28))
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "打开火狐浏览成功", self.s.res_result())
        except Exception as e:
            self.log.error("打开火狐浏览器失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            try:
                self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
            except Exception as w:
                self.log.warning("类型错误：%s" % w)
                self.w.sheet_write(self.w.row, self.w.col + 1, str(traceback.format_exc()), self.s.res_result())
        finally:
            return self.dr.get(self.url)

    def sava_url(self, s, r):
        url = self.dr.current_url
        try:
            u = url.replace(s, '')
            self.u = u.split(r)[-1]
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, self.u, self.s.res_result())
            return self.u
        except Exception as e:
            self.log.error("地址截取失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            try:
                self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
            except Exception as w:
                self.log.warning("类型错误：%s" % w)
                self.w.sheet_write(self.w.row, self.w.col + 1, str(traceback.format_exc()), self.s.res_result())

    def screen_png(self, path):
        time.sleep(random.randint(2, 5))
        name = time.strftime('%Y%m%d_%H%M%S', time.localtime()) + '.png'
        if not os.path.exists(path):
            os.mkdir(path)
        else:
            os.path.isfile(path)
        try:
            self.dr.save_screenshot(os.path.join(path, name))
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            try:
                self.w.sheet_write(self.w.row, self.w.col + 1, name, self.s.res_result())
            except Exception as w:
                self.log.warning("类型错误：%s" % w)
                self.w.sheet_write(self.w.row, self.w.col + 1, str(name), self.s.res_result())
        except Exception as e:
            self.log.error("截图失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            try:
                self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
            except Exception as w:
                self.log.warning("类型错误：%s" % w)
                self.w.sheet_write(self.w.row, self.w.col + 1, str(traceback.format_exc()), self.s.res_result())


class UI(Extend):
    def class_action(self, types, ele, text=None):
        wait = WebDriverWait(self.dr, self.to, self.pf)
        try:
            el = wait.until(lambda res: self.dr.find_element_by_class_name(ele))
            if types == 'input':
                try:
                    el.clear()
                except Exception as w:
                    self.log.warning("没有数据，不需要清空：%s" % w)
                finally:
                    el.send_keys(text)
            elif types == 'click':
                el.click()
            elif types == ('key', 'Key', 'KEY', 'keys', 'Keys', 'KEYS'):
                self.__keyboard_action(el, text)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "操作成功", self.s.res_result())
        except Exception as e:
            self.log.error("操作失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())

    def click_action(self, types, ele):
        ele = self.dr.find_element(types, ele)
        try:
            ele.click()
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "点击成功", self.s.res_result())
        except Exception as e:
            self.log.error("点击失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())

    def double_click(self, types, ele):
        ele = self.dr.find_element(types, ele)
        try:
            ActionChains(self.dr).double_click(ele).perform()
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "点击成功", self.s.res_result())
        except Exception as e:
            self.log.error("点击失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())

    def id_action(self, types, ele, text=None):
        wait = WebDriverWait(self.dr, self.to, self.pf)
        try:
            el = wait.until(lambda res: self.dr.find_element_by_id(ele))
            if types == 'input':
                try:
                    el.clear()
                except Exception as w:
                    self.log.warning("没有数据，不需要清空：%s" % w)
                finally:
                    el.send_keys(text)
            elif types == 'click':
                el.click()
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "操作成功", self.s.res_result())
        except Exception as e:
            self.log.error("操作失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())

    def __keyboard_action(self, ele, text=None):
        try:
            if text == ('enter', 'Enter', 'ENTER'):
                ele.send_keys(Keys.ENTER)
            elif text == ('space', 'Space', 'SPACE'):
                ele.send_keys(Keys.SPACE)
            elif text == ('up', 'Up', 'UP'):
                ele.send_keys(Keys.PAGE_UP)
            elif text == ('down', 'Down', 'DOWN'):
                ele.send_keys(Keys.PAGE_DOWN)
        except Exception as e:
            self.log.error("键盘输入失败：{}；{}".format(e, traceback.format_exc()))

    def name_action(self, types, ele, text=None):
        wait = WebDriverWait(self.dr, self.to, self.pf)
        try:
            el = wait.until(lambda res: self.dr.find_element_by_name(ele))
            if types == 'input':
                try:
                    el.clear()
                except Exception as w:
                    self.log.warning("没有数据，不需要清空：%s" % w)
                finally:
                    el.send_keys(text)
            elif types == 'click':
                el.click()
            elif types == ('key', 'Key', 'KEY', 'keys', 'Keys', 'KEYS'):
                self.__keyboard_action(el, text)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "操作成功", self.s.res_result())
        except Exception as e:
            self.log.error("操作失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())

    def mouse_action(self, types, ele):
        try:
            action = ActionChains(self.dr)
            action.move_to_element(self.dr.find_element(types, ele))
            action.click(self.dr.find_element_by_xpath('//*[@id="el-popover-7291"]'))
            action.perform()
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "定位成功", self.s.res_result())
        except Exception as e:
            self.log.error("元素加载失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())

    def oa_input(self, types, ele, text):
        ele = self.dr.find_element(types, ele)
        try:
            if text == '日期':
                ele.send_keys(datetime.datetime.strftime(datetime.datetime.today(), '%Y-%m-%d'))
                ele.send_keys(Keys.ENTER)
            else:
                try:
                    ele.clear()
                except Exception as w:
                    self.log.warning("没有数据，不需要清空：%s" % w)
                finally:
                    ele.send_keys(text)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "输入成功", self.s.res_result())
        except Exception as e:
            self.log.error("输入失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())

    def script_call(self, js):
        # script = js.split(',', '，')
        time.sleep(random.randint(2, 5))
        try:
            self.dr.execute_script(js)
            # for s in script:
            #     self.dr.execute_script(s)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "执行成功", self.s.res_result())
        except Exception as e:
            self.log.error("执行失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            try:
                self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
            except Exception as w:
                self.log.warning("类型错误：%s" % w)
                self.w.sheet_write(self.w.row, self.w.col + 1, str(traceback.format_exc()), self.s.res_result())

    def select_action(self, ele, text):
        el = Select(self.dr.find_element_by_xpath(ele))
        try:
            el.select_by_index(int(text))
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "获取元素成功", self.s.res_result())
        except Exception as e:
            self.log.error("获取元素失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())

    def pause(self, text=None):
        try:
            if text == '' or text is None:
                time.sleep(random.randint(10, 20))
            else:
                time.sleep(int(text))
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "页面暂停成功", self.s.res_result())
        except Exception as e:
            self.log.error("页面暂停失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())

    def window_action(self, text=None):
        if text == '' or text is None:
            tab = self.dr.current_window_handle
        else:
            tab = text
        try:
            for i in self.dr.window_handles:
                if i != tab:
                    self.dr.switch_to_window(i)
                else:
                    continue
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "切换成功", self.s.res_result())
        except Exception as e:
            self.log.error("切换失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())

    def xpath_action(self, types, ele=None, text=None):
        wait = WebDriverWait(self.dr, self.to, self.pf)
        try:
            el = wait.until(lambda res: self.dr.find_element_by_xpath(ele))
            if types == 'input':
                try:
                    el.clear()
                except Exception as w:
                    self.log.warning("没有数据，不需要清空：%s" % w)
                finally:
                    el.send_keys(text)
            elif types == 'click':
                el.click()
            elif types == ('key', 'Key', 'KEY', 'keys', 'Keys', 'KEYS'):
                self.__keyboard_action(el, text)
            self.w.sheet_write(self.w.row, self.w.col, "Pass", self.s.pass_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, "操作成功", self.s.res_result())
        except Exception as e:
            self.log.error("操作失败：{}；{}".format(e, traceback.format_exc()))
            self.w.sheet_write(self.w.row, self.w.col, "Fail", self.s.fail_result())
            self.w.sheet_write(self.w.row, self.w.col + 1, traceback.format_exc(), self.s.res_result())
