from req_module import HTTP_SJSC
from excel_module import EXCEL_TEST
from log_module import LOG_TEST
from web_module import WEB_TEST
import os


def run_case(mode, line):
    if len(line[0]) > 0 or len(line[1]) > 0:
        pass
    else:
        if line[3] == '' or line[3] is None:
            pass
        else:
            lf.info("测试用例：%s" % line)
            func = getattr(mode, line[3])
            lf.info("关键字：%s" % line[3])
            p = [i for i in line if i != '']
            if len(p) == 3:
                func(line[4])
            elif len(p) == 4:
                func(line[4], line[5])
            elif len(p) == 5:
                func(line[4], line[5], line[6])
            else:
                func()


if __name__ == '__main__':
    src = 'C:\\Users\\Administrator\\Documents\\H5-Test.xls'
    dst = 'C:\\Users\\Public\\Documents\\H5-Test.xls'
    r_data = EXCEL_TEST.Reader()
    w_data = EXCEL_TEST.Writer()
    f_data = EXCEL_TEST.Format_W()
    t_data = r_data.open_read(src)
    w_data.open_copy(src, dst)
    # 生成日志
    l = LOG_TEST.LOG()
    # os.path.abspath('..') + '\\test_log'
    # p = os.path.join(r'C:\Users\Public\Documents\test_log', '\\test_log')
    l.log_dir(r'C:\Users\Public\Documents\test_log', r_data.row_col(0, 1))
    lf = l.log_msg(r_data.row_col(0, 1))
    # 测试执行
    if r_data.row_col(0, 1) == 'HTTP' or r_data.row_col(1, 1) == 'HTTP':
        mode = HTTP_SJSC.HTTP(w_data, f_data, lf)
    else:
        mode = WEB_TEST.UI(w_data, f_data, lf)
    for i in range(t_data):
        num = r_data.sheet_index(i)
        w_data.set_sheet(i)
        for j in range(num):
            print(r_data.sheet_value(j))
            w_data.row = j
            w_data.col = 7
            runner = run_case(mode, r_data.sheet_value(j))
            if runner == "" or runner is None:
                pass
            else:
                print(runner)
    w_data.save_close()
