from req_module import HTTP_TEST
from excel_module import EXCEL_TEST
from log_module import LOG_TEST
from web_module import WEB_TEST
import os


def runcase(mode, line):
    if len(line[0]) > 0 or len(line[1]) > 0:
        pass
    else:
        if line[3] == '' or line[3] is None:
            pass
        else:
            lf.info("测试用例：%s" % line)
            func = getattr(mode, line[3])
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
    file_list = ['OA报工.xls', 'OA-Test.xls', ]
    src = os.path.join('C:\\Users\\Administrator\\Documents\\', file_list[0])
    dst = os.path.join('C:\\Users\\Public\\Documents\\', file_list[0])
    r_data = EXCEL_TEST.Reader()
    w_data = EXCEL_TEST.Writer()
    f_data = EXCEL_TEST.Format_W()
    t_data = r_data.open_read(src)
    w_data.open_copy(src, dst)
    # 生成日志
    l = LOG_TEST.LOG()
    p = os.path.abspath('..') + '\\test_log'
    l.log_dir(p, r_data.row_col(0, 1))
    lf = l.log_msg(r_data.row_col(0, 1))
    # 测试执行
    if r_data.row_col(0, 1) == 'HTTP' or r_data.row_col(1, 1) == 'HTTP':
        mode = HTTP_TEST.HTTP(w_data, f_data, lf)
    else:
        mode = WEB_TEST.UI(w_data, f_data, lf)
    for i in range(t_data):
        num = r_data.sheet_index(i)
        w_data.set_sheet(i)
        for j in range(num):
            print(r_data.sheet_value(j))
            w_data.row = j
            w_data.col = 7
            runner = runcase(mode, r_data.sheet_value(j))
            if runner == "" or runner is None:
                pass
            else:
                print(runner)
    # a = w_data.df
    w_data.save_close()
