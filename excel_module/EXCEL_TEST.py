import os, sys, time, xlrd, xlwt
from xlutils.copy import copy


class Reader:
    def __init__(self):
        self.workbook = self.sheet = self.index = None
        self.rows = 0

    def open_read(self, srcfile):
        """打开用例源文件"""
        if os.path.exists(srcfile):
            xlrd.Book.encoding = 'utf-8'
            self.workbook = xlrd.open_workbook(filename=srcfile)
            self.sheet = self.workbook.nsheets
            return self.sheet
        else:
            print('%s not exist,testing will over' % srcfile)
            sys.exit()

    def sheet_index(self, i_no):
        """定位用例源文件的Sheet页面"""
        self.index = self.workbook.sheet_by_index(i_no)
        self.rows = self.index.nrows
        return self.rows

    def sheet_value(self, r_no):
        """获取整行数据"""
        return self.index.row_values(r_no)

    def row_col(self, r, c):
        """获取指定行指定列数据"""
        sheet = self.workbook.sheet_by_index(0)
        return sheet.cell(r, c).value


class Writer:
    def __init__(self):
        self.workbook = self.wb = self.sheet = self.df = self.tables = None
        self.row = self.col = 0

    def open_copy(self, srcfile, dstfile):
        """打开用例源文件并复制源文件的数据及格式"""
        if os.path.exists(srcfile):
            xlrd.Book.encoding = 'utf-8'
            self.df = dstfile
            self.workbook = xlrd.open_workbook(filename=srcfile, formatting_info=True)
            self.wb = copy(self.workbook)
            return self.wb
        else:
            print('%s not exist,testing will over' % srcfile)
            sys.exit()

    def set_sheet(self, s_no):
        """控制Sheet页面切换"""
        self.tables = self.wb.get_sheet(s_no)
        return self.tables

    def sheet_write(self, row, col, value, style):
        """测试数据回写"""
        return self.tables.write(row, col, value, style)

    def save_close(self):
        """保存测试结果文件"""
        (filepath, tempfilename) = os.path.split(self.df)
        (filename, extension) = os.path.splitext(tempfilename)
        filename = os.path.join(filepath, filename)
        self.df = filename + time.strftime("-%Y%m%d_%H%M%S", time.localtime()) + extension
        return self.wb.save(self.df)


class Format_W:
    def __init__(self):
        """控制表格边框，属于公共区块"""
        self.borders = xlwt.Borders()
        self.borders.left = 1
        self.borders.left = xlwt.Borders.THIN
        self.borders.right = 1
        self.borders.top = 1
        self.borders.bottom = 1

    def pass_result(self):
        """Pass结果的写入格式"""
        style = xlwt.XFStyle()
        a1 = xlwt.Alignment()
        a1.horz = xlwt.Alignment.HORZ_CENTER
        a1.vert = xlwt.Alignment.VERT_CENTER
        style.alignment = a1
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 3
        style.pattern = pattern
        style.borders = self.borders
        return style

    def fail_result(self):
        """Fail结果的写入格式"""
        style = xlwt.XFStyle()
        a1 = xlwt.Alignment()
        a1.horz = xlwt.Alignment.HORZ_CENTER
        a1.vert = xlwt.Alignment.VERT_CENTER
        style.alignment = a1
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 2
        style.pattern = pattern
        style.borders = self.borders
        return style

    def res_result(self):
        """实际结果的写入格式"""
        style = xlwt.XFStyle()
        a1 = xlwt.Alignment()
        a1.horz = xlwt.Alignment.HORZ_LEFT
        a1.vert = xlwt.Alignment.VERT_CENTER
        style.alignment = a1
        style.borders = self.borders
        return style


def file_lists(path):
    return os.listdir(path)


if __name__ == '__main__':
    r = Reader()
    file = r.open_read('C:\\Users\\melon\\Documents\\yxzngkpt.xls')
    print(file)
