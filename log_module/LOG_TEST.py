import logging, os, time


class LOG:
    def __init__(self):
        self.path_name = ''
        self.name = None
        self.formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s', '%Y-%m-%d %H:%M:%S')

    def log_dir(self, d_path, d_name):
        """初始化日志文件保存路径"""
        if not os.path.exists(d_path):
            os.mkdir(d_path)
        else:
            os.path.isfile(d_path)
        self.name = d_name + time.strftime('-%Y%m%d_%H%M%S', time.localtime()) + '.log'
        self.path_name = os.path.join(d_path, self.name)
        if os.path.exists(self.path_name):
            self.path_name = os.path.isfile(self.path_name)
        else:
            pass
        return self.path_name

    def log_msg(self, l_name):
        """生成日志文件并进行日志内容输出"""
        logger = logging.getLogger(l_name)
        logger.setLevel(logging.DEBUG)
        '''日志输出到控制台'''
        ch = logging.StreamHandler()
        ch.setLevel(logging.DEBUG)
        '''写入所有日志'''
        fh = logging.FileHandler(filename=self.path_name, encoding='utf-8', )
        fh.setLevel(logging.DEBUG)
        '''写入错误日志'''
        eh = logging.FileHandler(filename=self.path_name, encoding='utf-8', )
        eh.setLevel(logging.ERROR)
        ch.setFormatter(self.formatter)
        fh.setFormatter(self.formatter)
        eh.setFormatter(self.formatter)
        logger.addHandler(ch)
        logger.addHandler(fh)
        logger.addHandler(eh)
        return logger


if __name__ == '__main__':
    cs = LOG()
    p = os.path.abspath('..') + '\\testlog'
