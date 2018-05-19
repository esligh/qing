# coding=utf-8
import common.constants as constants
import fileutil as fu
import lemu as lm
import jingmao as jm
import fenzhong as fz
import yuepiao as yp
import wanda as wd


def parseall(app):
    # 查找到所有的excel文件
    excels = fu.lookup(constants.DOCDIR)
    cinema = []
    result = []
    dest_file = ''
    for i in range(0, len(excels)):
        name = excels[i].decode('gbk')
        print u'正在解析: '+name
        if u'晶茂' in name:
            a = jm.parse(app, name)
            cinema.extend(a)
        elif u'乐幕' in name:
            a = lm.parse(app, name)
            cinema.extend(a)
        elif u'分众' in name:
            a = fz.parse(app, name)
            cinema.extend(a)
        elif u'万达' in name:
            a = wd.parse(app, name)
            cinema.extend(a)
        elif u'月票房' in name:
            result = yp.parse(app, name)
            dest_file = name
    yp.save_result(app, dest_file, result, cinema)

