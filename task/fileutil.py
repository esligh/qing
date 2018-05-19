# coding=utf-8
import os


def docfilter(name):
    '''
    这个方法主要针对目录下的文件进行过滤，满足要求的返回True,否则返回False
    :param name: 这个是excel的文件名称
    :return: True 合格的文件  False不合格的文件
    '''
    if u'晶茂' in name:
        return True
    else:
        return False


def lookup(rootdir):
    '''
    这个方法遍历文档的目录，取到excel的文件列表，并按照过滤规则过滤文件，
    返回过滤后的文件列表
    :param rootdir:
    :return:
    '''
    result = []
    files = os.listdir(rootdir)
    for i in files:
        path = os.path.join(rootdir, i)
        # print path.decode('gbk')
        result.append(path)
        # name = path.decode('gbk')
        # if docfilter(name):
        #     result.append(name)
    return result


def detect_left_bound(st):
    idx = 0
    start_c = 'A'
    for i in range(1, 10):
        for j in range(1, 10):
            rv = st.range('%s%s' % (start_c, j)).expand('right')
            if len(rv) > 5:
                idx = j
                break
        if idx > 0:
            start_c = 'A'
            for k in range(1, 10):
                dv = st.range('%s%s' % (start_c, idx+2)).expand('down')
                if len(dv) > 5:
                    print 'result:idx = %d , start_c = %s' % (idx, start_c)
                    break
                else:
                    start_c = chr(ord(start_c) + 1)
            break
        else:
            start_c = chr(ord(start_c) + 1)

    return start_c, idx