# coding=utf-8
# 晶茂 excel的解析
from classes.Filed import Field
from classes.Cinema import Cinema
from classes.Providor import Providor
import fileutil as fu
from classes import progress
import pickle
import os

import sys
reload(sys)
sys.setdefaultencoding('utf8')


def parse(app, name):
    cinemas = []
    wb = app.books.open(name)
    try:
        idx = find_sheet_index(wb.sheets)
        print 'sheet index : '+str(idx)
        st = wb.sheets[idx]
        result = fu.detect_left_bound(st)
        start_c = result[0]
        start_idx = result[1]
        rows = st.range('%s%s' % (start_c, start_idx+1)).expand('down')
        row_count = len(rows)
        print 'row count=%d' % row_count
        if start_idx > 0:
            titles = get_titles(st, start_c, 'BI', start_idx)
            cinemas = read_rows_data(st, titles, start_c, 'BI', start_idx+1, row_count)
    finally:
        wb.close()
    return cinemas


# 获取excel有效的title field
def get_titles(st, start_c, end_c, start_idx):
    result = []
    titles = st.range('%s%s:%s%s' % (start_c, start_idx, end_c, start_idx))
    if len(titles) > 0:
        for t in range(0, len(titles)):
            title = titles[t].value
            c = ord(start_c)+t
            if title is None:
                if len(result) > 0:
                    f = result[-1]
                    field = Field(t, f.name, chr(c))
                    result.append(field)
            else:
                title = title.replace(' ', '')
                if u'序号' == title :
                    field = Field(t, 'seq', chr(c))
                elif u'影院编号' == title:
                    field = Field(t, 'cid', chr(c))
                elif u'省份' == title :
                    field = Field(t, 'province', chr(c))
                elif u'城市' == title :
                    field = Field(t, 'city', chr(c))
                elif u'城市级别' == title :
                    field = Field(t, 'city_level', chr(c))
                elif u'区县' == title :
                    field = Field(t, 'district', chr(c))
                elif u'票房名称' == title or u'影院名称' == title:
                    field = Field(t, 'name', chr(c))
                elif u'资源属性' == title :
                    field = Field(t, 'res_attr', chr(c))
                elif u'地址' in title :
                    field = Field(t, 'address', chr(c))
                elif u'专资编码' == title :
                    field = Field(t, 'unique_code', chr(c))
                elif u'总座位数' in title:
                    field = Field(t, 'seats_count', chr(c))
                elif u'总影厅数' in title:
                    field = Field(t, 'hall_count', chr(c))
                elif u'院线归属' in title:
                    field = Field(t, 'ascription', chr(c))
                else:
                    field = Field(t, '', '')
                if field.seq > -1:
                    result.append(field)

    else:
        print 'not found valide titles '
    return result


# 查找要打开的sheet索引
def find_sheet_index(sheets):
    idx = 0
    for sheet in sheets:
        if u'资源表' in sheet.name:
            break
        else:
            idx += 1
    return idx


def read_rows_data(st, titles, start_c, end_c, start_idx, count):
    if os.path.exists(os.getcwd()+'\\jingmao.pk'):
        with open(os.getcwd()+'\\jingmao.pk', 'rb') as f:
            result = pickle.load(f)
        progress.show_progress(99)
        print '\n'
    else:
        result = []
        for i in range(0, count):
            idx = start_idx+i
            row_data = st.range('%s%s:%s%s' % (start_c, idx, end_c, idx))
            if len(row_data) > 0:
                cinema = Cinema()
                for n in range(0, len(row_data)):
                    field = titles[n]
                    if len(field.name) > 0 and row_data[n].value:
                        cinema.__dict__[field.name] = row_data[n].value
            if cinema.unique_code is not None and cinema.unique_code != '':
                cinema.providor = Providor.JINGMAO
                cinema.hall_count = int(cinema.hall_count)
                cinema.seats_count = int(cinema.seats_count)
                cinema.compute()
                result.append(cinema)
            # show progress
            percent = int((i*1.0/count)*100)
            progress.show_progress(percent)
        print '\n'
        with open(os.getcwd()+'\\jingmao.pk', 'wb') as f:
            pickle.dump(result, f)
    return result
