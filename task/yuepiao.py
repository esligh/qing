# coding=utf-8
# 月票房 excel的解析
import sys
from classes.Filed import Field
import fileutil as fu
import pickle
import os
from classes import progress
from classes.Providor import Providor

reload(sys)
sys.setdefaultencoding('utf8')


def save_result(app, name, result, cinemas):
    index_code = result[0]
    offset = result[1]
    no_index = 9540;
    print '\n\n===========================result==============================\n\n'
    try:
        wb = app.books.open(name)
        st = wb.sheets[0]
        st.range('U%s' % (offset-1)).value = u'晶茂'
        st.range('V%s' % (offset - 1)).value = u'乐幕'
        st.range('W%s' % (offset - 1)).value = u'分众晶视'
        st.range('X%s' % (offset - 1)).value = u'万达'
        total_count = jm_count = lm_count = fz_count = wd_count = 0
        count = len(cinemas)
        for c in cinemas:
            # print c
            total_count += 1
            b = False
            for i in range(0, len(index_code)):
                if c.unique_code == index_code[i]:
                    index = (i + offset)
                    b = True
                    # print 'key =%s ,find index== %d ' % (key, index)
                    if c.providor == Providor.JINGMAO and u'独家' in c.res_attr:
                        key = chr(ord('U'))
                        st.range('%s%s' % (key, index)).value = u'晶茂'
                        jm_count += 1
                    elif c.providor == Providor.LEMU:
                        key = chr(ord('U')+1)
                        st.range('%s%s' % (key, index)).value = u'乐幕'
                        lm_count += 1
                    elif c.providor == Providor.FENZHONG:
                        key = chr(ord('U')+2)
                        st.range('%s%s' % (key, index)).value = u'分众晶视'
                        fz_count += 1
                    elif c.providor == Providor.WANDA:
                        key = chr(ord('U')+3)
                        st.range('%s%s' % (key, index)).value = u'万达'
                        wd_count += 1
                    break
            if not b:
                if c.providor == Providor.JINGMAO:
                    attr = u'晶茂'
                elif c.providor == Providor.LEMU:
                    attr = u'乐幕'
                elif c.providor == Providor.FENZHONG:
                    attr = u'分众晶视'
                elif c.providor == Providor.WANDA:
                    attr = u'万达'
                st.range('A%s' % no_index).value = [no_index, c.name, c.hall_count, c.seats_count, c.box_office,
                                                    c.show_count, c.viewer_count, '', '', '', '', '', '',
                                                    c.district, c.province, c.ascription, c.open_time, c.unique_code,
                                                    '', '', attr]
                no_index += 1
            percent = int((total_count * 1.0 / count) * 100)
            progress.show_progress(percent)
        print '\n'
        print 'total_count =%d, jm_count = %d , lm_count = %d, fz_count = %d, wd_count = %d' % \
              (total_count, jm_count, lm_count, fz_count, wd_count)
        wb.save()
    finally:
        wb.close()


def parse(app, name):
    cinemas = []
    wb = app.books.open(name)
    try:
        idx = find_sheet_index(wb.sheets)
        print 'sheed index : '+str(idx)
        st = wb.sheets[idx]
        result = fu.detect_left_bound(st)
        start_c = result[0]
        start_idx = result[1]
        rows = st.range('%s%s' % (start_c, start_idx + 1)).expand('down')
        row_count = len(rows)
        print 'row count=%d' % row_count
        if start_idx > 0:
            titles = get_titles(st, start_c, start_idx)
            cinemas = read_rows_data(st, titles, start_idx+1, row_count)
    finally:
        wb.close()
    return cinemas, start_idx+1


# 获取excel有效的titles
def get_titles(st, start_c, start_idx):
    result = []
    titles = st.range('%s%s' % (start_c, start_idx)).expand('right')
    if len(titles) > 0:
        for t in range(0, len(titles)):
            title = titles[t].value.replace(' ', '')
            c = ord(start_c)+t
            if u'序号' == title :
                field = Field(t, 'seq', chr(c))
            elif u'排名' == title :
                field = Field(t, 'rank', chr(c))
            elif u'影院编号' == title or u'影院ID' == title:
                field = Field(t, 'cid', chr(c))
            elif u'省份' == title :
                field = Field(t, 'province', chr(c))
            elif u'城市' == title :
                field = Field(t, 'city', chr(c))
            elif u'区县' == title or u'地区' == title:
                field = Field(t, 'district', chr(c))
            elif u'票房名称' == title or u'影院名称' == title:
                field = Field(t, 'name', chr(c))
            elif u'资源属性' == title :
                field = Field(t, 'res_attr', chr(c))
            elif u'院线归属' == title:
                field = Field(t, 'ascription', chr(c))
            elif u'地址' in title :
                field = Field(t, 'address', chr(c))
            elif u'专资编码' == title or u'专资代码' == title:
                field = Field(t, 'unique_code', chr(c))
            elif u'城市级别' == title :
                field = Field(t, 'city_level', chr(c))
            elif u'座位数' in title:
                field = Field(t, 'seats_count', chr(c))
            elif u'厅数' in title:
                field = Field(t, 'hall_count', chr(c))
            else:
                field = Field(t, '', '')
            result.append(field)

    else:
        print 'not found valide titles '
    return result


# 查找要打开的sheet索引
def find_sheet_index(sheets):
    idx = 0
    for sheet in sheets:
        if u'月票房' in sheet.name:
            break
        else:
            idx += 1
    return idx


def read_rows_data(st, titles, start, count):
    if os.path.exists(os.getcwd()+'\\temp.pk'):
        with open(os.getcwd()+'\\temp.pk', 'rb') as f:
            result = pickle.load(f)
        progress.show_progress(99)
        print '\n'
    else:
        result = []
        c = ''
        for title in titles:
            if title.name == 'unique_code':
                c = title.key
        for i in range(0, count):
            idx = start+i
            row_data = st.range('%s%s' % (c, idx))
            row_data.value = row_data.value.strip()
            if row_data.value is None or row_data.value == '':
                result.append(-1)
            else:
                result.append(int(row_data.value))
            # show progress
            percent = int((i * 1.0 / count) * 100)
            progress.show_progress(percent)
        print '\n'
        with open(os.getcwd()+'\\temp.pk', 'wb') as f:
            pickle.dump(result, f)
    return result
