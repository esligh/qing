# coding=utf-8
import sys
import xlwings as xw
reload(sys)
sys.setdefaultencoding('utf8')


app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False


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
            print 'idx = %s' % idx
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

name = 'F:\\workspace\\python\\qing\\doc\\1.xlsx'
wb = app.books.open(name)

try:
    st = wb.sheets[3]
    result = detect_left_bound(st)
    start_c = result[0]
    start_idx = result[1]
    print 'startc_c=%s , start_idx = %s' % (start_c, start_idx)
    values = st.range('C2').expand('down')
    print len(values)
finally:
    wb.close()

app.quit()



