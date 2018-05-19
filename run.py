# coding=utf-8

import xlwings as xw
import task.task as pt
import  os

app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False

print 'start task ... '

pt.parseall(app)

app.quit()
