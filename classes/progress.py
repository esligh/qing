#!/usr/local/lib
# -*- coding: UTF-8 -*-

import sys, time


def show_progress(percent):
    k = percent
    str = '>' * (k // 2) + ' ' * ((100 - k - 1) // 2)
    sys.stdout.write('\r' + str + '[%s%%]' % (percent + 1))
    sys.stdout.flush()
