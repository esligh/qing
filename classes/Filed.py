# coding=utf-8


class Field(object):
    def __init__(self, seq=-1, name='', key=''):
        self.seq = seq
        self.name = name
        self.key = key

    def __str__(self):
        return '[Field: seq=%s,name=%s,key=%s]' % (self.seq, self.name, self.key)

