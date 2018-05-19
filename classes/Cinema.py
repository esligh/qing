# coding=utf-8
from common import constants


# 资源项 即excel中一行资源的描述
# 注：不能修改该类字段的名称
class Cinema(object):
    def __init__(self, seq=-1, province='', district='', city='', city_level='',
                 name='', cid='', address='', unique_code='', res_attr='',
                 ascription='', hall_count=0, seats_count=0, rank=0, providor=0,
                 box_office=0, show_count=0, viewer_count=0, open_time=''):
        self.seq = seq                      # 序号
        self.province = province            # 所属省份
        self.city = city                    # 市
        self.district = district            # 地区
        self.city_level = city_level        # 城市等级
        self.name = name                    # 影院名称
        self.cid = cid                      # 影院id
        self.address = address              # 影院地址
        self.unique_code = unique_code      # 专资编码
        self.res_attr = res_attr            # 资源属性
        self.ascription = ascription        # 院线归属
        self.hall_count = hall_count        # 院线厅数
        self.seats_count = seats_count      # 座位数
        self.rank = rank                    # 排名
        self.providor = providor            # 表示资源提供者 即谁的资源
        self.box_office = box_office        # 票房
        self.show_count = show_count        # 场次
        self.viewer_count = viewer_count    # 人次
        self.open_time = open_time          # 开业时间

    def __iter__(self):
        return self

    def __str__(self):
        return '[seq=%s,province=%s,city=%s,district=%s,city_level=%s,cid=%s,name=%s,address=%s,\n' \
               'unique_code=%s ,res_attr=%s,ascription=%s,hall_count=%s,seats_count=%s,providor=%s,\n' \
               'box_office=%s,viewer_count=%s,show_count=%s,open_time=%s]' \
               % (self.seq, self.province, self.city, self.district, self.city_level, self.cid, self.name,
                  self.address, self.unique_code, self.res_attr, self.ascription, self.hall_count, self.seats_count,
                  self.providor, self.box_office, self.viewer_count, self.show_count, self.open_time)

    def compute(self):
        self.box_office = (int(self.hall_count)+1)*100      # 票房数=(厅数+1)*100
        self.viewer_count = (int(self.hall_count) * 365)*6*40  # 人次 = 厅数*365*6*40


if __name__ == '__main__':

    cinema = Cinema()
    # s = "seq"
    # cinema.__dict__[s] = 1
    # d = "province"
    # cinema.__dict__[d] = '河北省'
    # cinema.city = '北京'
    # cinema.seats_count = '1234'
    # cinema.hall_count = '5'
    # cinema.unique_code = '23423223'
    # cinema.address = '北京市人民路'
    # cinema.ascription = '上海联合影院'
    # cinema.city_level = '一线'
    # cinema.res_attr = '独家'
    # cinema.name = '中影国际'
    # cinema.cid = '21'
    # cinema.providor = Providor.JINGMAO

    print cinema

    print '==%s' % constants.PROVIDOR.get(cinema.providor)
