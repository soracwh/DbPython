import xlrd
import SqliteUtil
import datetime


def data(table):
    d = []
    row_num = table.nrows
    col_num = table.ncols
    for i in range(row_num):
        if i == 0:
            continue
        dev = []
        for j in range(col_num):
            dev.append(table.cell_value(i, j))
        d.append(dev)
    return d


class excel_ex:
    def __init__(self):
        self.device = []
        self.repair = []
        self.defect = []
        self.measure = []

    def get_data(self, path):
        excel = xlrd.open_workbook(path)
        dv_table = excel.sheet_by_name("设备表")
        r_table = excel.sheet_by_name("检修记录表")
        df_table = excel.sheet_by_name("缺陷表")
        m_table = excel.sheet_by_name("反措表")
        self.device = data(dv_table)
        self.repair = data(r_table)
        self.defect = data(df_table)
        self.measure = data(m_table)

    def sql_dev(self, su):
        station = su.query_one("select id from station where name =?", (self.device[0][0],))
        if station is None:
            su.operate_one('insert into interval (name,level) values(?,?)', (self.device[0][0], '220kV'))
            station = su.query_one("select id from station where name =?", (self.device[0][0],))
        for dev in self.device:
            dev[0] = station[0]
            interval = su.query_one("select id from interval where station_id =? and name=?", (dev[0], dev[1]))
            if interval is None:
                su.operate_one('insert into interval (station_id,name) values(?,?)', (dev[0], dev[1]))
                interval = su.query_one("select id from interval where station_id =? and name=?", (dev[0], dev[1]))
            dev[1] = interval[0]
            tp = su.query_one("select id from type where name=?", (dev[2],))
            if tp is None:
                su.operate_one('insert into type (name) values(?)', (dev[2],))
                tp = su.query_one("select id from type where name=?", (dev[2],))
            dev[2] = tp[0]
            sql_1 = "select id from device where station_id = ? and interval_id = ? and type_id = ? and name=?"
            d = su.query_one(sql_1, dev)
            if d is None:
                station_num = station[0]
                while len(str(station_num)) < 2:
                    station_num = '0' + str(station_num)
                station_num = '1' + station_num
                num = su.query_one('select count() from device where station_id = ?', (station[0],))[0]+1
                num = str(num)
                while len(str(num)) < 3:
                    num = '0' + str(num)
                dev.append(station_num+num)
                su.operate_one('insert into device (station_id,interval_id,type_id,name,SD_id) values(?,?,?,?,?)', dev)

    def sql_def(self, su):
        for d in self.defect:
            dev = su.query_one("select id from device where name=? and station_id=?", (d[5], self.device[0][0]))
            d[5] = dev[0]
            if d[3] is None:
                continue
            else:
                date = datetime.datetime.fromtimestamp((d[3] - 25569) * 86400.0)
                d[3] = date.strftime('%Y-%m-%d')
            su.operate_one('insert into defect (content,flag,person,time,level,device_id) values(?,?,?,?,?,?)', d)

    def sql_repair(self, su):
        for r in self.repair:
            d = r[3].split(';')
            did = []
            r[1] = r[1].split('-')[0]
            su.operate_one('insert into repair (content,time,person) values(?,?,?)', (r[0], r[1], r[2]))
            rid = su.query_one("select id from repair ORDER BY id DESC LIMIT 1")
            for i in d:
                if i != '':
                    dev = su.query_one("select id from device where name=? and station_id=?", (i, self.device[0][0]))
                    if dev is not None:
                        did.append([dev[0], rid[0]])
            su.operate_many('insert into repair_device (device_id,repair_id) values(?,?)', did)

    def sql_meas(self, su):
        for m in self.measure:
            if m[5] != '':
                date = datetime.datetime.fromtimestamp((m[5] - 25569) * 86400.0)
                m[5] = date.strftime('%Y-%m-%d')
            dev = su.query_one("select id from device where name=? and station_id=?", (m[2], self.device[0][0]))
            m[2] = dev[0]
            su.operate_one('insert into measure (content,flag,device_id,target,person,time) values(?,?,?,?,?,?)', m)


if __name__ == '__main__':
    a = excel_ex()
    a.get_data("检修挂牌模版（1125）.xls")
    s = SqliteUtil.SqliteTool("electric.db")
    a.sql_dev(s)
    a.sql_def(s)
    a.sql_repair(s)
    a.sql_meas(s)
    s.close_con()
