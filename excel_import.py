import SqliteUtil
import xlwt


def im():
    s = SqliteUtil.SqliteTool("electric.db")
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('get')
    sql_str = "select d.name,d.SD_id,i.name from device AS d left join interval AS i on d.interval_id=i.id where i.station_id=7"
    data = s.query_many(sql_str)
    count = 0
    print(data)
    for i in data:
        worksheet.write(count, 0, "220kV玉鹤变")
        worksheet.write(count, 1, i[2])
        worksheet.write(count, 2, i[0])
        worksheet.write(count, 3, i[1])
        count += 1
    workbook.save('import.xls')


if __name__ == '__main__':
    im()
