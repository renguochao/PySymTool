# coding=UTF-8

import mysql.connector
import xlrd
import xlsxwriter
import os

from mysql.connector import errorcode
from datetime import datetime

# 符号化后的 Excel 文件名
EXCEL_NAME = '20170223_4.0.1_feedback_result_py'
DB_NAME = 'zl_crash'

config = {
    'user': 'root',
    'password': '123456',
    'host': '127.0.0.1',
    'database': 'zl_crash',
}

class Report(object):
    '''
    Report class used to encapsulate the row data in EXCEL
    '''

    def __init__(self, report_id, exception_type, device_id, exception_symbols, os_version):
        self.report_id = report_id;
        self.exception_type = exception_type;
        self.device_id = device_id;
        self.exception_symbols = exception_symbols;
        self.os_version = os_version;


def main():
    begin_time = datetime.now()

    # 表名
    table_name = 'report_' + begin_time.strftime("%Y_%m_%d_%H_%M_%S")

    # 建表
    create_table_in_db(table_name)

    # 插入数据
    insert_symbolication_result_into_db(table_name)

    # 对数据进行分组并导出
    generate_grouped_exception(table_name)

    end_time = datetime.now()

    print('耗时:' + str(end_time - begin_time))


def create_table_in_db(table_name):
    '''
    Create a table in database, and named as `table_name`
    :param table_name: table_name
    '''

    SQLS = {}
    SQLS['drop_report'] = (
        "DROP TABLE IF EXISTS `" + table_name + "`")

    SQLS['report'] = (
        "CREATE TABLE `" + table_name + "` ( "
        "`report_id` int(11) NOT NULL AUTO_INCREMENT, "
        "`exception_type` varchar(255) DEFAULT NULL, "
        "`device_id` varchar(255) DEFAULT NULL, "
        "`exception_symbols` longtext, "
        "`os_version` varchar(255) DEFAULT NULL, "
        "PRIMARY KEY (`report_id`)"
        ") ENGINE=InnoDB DEFAULT CHARSET=utf8")

    try:
        conn = mysql.connector.connect(**config)
        cursor = conn.cursor();

        for name, sql in SQLS.items():
            try:
                print("Executing sql {}.".format(name))
                cursor.execute(sql)
            except mysql.connector.Error as err:
                if err.errno == errorcode.ER_TABLE_EXISTS_ERROR:
                    print('Table already exists.')
                else:
                    print(err.msg)

    except mysql.connector.Error as err:
        if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
            print("Something is wrong with your user name or password")
        elif err.errno == errorcode.ER_BAD_DB_ERROR:
            print("Database does not exist")
        else:
            print(err.msg)
    finally:
        cursor.close()
        conn.close()


def insert_symbolication_result_into_db(table_name):
    '''
    Insert the symbolicated result into database
    :param table_name: table_name in database
    '''

    try:
        conn = mysql.connector.connect(**config)
        # print('connected to db')

        cursor = conn.cursor()
        insert_report = (
            "INSERT INTO " + table_name + " "
            "(exception_type, device_id, exception_symbols, os_version) "
            "VALUES (%s, %s, %s, %s)")

        work_book = xlrd.open_workbook(EXCEL_NAME + '.xlsx')
        sheet = work_book.sheets()[0]
        nrows = sheet.nrows
        ncols = sheet.ncols
        row_index = 1

        for row_index in range(1, nrows):
            data_row = sheet.row_values(row_index)

            # assert col < ncols
            device_id = data_row[0]
            os_version = data_row[1]
            exception_type = data_row[2]
            exception_symbols = data_row[3]

            if exception_symbols == '':
                continue

            data_report = (exception_type, device_id, exception_symbols, os_version)

            # insert report data
            cursor.execute(insert_report, data_report)

        conn.commit()

    except mysql.connector.Error as err:
        if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
            print("Something is wrong with your user name or password")
        elif err.errno == errorcode.ER_BAD_DB_ERROR:
            print("Database does not exist")
        else:
            print(err.msg)
    finally:
        cursor.close()
        conn.close()


def generate_grouped_exception(table_name):
    '''
    According the group data in database, make all exception to group data.
    :param table_name: table_name in zl_crash database
    '''

    EXCEPTION_TYPE_COUNT = {}
    EXCEPTION_MAPPING = {}
    try:
        conn = mysql.connector.connect(**config)
        cursor = conn.cursor()

        group_exception_type = (
            "SELECT exception_type, COUNT(*) as nums "
            "FROM " + table_name + " GROUP BY exception_type")
        query_specific_exception = (
            "SELECT * FROM " + table_name + " "
            "WHERE exception_type = %s")

        cursor.execute(group_exception_type)

        for (exception_type, nums) in cursor:
            EXCEPTION_TYPE_COUNT[exception_type] = nums
            # print("exception_type:" + exception_type + ", nums:" + str(nums))

        for exception_type in EXCEPTION_TYPE_COUNT.keys():
            cursor.execute(query_specific_exception, (exception_type,))
            exception_list = []
            for (report_id, exception_type, device_id, exception_symbols, os_version) in cursor:
                report = Report(report_id, exception_type, device_id, exception_symbols, os_version)
                exception_list.append(report)

            EXCEPTION_MAPPING[exception_type] = exception_list

        write_grouped_exception_to_file(EXCEPTION_TYPE_COUNT, EXCEPTION_MAPPING)

    except mysql.connector.Error as err:
        if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
            print("Something is wrong with your user name or password")
        elif err.errno == errorcode.ER_BAD_DB_ERROR:
            print("Database does not exist")
        else:
            print(err.msg)
    finally:
        cursor.close()
        conn.close()


def write_grouped_exception_to_file(count, mapping):
    '''
    Export grouped exception to file
    :param count: 字典 key:exception_type value:count
    :param mapping: 字典 key:exception_type value:exception_list
    '''

    output_file_name = EXCEL_NAME + '_grouped.xlsx'
    os.system('rm -rf ' + output_file_name)
    workbook = xlsxwriter.Workbook(output_file_name)
    worksheet = workbook.add_worksheet()

    # 设置列宽
    worksheet.set_column('A:A', 25)
    worksheet.set_column('B:B', 10)
    worksheet.set_column('C:C', 25)
    worksheet.set_column('D:D', 40)
    worksheet.set_column('E:E', 500)

    # 粗体格式
    bold = workbook.add_format({'font_size': 14,
                                'align': 'center',
                                'bold': True})

    # 标题行
    worksheet.write('A1', 'exception_type', bold)
    worksheet.write('B1', 'count', bold)
    worksheet.write('C1', 'os_version', bold)
    worksheet.write('D1', 'device_id', bold)
    worksheet.write('E1', 'symbols', bold)

    # 写入 Excel Index 指示器
    row_index = 1
    col_index = 0

    colors = ('#A8BAAA', '#FFF6CF', '#DCCDAE', '#B49D7E',
              '#816854', '#334D5C', '#45B29D', '#EFC94C')

    count_index = 0
    pattern = 0.5
    for (type, num) in count.items():
        bg_color = colors[count_index % len(colors)]
        col_format = workbook.add_format({'pattern': pattern,
                                          'bg_color': bg_color})
        num_col_format = workbook.add_format({'pattern': pattern,
                                              'bg_color': bg_color,
                                              'bold': True,
                                              'align': 'center'})
        count_index += 1

        list = mapping[type]

        for i in range(num):
            report_item = list[i]
            if i == 0:
                worksheet.write(row_index, col_index, report_item.exception_type, col_format)
                col_index += 1
                worksheet.write(row_index, col_index, num, num_col_format)
                col_index += 1
            else:
                worksheet.write(row_index, col_index, '', col_format)
                col_index += 1
                worksheet.write(row_index, col_index, '', col_format)
                col_index += 1
            worksheet.write(row_index, col_index, report_item.os_version, col_format)
            col_index += 1
            worksheet.write(row_index, col_index, report_item.device_id, col_format)
            col_index += 1
            worksheet.write(row_index, col_index, report_item.exception_symbols, col_format)

            # 设置 index
            row_index += 1
            col_index = 0

    # 关闭文件
    workbook.close()
    print("Exporting grouped data to " + output_file_name)


if __name__ == '__main__':
    main()
