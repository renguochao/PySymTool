# coding=UTF-8

import os
import time
import xlrd
import xlsxwriter
import threading
import math
import datetime

# excel 名称
excel_name = '20170223_4.0.1_crash'

app_dsym_mapping = {
    '3.4.2' : '882',
    '4.0.0' : '1365',
    '4.0.1' : '1494',
}

only_symbolicate_latest_version = True
latest_version = '4.0.1'
symbol_col = 13
deviceId_col = 5

def get_raw_workbook():
    workbook = xlrd.open_workbook(excel_name + '.xlsx')
    return workbook
    
def get_output_workbook():
    output_excel_name = excel_name + '_multithread_result_py.xlsx'
    output_workbook = xlsxwriter.Workbook(output_excel_name)

    return output_workbook

def get_output_worksheet(output_workbook):
    output_worksheet = output_workbook.add_worksheet()
    # 设置列宽
    output_worksheet.set_column('A:D', 20)
    # 粗体格式
    bold = output_workbook.add_format({'bold': True})
    # 标题行
    output_worksheet.write('A1', 'deviceId', bold)
    output_worksheet.write('B1', 'os_version', bold)
    output_worksheet.write('C1', 'exception type', bold)
    output_worksheet.write('D1', 'symbolication result', bold)

    return output_worksheet


class WorkThread (threading.Thread):
    def __init__(self, name, start_index, end_index, output_worksheet):
        threading.Thread.__init__(self)
        self.name = name
        self.start_index = start_index
        self.end_index = end_index
        self.output_worksheet = output_worksheet
        
    def run(self):
        raw_workbook = get_raw_workbook()
        raw_worksheet = raw_workbook.sheets()[0]

        symbolicate_crash(raw_worksheet, self.output_worksheet, self.start_index, self.end_index, self.name)
        
def symbolicate_crash(raw_worksheet, output_worksheet, start, end, thread_name):
    if (start > end):
        return

    # 写入 Excel Index 指示器
    result_row_index = start
    result_col_index = 0
    for row_index in range(start, end):
        row = raw_worksheet.row_values(row_index)
        
        # 取得 crash_symbol 那一列数据
        # print('thread_name:' + thread_name + ' row_index:' + str(row_index) + ' col_index:' + str(symbol_col))
        crash_symbol = row[symbol_col]
        device_id = row[deviceId_col]
        if crash_symbol == None:
            print('Empty Column in row ' + str(row_index))
            continue
        
        # 清理数据
        crash_symbol = crash_symbol.replace('\"', '')
        crash_symbol = crash_symbol.replace('\n', '\\n')
        # print(str(len(crash_symbol)) + ' ' + crash_symbol)
        
        # 分割字符串
        if len(crash_symbol) > 0:
            array = crash_symbol.split('\\n')
            if not array[0].startswith('Incident'):
                print('Column ' + str(symbol_col) + ' does not contain Incident')
                continue
            
            binary_images_found = 0
            previous_part = ''
            duplicate_line_num = []
            for array_index in range(len(array)):
                part = array[array_index]
                if previous_part != part:
                    previous_part = part
                else:
                    # 当前行与上一行相同会导致符号化失败, 需删除
                    duplicate_line_num.append(array_index)

                if part.find('Binary Images:') != -1:
                    binary_images_found = 1

            array_copy = []
            for array_index in range(len(array)):
                if array_index in duplicate_line_num:
                    continue

                array_copy.append(array[array_index])

            array = array_copy

            if binary_images_found == 0:
                print('Row: ' + str(row_index) + ' Binary Images Not Found!')
                continue
                
            count = len(array)
            # print('count: ' + str(count) + ' array:' + str(array))
            hardware_model = array[2][27:28]
            platform = 'armv7' if int(hardware_model) <=5 else 'arm64'
            version = array[6][17:22]
            build = array[6][24:][:-1]
            os_version = array[11][17:]
            exception_type = array[14][17:]
            
            if only_symbolicate_latest_version:
                if version != latest_version:
                    print('Row:' + str(row_index) + ' crash log version: ' + version + ' binary version: ' + latest_version + ' does not match!')
                    continue
            
            print('Row: ' + str(row_index) + ' version: ' + version + ' build: ' + build + ' os_version:' + os_version + ' platform: ' + platform + ' exception_type:' + exception_type)
            crash_symbol = '\n'.join(array)
            # print(crash_symbol)
            
            # 将 crash 堆栈写入临时文件
            timestamp = int(time.time())
            tmp_result = '{0}_{1}'.format(thread_name, str(timestamp))
            tmp_f = open(str(tmp_result), 'w')
            tmp_f.write(crash_symbol)
            tmp_f.close()
            
            # 符号化堆栈
            os.putenv('DEVELOPER_DIR', '/Applications/Xcode7.app/Contents/Developer')
            symbolicate_command = './symbolicatecrash ' + str(tmp_result) + ' -d tztHuaTaiZLMobile.app.' + build + '.dSYM -o ' + str(tmp_result) + '.result'
            print(symbolicate_command)
            os.system(symbolicate_command)
            # 结果读入内存
            with open(str(tmp_result) + '.result') as result_f:
                result_symbol = result_f.readlines()    

            # 结果写入文件
            output_worksheet.write(result_row_index, result_col_index, device_id)
            result_col_index += 1
            output_worksheet.write(result_row_index, result_col_index, os_version)
            result_col_index += 1
            output_worksheet.write(result_row_index, result_col_index, exception_type)
            result_col_index += 1
            output_worksheet.write(result_row_index, result_col_index, ''.join(result_symbol))
            
            result_row_index += 1
            result_col_index = 0
            
            # 删除临时文件
            rm_command_01 = 'rm ' + str(tmp_result)
            rm_command_02 = 'rm ' + str(tmp_result) + '.result'
            os.system(rm_command_01)
            os.system(rm_command_02)
        
def main():
    begin_time = datetime.datetime.now()

    raw_worksheet = get_raw_workbook().sheets()[0]
    row_num = raw_worksheet.nrows

    row_per_thread = 30
    thread_num = math.ceil(row_num / row_per_thread)

    output_workbook = get_output_workbook()
    output_worksheet = get_output_worksheet(output_workbook)

    threads = []

    for thread_index in range(0, thread_num):
        threadName = 'Thread{}'.format(thread_index)
        start_index = 1 + thread_index * row_per_thread
        if thread_index == thread_num - 1:
            end_index = row_num
        else:
            end_index = 1 + row_per_thread + thread_index * row_per_thread

        t = WorkThread(threadName, start_index, end_index, output_worksheet)
        threads.append(t)

    for i in range(0, thread_num):
        threads[i].start()

    for i in range(0, thread_num):
        threads[i].join()

    # 关闭 output_workbook
    output_workbook.close();

    end_time = datetime.datetime.now()
    print('程序耗时:' + str(end_time - begin_time))

    print('Exit main')
  

if __name__ == '__main__':
    main()