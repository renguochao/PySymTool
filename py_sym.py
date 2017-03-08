# coding=UTF-8

import os
import time
import xlrd
import xlsxwriter
from datetime import datetime

# excel 名称
excel_name = '20170223_4.0.1_crash'

app_dsym_mapping = {
    '3.4.2' : '882',
    '4.0.0' : '1365',
    '4.0.1' : '1494',
}

only_symbolicate_latest_version = True
latest_version = '4.0.1'

def main():
    begin_time = datetime.now()
    
    excel = xlrd.open_workbook(excel_name + '.xlsx')
    symbol_col = 13
    deviceId_col = 5
    
    symbolicate(excel, symbol_col, deviceId_col)
    
    end_time = datetime.now()
    print('符号化耗时:' + str(end_time - begin_time))
    
def symbolicate(excel, symbol_col, deviceId_col):
    sheet = excel.sheets()[0]
    
    # Excel 的行和列
    nrows = sheet.nrows
    ncols = sheet.ncols
    
    # 创建符号化结果保存的 excel
    result_excel_name = excel_name + '_result_py.xlsx'
    os.system('rm -rf ' + result_excel_name)
    workbook = xlsxwriter.Workbook(result_excel_name)
    worksheet = workbook.add_worksheet()

    # 设置列宽
    worksheet.set_column('A:D', 20)
    
    # 粗体格式
    bold = workbook.add_format({'bold' : True})
    
    # 标题行
    worksheet.write('A1', 'deviceId', bold)
    worksheet.write('B1', 'os_version', bold)
    worksheet.write('C1', 'exception type', bold)
    worksheet.write('D1', 'symbolication result', bold)

    # 写入 Excel Index 指示器
    result_row_index = 1
    result_col_index = 0
    
    for row_index in range(1, nrows):
        row = sheet.row_values(row_index)
        
        # 取得 crash_symbol 那一列数据
        crash_symbol = row[symbol_col]
        deviceId = row[deviceId_col]
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
            tmp_result = int(time.time())
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
            worksheet.write(result_row_index, result_col_index, deviceId)
            result_col_index += 1
            worksheet.write(result_row_index, result_col_index, os_version)
            result_col_index += 1
            worksheet.write(result_row_index, result_col_index, exception_type)
            result_col_index += 1
            worksheet.write(result_row_index, result_col_index, ''.join(result_symbol))
            
            result_row_index += 1
            result_col_index = 0
            
            # 删除临时文件
            rm_command_01 = 'rm ' + str(tmp_result)
            rm_command_02 = 'rm ' + str(tmp_result) + '.result'
            os.system(rm_command_01)
            os.system(rm_command_02)
            
    # 关闭文件
    workbook.close()    
  

if __name__ == '__main__':
    main()