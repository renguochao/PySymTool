# coding=UTF-8

import os
import time
import xlrd
import xlsxwriter

# excel 名称
excel_name = 'query_result'

app_dsym_mapping = {
    '3.4.2' : '882',
    '4.0.0' : '1365',
}

def main():
    excel = xlrd.open_workbook(excel_name + '.xlsx')
    symbol_col = 8
    
    symbolicate(excel, symbol_col)
    
def symbolicate(excel, col):
    sheet = excel.sheets()[0]
    
    nrows = sheet.nrows
    ncols = sheet.ncols
    
    # 创建符号化结果保存的 excel
    result_excel_name = excel_name + '_result_py.xlsx'
    os.system('rm -rf ' + result_excel_name)
    workbook = xlsxwriter.Workbook(result_excel_name)
    worksheet = workbook.add_worksheet()

    # 设置列宽
    worksheet.set_column('A:C', 20)
    
    # 粗体格式
    bold = workbook.add_format({'bold' : True})
    
    # 标题行
    worksheet.write('A1', 'os_version', bold)
    worksheet.write('B1', 'exception type', bold)
    worksheet.write('C1', 'symbolication result', bold)

    # 写入 Excel Index 指示器
    result_row_index = 1
    result_col_index = 0
    
    for row_index in range(1, nrows):
    # for row_index in range(35, 38):
        row = sheet.row_values(row_index)
        
        # 取得 crash_symbol 那一列数据
        crash_symbol = row[col]
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
                continue
            
            binary_images_found = 0    
            for part in array:
                if part.find('Binary Images:') != -1:
                    binary_images_found = 1
                    
            if binary_images_found == 0:
                print('Row: ' + str(row_index) + ' Binary Images Not Found!')
                continue
                
            count = len(array)
            # print('count: ' + str(count) + ' array:' + str(array))
            hardware_model = array[2][27:28]
            platform = 'armv7' if int(hardware_model) <=5 else 'arm64'
            version_with_build = array[6][17:22]
            version = array[6][24:][:-1]
            os_version = array[11][17:]
            exception_type = array[14][17:]
            
            print('Row: ' + str(row_index) + ' version: ' + version + ' os_version:' + os_version + ' build: ' + version_with_build + ' platform: ' + platform + ' exception_type:' + exception_type)
            crash_symbol = '\n'.join(array)
            # print(crash_symbol)
            
            # 将 crash 堆栈写入临时文件
            tmp_result = int(time.time())
            tmp_f = open(str(tmp_result), 'w')
            tmp_f.write(crash_symbol)
            tmp_f.close()
            
            # 符号化堆栈
            os.putenv('DEVELOPER_DIR', '/Applications/XCode.app/Contents/Developer')
            symbolicate_command = './symbolicatecrash ' + str(tmp_result) + ' -d tztHuaTaiZLMobile.app.' + version + '.dSYM -o ' + str(tmp_result) + '.result'
            os.system(symbolicate_command)
            # 结果读入内存
            with open(str(tmp_result) + '.result') as result_f:
                result_symbol = result_f.readlines()    

            # 结果写入文件
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