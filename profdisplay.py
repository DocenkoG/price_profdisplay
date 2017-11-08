# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
#import io
import sys
import configparser
#import time
import shutil
import openpyxl                       # Для .xlsx
# import xlrd                         # для .xls
from   price_tools import getCellXlsx, quoted, dump_cell, currencyType, subInParentheses



def convert_sheet( book, sheetName):
    confName = ('cfg_'+sheetName.replace(' ','').replace('.','')+'.cfg').lower()
    csvFName = ('csv_'+sheetName.replace(' ','').replace('.','')+'.csv').lower()
    if not os.path.exists( confName ) :
        log.error( 'Нет файла конфигурации '+confName)
        return
    # Прочитать конфигурацию из файла
    in_columns_j, out_columns = config_read( confName )
    sh = book[sheetName]                                     # xlsx
    ssss = []
    for i in range(sh.min_row, sh.max_row) :
        i_last = i
        try:
            ccc = sh.cell(row=i, column= 2 )
            if ccc.value == 'Категория дисплея' :                   # Заголовок таблицы
                pass
    
            elif ccc.value == None  :                               # Пустая строка
                pass
                #print( 'Пустая строка. i=', i )
    
            else :                                                  # Информационная строка
                impValues = getXlsxString(sh, i, in_columns_j)
                sss = []                                            # формируемая строка для вывода в файл
                for outColName in out_columns.keys() :
                    shablon = out_columns[outColName]
                    for key in impValues.keys():
                        if shablon.find(key) >= 0 :
                            shablon = shablon.replace(key, impValues[key])
                    if outColName == 'описание':    shablon = appendSensor( shablon, impValues)
                    sss.append( quoted( shablon))
                ssss.append(','.join(sss))
                    
        except Exception as e:
            log.debug('Exception: <' + str(e) + '> при обработке строки ' + str(i) +'<' + '>' )
            raise e

    log.info('Обработано ' +str(i_last)+ ' строк.')
    f2 = open( csvFName, 'w', encoding='cp1251')
    strHeader = ','.join( out_columns.keys() ) + ','
    f2.write( strHeader + '\n' )
    data = ',\n'.join(ssss) +','
    bbbb = data.encode(encoding='cp1251', errors='replace')
    data = bbbb.decode(encoding='cp1251')
    f2.write(data)
    f2.close()



def appendSensor( shablon, impValues):
    ss = impValues['тип_сенсора']
    if ss != 'нет' :  shablon = shablon + '\nтип сенсора: ' + ss
    ss = impValues['количество_точек_касания']
    if ss != 'нет' :  shablon = shablon + '\nколичество точек касания: ' + ss
    return shablon



def getXlsxString(sh, i, in_columns_j):
    impValues = {}
    for item in in_columns_j.keys() :
        if item in ('закупка','продажа','цена') :
            impValues[item] = getCellXlsx(row=i, col=in_columns_j[item], isDigit='Y', sheet=sh) 
        else:
            impValues[item] = getCellXlsx(row=i, col=in_columns_j[item], isDigit='N', sheet=sh)
    return impValues



def config_read( cfgFName ):
    log.debug('Reading config ' + cfgFName )
    
    config = configparser.ConfigParser()
    if os.path.exists(cfgFName):     config.read( cfgFName)
    else : log.debug('Не найден файл конфигурации.')

    # в разделе [cols_in] находится список интересующих нас колонок и номера столбцов исходного файла
    in_columns_names = config.options('cols_in')
    in_columns_j = {}
    for vName in in_columns_names :
        if ('' != config.get('cols_in', vName)) :
            in_columns_j[vName] = config.getint('cols_in', vName) 
    
    # По разделу [cols_out] формируем перечень выводимых колонок и строку заголовка результирующего CSV файла
    out_columns_names = config.options('cols_out')
    out_columns = {}
    for vName in out_columns_names :
        if ('' != config.get('cols_out', vName)) :
            out_columns[vName] = config.get('cols_out', vName) 

    return in_columns_j, out_columns



def convert2csv( dealerName ):
    fileNameIn = 'new_'+dealerName+'.xlsx'
    book = openpyxl.load_workbook(filename = fileNameIn, read_only=False, keep_vba=False, data_only=False)
#   book = xlrd.open_workbook( fileNameIn.encode('cp1251'), formatting_info=True)
    sheetNames = book.get_sheet_names()
    for sheetName in sheetNames :                                # Организую цикл по страницам
        log.info('-------------------  '+sheetName +'  ----------')
        if   sheetName == 'Samsung': convert_sheet( book, sheetName)
        elif sheetName == 'LG'     : convert_sheet( book, sheetName)
        elif sheetName == 'NEC'    : convert_sheet( book, sheetName)
        #else : log.debug('Не конвертируем лист '+sheetName )



def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')



def main( dealerName):
    make_loger()
    log.info('         '+dealerName )
    convert2csv( dealerName )
    if os.path.exists( 'python.log') : shutil.copy2( 'python.log', 'c://AV_PROM/prices/' + dealerName +'/python.log')
    if os.path.exists( 'python.1'  ) : shutil.copy2( 'python.log', 'c://AV_PROM/prices/' + dealerName +'/python.1'  )



if __name__ == '__main__':
    myName = os.path.basename(os.path.splitext(sys.argv[0])[0])
    mydir    = os.path.dirname (sys.argv[0])
    print(mydir, myName)
    main( 'profdisplay')
