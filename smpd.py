# -*- coding: utf-8 -*-
"""
Created on Wed Jun  5 17:39:21 2019
https://xlsxwriter.readthedocs.io/

@author: Stavrovskiy
"""
import sys
import numpy as np
import pandas as pd
import io

# = input, output
def pdxlsread(pathfile, wstname, header=0, skiprows=0, index_col=None, converters=None):
    '''
    Назначение: импорт листа Excel - в pandas.dataframe 
    Аргументы:
        header - номер строки с именами столбцов,
        index_col = None имена столбцов задаются как 0, 1,... 
        skiprows - число строк которые надо пропустить начиная с первой строки листа
    пример:
        
    '''
    df = None
    try:
        xlsx= pd.ExcelFile(pathfile)
        df = pd.read_excel(xlsx, sheet_name=wstname, header=header, skiprows=skiprows, index_col=index_col, converters=converters)
#        df = pd.read_excel(xlsx, sheetname=wstname, header=header, skiprows=skiprows )
    except:
        print("dfxlsread: Неожиданная ошибка (Excel): {0}\nПуть к файлу:{1}".format( sys.exc_info()[0], pathfile ))
    finally:
        return df
    
def pdxlswritedfs(pathfile,  header=True, firstrow=3, firstcol=0, **sheet_df):
    '''
    Назначение: записать в pathfile **sheet_df (имя листа1 = df1,...)
    Примечание:
        (0, 0)      # Row-column notation, ('A1') # The same cell in A1 notation.
        https://colorscheme.ru/html-colors.html
         A5:A18 формулу =НЕЧЁТ(СТРОКА())=СТРОКА()
    '''
    
    try:
        # import xlswriter
        writer = pd.ExcelWriter(pathfile, engine='xlsxwriter')
        import pandas.io.formats.excel
#        pandas.io.formats.excel.ExcelFormatter.header_style = None # обнуляет установку стиля
# формат заголовка
        cell_format = writer.book.add_format({'bold': True, 'bg_color': '#808080'})  # для заголовка
        cell_format.set_align('center'); cell_format.set_align('vcenter');cell_format.set_text_wrap()
        border_index = 5
        cell_format.set_bottom(border_index); cell_format.set_top(border_index);cell_format.set_left(border_index);cell_format.set_right(border_index)
# формат для чередования строк
        cell_format_1 = writer.book.add_format({'bg_color': '#DCDCDC'}) # для чередования строк
# формат для таблицы границы
        cell_format_2 = writer.book.add_format() # border
        border_index = 1
        cell_format_2.set_bottom(border_index); cell_format_2.set_top(border_index);cell_format_2.set_left(border_index);cell_format_2.set_right(border_index)

        # cell_format.set_pattern(1)
        for sheet, df in sheet_df.items():
            df.to_excel(writer, sheet, index=False, startrow=firstrow, startcol=firstcol, header=header)
            worksheet = writer.sheets[sheet]; dfrows, dfcols = df.shape
            worksheet.autofilter(firstrow, firstcol, firstrow+dfrows, firstcol+dfcols-1)
            worksheet.freeze_panes(firstrow+1, firstcol+1)
#           worksheet.set_column(firstcol, firstcol+dfcols-1, None, cell_format_2 ) # форматирование столбцов весь столбец
            worksheet.conditional_format(firstrow+1, firstcol, firstrow+dfrows, firstcol+dfcols-1, 
                                         {'type': 'no_errors',   'format': cell_format_2}) #  только таблицу
            # условное форматирование =ЕНЕЧЁТ(СТРОКА(XFB1048568))   ISEVEN(ROW())
            # НЕЧЁТ(СТРОКА())=СТРОКА() =ISODD(ROW())=ROW()
            worksheet.conditional_format(firstrow+1, firstcol, firstrow+dfrows, firstcol+dfcols-1, 
                                         {'type': 'formula', 'criteria': '=ISEVEN(ROW())',  'format':  cell_format_1})
            worksheet.set_row(firstrow, 32)
            for columnnum, columnname in enumerate(list(df.columns)):
                worksheet.write(firstrow, firstcol+columnnum, columnname, cell_format)
                
        writer.save()
    except:
        print (f"pdxlswritedfs Неожиданная ошибка. \n \
               {sys.exc_info()[0]}\n{sys.exc_info()[1]}\n{sys.exc_info()[2]}\n")
    finally:
        pass
        
def pdxlswritedfs_tbl(pathfile,  header=True, firstrow=1, firstcol=0, 
                      tablename='tTable01', tablestyle={'autofilter': True,'total_row': False, 'style': 'Table Style Medium 2'}, 
                      tableyn=False, **sheet_df):
    '''
    Назначение: записать в pathfile **sheet_df (имя листа1 = df1,...)
    Примечание:
        (0, 0)      # Row-column notation, ('A1') # The same cell in A1 notation.
        с форматированием таблица
        https://colorscheme.ru/html-colors.html
    '''
    
    try:
        # import xlswriter
        writer = pd.ExcelWriter(pathfile, engine='xlsxwriter')
        workbook = writer.book
        header_format = workbook.add_format({'bold': True})
        header_format.set_align('center'); header_format.set_align('vcenter');header_format.set_text_wrap()
        for sheet, df in sheet_df.items():
            df.to_excel(writer, sheet, index=False, startrow=firstrow, startcol=firstcol, header=header)
            worksheet = writer.sheets[sheet]; dfrows, dfcols = df.shape
            # header_format = workbook.add_format({'bold': True, 'bottom': 2, 'bg_color': '#F9DA04'})
            tablestyle['name'] = sheet + 'tbl'
            tablestyle['columns']=[{'header':col, 'header_format':header_format} for col in df.columns]
            worksheet.add_table(firstrow, firstcol, firstrow + dfrows, firstcol + dfcols-1, tablestyle )
        writer.save()
    except:
        print (f"pdxlswritedfs Неожиданная ошибка. \n \
               {sys.exc_info()[0]}\n{sys.exc_info()[1]}\n{sys.exc_info()[2]}\n")
    finally:
        pass
        

def pdxlswrite( df, pathfile, wstname, header=True, firstrow=0, firstcol=0, tablename='tTable01', tablestyle={'autofilter': True,'total_row': None, 'style': 'Table Style Medium 15'}, tableyn=False):
    '''
    Назначение: экспорт pandas.dataframe  в лист Excel 
    Аргументы:
         df     - pandas.dataframe
         wsname - имя листа
    Пример:
    '''
#  print(win32.constants.xlRight)
    work=False
    try:
        writer = pd.ExcelWriter(pathfile, engine='xlsxwriter')
        df.to_excel(writer, wstname, index=False, startrow=firstrow, startcol=firstcol, header=header)
        workbook = writer.book
        worksheet = writer.sheets[wstname]
        dfrow, dfcol = df.shape
        tablestyle['name']=tablename
        tablestyle['columns']=[{'header':col} for col in df.columns]
        worksheet.add_table(firstrow, firstcol, firstrow+dfrow+1, firstcol+dfcol-1, tablestyle )
        worksheet.freeze_panes(firstrow + 1, firstcol + 1)
#       worksheet.autofilter(firstrow, firstcol, firstrow+dfrow+1, firstcol+dfcol-1)
        writer.save()
        if not tableyn:
            # https://gist.github.com/airstrike/5469478
            import win32com.client as win32
            try:
                excel = win32.gencache.EnsureDispatch('Excel.Application') # ранне связывание
                excel.Visible = False
                wbk = excel.Workbooks.Open(pathfile)
                wst= wbk.Worksheets(wstname)
                wst.ListObjects[tablename].Unlist()
                rng = wst.Range(wst.Cells(firstrow+1, firstcol+1),wst.Cells(firstrow+dfrow+2, firstcol+dfcol))
#                rng.AutoFilter(1,Operator=excel.XlAutoFilterOperator.xlAnd,  VisibleDropDown=True)
                rng.AutoFilter(1)
            except:
                print("dfxlswrite com: Неожиданная ошибка win32com.client (Excel):", sys.exc_info()[0])
            finally:
                wbk.Save()
                excel.Application.Quit()
                del excel
    except:
        print ("dfxlswrite Неожиданная ошибка.", sys.exc_info()[0])
    else:
        work=True
    finally:
        return work


# === Преоразование типов
def pdastype(df, fieldlist, strsettype):
    '''
    Назначение: приводит список fieldlist полей в df к типу strsettype
    Аргументы: df - dataframe, fieldlist - список полей, strsettype - тип 
    Пример:
        df = dfastype(df, fieldlist, 'uint8' )
    Примечание:
        инфа по диапазону значений для типа 
        np.iinfo(np.int16) - целые, np.finfo(np.int16)
    '''
    df[fieldlist] = df[fieldlist].astype(strsettype)
    return (df)

def pdastypedict(df, fieldtypedict):
    '''
    Назначение: приводит список полей в df к типу на основе словаря
    Аргументы: df - dataframe, fieldtypedict - словарь  поля : типы 
    Пример:
        fieldtypedict = { 'lttype' : 'uint8', 'lt' : 'uint16'}
        df = dfastype(df, fieldtypedict )
    Примечание:
        инфа по диапазону значений для типа 
        np.iinfo(np.int16) - целые, np.finfo(np.int16)
    '''
    df = df.astype(fieldtypedict)
    return (df)
# == трансформация
def pdmelt(dft, id_fields, var_fields, var_name, value_name, dict_rename = {}, dict_sort = {}):
    '''
    Назначение: преобразовать из 2D df в df 1D
    2D df имеет вид ключевые поля id_fields + поля имеющие вид prefix + separator + числа
    Пример: преобразовать
        dft = ['idLtType', 'idLt', LtN1, ..., LtN5 ] в 
        df  = [lttype, lt, ltnord, ltn]
        id_fields = ['idLtType', 'idLt'], varfields = [LtN1, ..., LtN5], 
        var_name = 'ltnord', value_name = 'ltn', 
        dict_rename = {'idLtType' : 'lttype', 'idLt' : 'lt' }
        dict_sort = {field : ascending }
            df.sort_values(['age', 'grade'], ascending=[True, False])
    '''
    df = None
    fields = id_fields + var_fields
    df = dft[fields].melt(id_vars=id_fields, var_name=var_name, value_name=value_name)
    if len(dict_rename)>0:
        df.rename(columns=dict_rename, inplace=True)
    if len(dict_sort)>0:
        columns = list(dict_sort.keys())
        ascending = list(dict_sort.values())
        df.sort_values(columns, ascending=ascending, inplace=True)
    return df

def pdmeltidnnums(dft, idfields, prefix, nnums, var_name, value_name, dict_rename ):
    '''
    Назначение: преобразовать из 2D df в df 1D
    2D df имеет вид ключевые поля id_fields + поля имеющие вид prefix + separator + числа
    Пример: преобразовать
        dft = ['idLtType', 'idLt', LtN1, ..., LtN5 ] в 
        df  = [lttype, lt, ltnord, ltn]
        id_fields = ['idLtType', 'idLt'], prefix = 'LtN', nums = 5, 
        var_name = 'ltnord', value_name = 'ltn', 
        dict_rename = {'idLtType' : 'lttype', 'idLt' : 'lt' }
    '''
    
    return None

def pdcolendtofirst(df):
    '''
    назначение: меняет порядок столбцов в df 
    последний становиться первым, остальные не меняются
    '''
    colname = df.columns.tolist()
    colname = colname[-1:] + colname[:-1]
    df = df[colname]
    return df

# == группировка, сводные показатели
def pdgar(df, list_colgrup, dict_colagg, dict_colrename, smobserved=True):
    '''
    назначение: для pandas dataframe df сгруппировать, агригировать, переименовать, 
    переустановить индекс (grup agg rename)
    df - 
    list_colgrup - список имен столбцов для групировки
    dict_colagg - словарь для агреггирования
    dict_colrename - словарь для переименования
        observed=smobserved, Это применимо только в том случае, если какой-либо из группировщиков является категориальным. 
        Если True: показывать только наблюдаемые значения для категориальных групперов. 
    Если False: показать все значения для категориальных групперов
    пример:
        colgrup=['id1']; colagg = {'id1': 'count'}; 
        colrename = {'id1': 'id1N'}
        dfWhat = smpd.pdgar(dfLink, colgrup,  colagg, colrename)
    '''
    dfout = df.groupby(list_colgrup, observed=smobserved).agg(dict_colagg).rename(
            columns=dict_colrename).reset_index()
    return dfout 
# 
def pdgar_mui(df, list_colgrup, dict_colagg, dict_colrename=None, sep_col_level='', smobserved=True, smdropna=True ):
    '''
    Особенность: при наличии MultiIndex в столбцах сворачивает до 1 уровня через разделитель sep_col_level
    назначение: для pandas dataframe df сгруппировать, агригировать, переименовать, 
    переустановить индекс (grup agg rename)
    df - 
    list_colgrup - список имен столбцов для групировки
    dict_colagg - словарь для агреггирования
    dict_colrename - словарь для переименования
    observed =smobserved Это применимо только в том случае, если какой-либо из группировщиков является категориальным. 
    Если True: показывать только наблюдаемые значения для категориальных групперов. 
    Если False: показать все значения для категориальных групперов
    smdropna = True,значения NA - игнорируются
        colgrup=['id1']; colagg = {'id1': 'count'}; 
        colrename = {'id1': 'id1N'}
        dfWhat = smpd.pdgar(dfLink, colgrup,  colagg, colrename)

    '''

    dfout = df.groupby(list_colgrup, observed=smobserved, dropna=smdropna).agg(dict_colagg).rename(
            columns=dict_colrename).reset_index()
    if isinstance(dfout.columns, pd.MultiIndex):
        dfout.columns = dfout.columns.to_series().apply(lambda x: sep_col_level.join(x))
        # столбцы в группировке без sep_col_level
        dd = {x + sep_col_level:x for x in list_colgrup}
        dfout.rename(columns = dd, inplace = True)
    return dfout

def pdgar_size(df, list_colgrup, name_size="Size", smdropna=True):
    '''
    назначение: для pandas dataframe df сгруппировать и вычислить размер группы  
    переустановить индекс (size -все включая NA, count - без NA)
    df - 
    list_colgrup - список имен столбцов для групировки
    name_size - имя столбца для числа значений в группе
    smdropna = True,значения NA - игнорируются
    пример:
        
    '''
    if name_size == "Size":
        dfout = df.groupby(list_colgrup, dropna=smdropna).size().to_frame(name_size).reset_index()
    else:
        dfout = df.groupby(list_colgrup, dropna=smdropna).count().to_frame(name_size).reset_index()
    return dfout 


def pdgruprank(df, list_colgrup, colforrank, colnamerank, rankmethod='first', reset_index = False, ascending=True):
    '''
    Назначение: вычисляет и добаляет в df новый столбец с рейтингом
    list_colgrup - список имен столбцов для групировки
    colforrank  - имя столбца по знаячению которого считается рейтинг
    colnamerank - имя столбца которому присваивается рейтинг
    rankmethod - способ вычисления рейтинга
        method : {‘average’, ‘min’, ‘max’, ‘first’, ‘dense’}
        average: average rank of group
        min: lowest rank in group,max: highest rank in group
        first: ranks assigned in order they appear in the array
        dense: like ‘min’, but rank always increases by 1 between groups
    reset_index = True выполн переустановку индекса df.reset_index(inplace = True)
    пример:
     field = ['lttype', 'ltw', 'rtgrow']
     dfc["lnrtgord"] = dfc.groupby(field)[rtgsel].rank(method='first')

    '''
    df.loc[:,colnamerank] = df.groupby(list_colgrup)[colforrank].rank(method=rankmethod, ascending=ascending)
    if reset_index:
        df.reset_index(inplace = True)
    return df

def pddfinfo_to_str(df):
    '''
    возвращает строковое представление df.info()
    '''
    buf = io.StringIO();  df.info(buf=buf); s = buf.getvalue()
    return s

def pddfdescribe_to_str(df):
    '''
    
    возвращает строковое представление df.info()
    with pd.option_context('display.max_columns', 40):
    print(df_pr.describe(include='all').T)
    '''
    buf = io.StringIO();  df.describe(buf=buf); s = buf.getvalue()
    return s
# ================================================
if __name__ == "__main__":
    df = pd.DataFrame({'Data_N1': [10, 20, 30, 20, 15, 30, 45], 'Data_N2': ['S10', 's20', 's30', 's20', 's2110', 's2220', 's45'], })
    pdxlswritedfs_tbl('pandas_test.xlsx', **{'test':df})
    pdxlswritedfs('pandas_test_1.xlsx', **{'test':df})
    