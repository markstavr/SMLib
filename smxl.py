# -*- coding: utf-8 -*-
"""
Created on Sat Apr 10 23:18:45 2021
Модуль для работы с Excel, основан на библиотеке xlwings
@author: Марк
http://m3.bars-open.ru/stories/style-guide-python.html#nomination
https://github.com/DagW/xlwings_examples/blob/master/examples.py
правила именований
имена функйий - маленькие буквы с подчеркиванием

"""
__autor__ = "MAS"
import xlwings  as xw
import pandas as pd


class CSMXl:
    _xlApp = None
    
    def __init__(self):
        pass
# xl_ - функции для работы с приложением Excel
    def xl_open(self, xlAppVisible = True, xlDisplayAlerts = False, xlScreenUpdate = True):

        self._xlApp = xw.App(visible=xlAppVisible, add_book=False)
        self._xlApp.display_alerts = True #xlDisplayAlerts False - подавляет предупреждения
        self._xlApp.screen_updating = xlScreenUpdate

    def xl_quit(self):
        self._xlApp.quit()

# === wb_ - функции для работы с wokbook
    def wb_get_wb(self, wbFullName):
        '''
        Parameters
        ----------
        wbFullName : строка, полное имя раб книги

        Returns
        -------
        возвращает рабочую книгу или None если книга не открыта

        '''
        wb = None
        try:
            wb = self._xlApp.books[wbFullName]
        except Exception as err:
            msg = f"ОШИБКА!? {__name__}.wb_get_wb\nТип:{type(err)}, Ошибка: {err}\n"
            msg += f"wb:{wb}, файл: {wbFullName}, не найден в открытых файлах \nxlApp.books: {self._xlApp.books} "
            print(msg)
        finally:
            return wb             
            
        # if len(self._xlApp.books)>0:
        #     wbdict = {wb_.fullname:wb_ for wb_ in self._xlApp.books}
        #     wb = wbdict.get(wbFullName) # если так писать .get(wbFullName, default = None) то не работает
        #     msg= f"функция: wb_get_wb\nxlApp.books {len(self._xlApp.books)},  wb:{wb} "
        #     print(msg)
        # return wb

    def wb_open(self, wbFullName):
        '''
        Открыть раб.книгу 
        Parameters
        ----------
        wbFullName : строка, полное имя или имя раб книги
        Returns
        -------
        возвращает wb, если все Ок, None в противном случаи
        '''             
        wb = None
        try:
            wb = self._xlApp.books.open(wbFullName)
            # print(f"2) функция: wb_open\nwb: {wb}\nself._xlApp.books: {len(self._xlApp.books)} ")
           
        except Exception as err:
            msg = f"функция: wb_open\nПроизошла ошибка\n{err} "
            print(msg)
        finally:
            return wb 
        
    def wb_save_as(self, wb, wbNewFullName = None):
        '''
        Сохранить раб.книгу в новом месте
        Parameters
        ----------
        wb: раб книги
        wbNewFullName  : строка, полное новое имя раб книги, если None ( путь не указан ) и файл не был сохранен ранее, 
        он сохраняется в текущем рабочем каталоге с текущим именем файла. Существующие файлы перезаписываются без запроса.
       
        Returns
        -------
        возвращает True, если все Ок
        ''' 
        is_save_as = False
        try:
        #     if wbNewFullName is None:
        #         wb.save()
        #     else:
        #         wb.save(wbNewFullName) #сохраняет под новым именем
            wb.save(wbNewFullName)
            is_save_as = True
        except Exception as err:
            msg = f"функция: wb_save_as\nПроизошла ошибка\n{err}\n wb.Name: {wb.name}, wbNewFullName: {wbNewFullName}"
            print(msg)
        finally:
            return is_save_as  
        
#  wst_ - функции для работы с woksheet
    def wst_get_wst(self, wb, wstName):
        '''
        Parameters
        ----------
        wbFullName : строка, полное имя раб книги или имя
        wstName : строка,  имя листа раб книги

        Returns
        -------
        возвращает кортеж (wb - раб книгу и wst - лист рабочей книги)
        wb = None если книга не открыта, wst = None нет листа в wb

        '''
        wst = None
        try:
            wst = wb.sheets[wstName]
        except Exception as err:
            msg = f"ОШИБКА!? {__name__}.wst_get_wst\nТип: {type(err)}, Ошибка: {err}\n"
            msg += f"wb:{wb}, файл: {wb.name}, лист: {wstName} не найден в коллекции листов\nwb.Sheets: {wb.sheets} "
            print(msg)
        finally:
            return wst 
   
# === rng_ - функции для работы с Range            
    def rng_get_rng(self, wb, wstName, rngNameStart, rngNameEnd = None, NameDirection=None, NameExpand=None):
        '''
        Получить диапазон 

        Parameters
        ----------
        wb :рабочая книга
        wstName: имя листа
        rngNameStart : начало диапазона лев край или имя именнованного диапазона (нотация A1, кортеж(r, c))
        rngNameУтв: конец диапазона прав край или имя именнованного диапазона (нотация A1, кортеж(r, c))
        NameDirection : для автоматическо формирования диапазона 'up', 'down', 'right', 'left'
        NameExpand : для автоматическо формирования диапазона 'table' (=down and right), 'down', 'right'.

        Returns
        -------
        range, или None

        '''
        rng = None
        try:
            if wstName is None:
                rng = wb.Range(rngNameStart) #именованный диапазон 
            else:
                wst = self.wst_get_wst(wb, wstName)
                if rngNameEnd is not None:
                    rng = wst.range(rngNameStart, rngNameEnd)
                elif NameDirection is not None:
                    rng = wst.range(rngNameStart).end(NameDirection)
                elif NameExpand is not None:
                    rng = wst.range(rngNameStart).expand(NameExpand)                    
                else:
                    rng = wst.range(rngNameStart)
        except Exception as err:            
            msg = f"ОШИБКА!? {__name__}.wst_get_wst\nТип: {type(err)}\nОшибка: {err}\n"
            msg += f"wb:{wb}, файл: {wb.name}, лист: {wstName}\nНе найден диапазон: "
            msg += f"rngNameStart: {rngNameStart}, rngNameEnd: {rngNameEnd}, NameDirection: {NameDirection}, NameExpand: {NameExpand}"
            print(msg)
        finally:
            return rng      

    def rng_copy_paste(self, rngCopy, rngPasteStart, PasteAs="all", PasteOperation = None, 
                       PasteSkipBlanks=False, PasteTranspose=False):
        '''
        Скопировать диапазон rngCopy вставить в диапазон 

        Parameters
        ----------
        rngCopy : Range for Copy
        rngPaste : Range
            начало диапазона вставки лев край 
        PasteAs : strings
            all_merging_conditional_formats, all, all_except_borders, all_using_source_theme, 
            column_widths, comments, formats, formulas, formulas_and_number_formats, validation, 
            values, values_and_number_formats. 
            The default is "all".
        PasteOperation : strings, 
            add”, “divide”, “multiply”, “subtract”.
            The default is None.
        PasteSkipBlanks : logical
            The default is False.
        PasteTranspose : logical
            The default is False.

        Returns
        -------
        True если Ок, False

        '''
        is_copy_paste = False
        try:
          rngCopy.copy()
          rngPasteStart.paste(paste=PasteAs, operation=PasteOperation, 
                                        skip_blanks=PasteSkipBlanks, transpose=PasteTranspose )
          is_copy_paste = True
            
        except Exception as err:
            msg = f"функция: rng_copy_paste\nПроизошла ошибка\n{err}\n COPY wb:{rngCopy.sheet.book}, wst: {rngCopy.sheet}"
            msg += f"\nPASTE wb: {rngPasteStart.sheet.book}, wst: {rngPasteStart.sheet}, rngPasteStart: {rngPasteStart}"
            msg += f"\nPASTE paste: {PasteAs}, operation: {PasteOperation}, skip_blanks: {PasteSkipBlanks}, transpose: {PasteTranspose}"
            print(msg)
            
        finally:
            return is_copy_paste
    
    def rng_copy_paste_tpl(self, rngRowTpl,  rngPasteTable):
        '''
       
        Вставить из шаблона (строка диапазон rngRowTpl) в таблицу rngPasteTable - форматы и формулы 

        Parameters
        ----------
        rngRowTpl : Range строка шаблона (форматы формулы)
        rngPasteTable : Range
            таблица для вставки форматов и формул из rngRowTpl 

        Returns
        -------
        True если Ок, False

        '''
        is_copy_paste = False
        try:
          # rngCopy.copy() sht.range('A1').api.HasFormula
          # rngPasteStart.paste(paste=PasteAs, operation=PasteOperation, 
          #                               skip_blanks=PasteSkipBlanks, transpose=PasteTranspose )
          
          for i, rng in enumerate(rngRowTpl.rows(1).columns):              
              if rng.api.HasFormula:
                  self.rng_copy_paste(rng,  rngPasteTable.columns[i], PasteAs='formulas')
              self.rng_copy_paste(rng,  rngPasteTable.columns[i], PasteAs='formats')
          
          is_copy_paste = True
            
        except Exception as err:
            # msg = f"функция: rng_copy_paste_tpl\nПроизошла ошибка\n{err}\n COPY wb:{rngCopy.sheet.book}, wst: {rngCopy.sheet}"
            # msg += f"\nPASTE wb: {rngPasteStart.sheet.book}, wst: {rngPasteStart.sheet}, rngPasteStart: {rngPasteStart}"
            # msg += f"\nPASTE paste: {PasteAs}, operation: {PasteOperation}, skip_blanks: {PasteSkipBlanks}, transpose: {PasteTranspose}"
            # print(msg)
            pass
        finally:
            return is_copy_paste
    
      
    def df_to_rng(self, rngPasteStart, df, dfIsIndex = False, dfIsHeader = False, dfIsDropna = False ):
        '''
        Вставить df в лист Excel начиная с rngPasteStart

        Parameters
        ----------

        rngPasteNameStart : range
            начало диапазона вставки (лев край) для df
        df : pd.DataFrame
            all_merging_conditional_formats, all, all_except_borders, all_using_source_theme, 

        dfIsIndex : logical, 
            The default is False.
        dfIsHeader : logical, 
            The default is False.
        dfIsDropna : logical, 
            The default is False.
        Returns
        -------
        True если Ок, False
        

        '''
         
        is_rng_set_df = False
        try:
          rngPasteStart.options(index=dfIsIndex, header=dfIsHeader, dropna=dfIsDropna).value = df
          is_rng_set_df = True
            
        except Exception as err:
            msg = f"функция: df_to_rng\nПроизошла ошибка\n{err}\n PASTE wb:{rngPasteStart.sheet.book}, wst: {rngPasteStart.sheet}, rngPasteStart: {rngPasteStart}"
            msg += f"\nDF df: {df}, index: {dfIsIndex}, header: {dfIsHeader}, dropna: {dfIsDropna}"
            print(msg)
            
        finally:
            return is_rng_set_df  
        
    def df_to_rng_tpl(self, wb, sheetName, intRowTpl, intColPasteStart, strRangeName, strRangeNameFlt,
                      df, dfIsIndex = False, dfIsHeader = False, dfIsDropna = False, strTextInfo = "~" ):
        '''
        Назначение:
            Вставить df в лист Excel начиная с rngPasteStart, 
            отформатировать по строке шаблона и вставить формулы из строки щаблона
            удалиттьстроку шаблона
            добавить именованный диапзон для данных = strRangeName (df)
            добавить именованный диапзон для для фильтра = strRangeNameFlt 
            
        wb - раб книга
        sheetName - имя раб листа
        intRowTpl - номер строки шаблона
        intColPasteStart - стартовый столбец для вставки df
        
        '''
        wb.sheets[sheetName].activate()
        # 1) вставить df c rngNameStart=(intRowTpl,+1, intColPasteStart)
        rng_paste= self.rng_get_rng(wb, wstName=sheetName, rngNameStart=(intRowTpl+1, intColPasteStart))
        self.df_to_rng(rngPasteStart=rng_paste, df=df, dfIsIndex = dfIsIndex, 
                  dfIsHeader = dfIsHeader, dfIsDropna = dfIsDropna)
        # 2) заполнить шаблоны форматирования и фомул
        rows, cols = df.shape
        rng_tpl = self.rng_get_rng(wb, wstName=sheetName, 
            rngNameStart=(intRowTpl, intColPasteStart), rngNameEnd=(intRowTpl, intColPasteStart+cols-1))
        
        rng_paste = self.rng_get_rng(wb, wstName=sheetName, 
            rngNameStart=(intRowTpl+1, intColPasteStart), rngNameEnd=(intRowTpl + rows, intColPasteStart+cols-1))
        
        self.rng_copy_paste_tpl(rngRowTpl=rng_tpl, rngPasteTable=rng_paste)
        
        # 3) удалить строку шаблона
        str_row_del =  str(intRowTpl) + ":" + str(intRowTpl)
        print(sheetName, str_row_del)
        wb.sheets[sheetName].range(str_row_del).delete()
        
        # 4) добавить именованный диапазон strRangeName для данных 
        rng_paste = self.rng_get_rng(wb, wstName=sheetName,
                            rngNameStart=(intRowTpl, intColPasteStart),
                            rngNameEnd=(intRowTpl + rows-1, intColPasteStart+cols-1))
        
        txt_range_adress_1='=' + sheetName +'!' + rng_paste.address
        wb.sheets[sheetName].names.add(strRangeName,  txt_range_adress_1)
        txt_range_adress_1 = rng_paste.address
        txt_range_adress_1= sheetName +'!' + txt_range_adress_1 #для openxl
        # добавлено 15.03.23
        wb.sheets[sheetName].names[strRangeName].delete()
        # wb.sheets[sheetName].names.add(strRangeName, rng_paste.address.strip('"'))
        
        # 5) установить высоту строки
        str_rng_rows = str(intRowTpl) + ":" + str(intRowTpl + rows-1)
        wb.sheets[sheetName][str_rng_rows].api.RowHeight = 24
        
        # 6) определить диапазон для автофильтра и добавить именованный диапазон
        rng_paste = self.rng_get_rng(wb, wstName=sheetName,
                            rngNameStart=(intRowTpl-1, intColPasteStart),
                            rngNameEnd=(intRowTpl + rows-1, intColPasteStart+cols-1))
        txt_range_adress='=' + sheetName +'!' + rng_paste.address
        wb.sheets[sheetName].names.add(strRangeNameFlt,  txt_range_adress)
        txt_range_adress=rng_paste.address
        txt_range_adress= sheetName +'!' + txt_range_adress #для openxl
        # добавлено 15.03.23
        wb.sheets[sheetName].names[strRangeNameFlt].delete()
        # 6) активировать ячейку (1, 1)
        wb.sheets[sheetName].activate()
        wb.sheets[sheetName].range((1,1)).value = strTextInfo 
        wb.sheets[sheetName].range((1,1)).select()
              
        self.wb_save_as(wb)


        return (sheetName, [strRangeNameFlt, txt_range_adress, strRangeName, txt_range_adress_1])
        
    def rng_to_df(self,  rngForDf, header = 1, index = False, dfIsDropna = False):
        '''
        Сформировать df из диапазона rngForDf

        Parameters
        ----------

        rngForDf : range
            диапазон ячеек для формирования df

        Returns
        -------
        df: pd.DataFrame
        

        '''
         
        try:
            df = rngForDf.options(index=index, header=header, dropna=dfIsDropna).value
            
        except Exception as err:
            msg = f"функция: rng_to_df\nПроизошла ошибка\n{err}\n PASTE wb:{rngForDf.sheet.book}, wst: {rngForDf.sheet}, rngForDf: {rngForDf}"
            msg += f"\index: {index}, header: {header}, dropna: {dfIsDropna}"
            print(msg)
            
        finally:
            return df      

    def rng_filter_add(self, rngFilterTable):
        '''
       
        добавит автофильтр для таблицы rngPasteTable 

        Parameters
        ----------
        rngPasteTable : Range
            таблица для автофильтра

        Returns
        -------
        True если Ок, False

        '''  
        is_rng_filter = False
        try:
            rngFilterTable.sheet.api.Range(rngFilterTable.get_address()).AutoFilter()
            is_rng_filter = True
        except Exception as err:
            print(err)
        finally:
            return is_rng_filter
        
def opxl_wb_shets_add_autofilter(wbFullName, **sheet_range):
    '''
    добавитm в wb wbFullName автофильтр для указанных листов:диапазонов

    Parameters
    ----------
    wbFullName : TYPE
        DESCRIPTION.
    **sheet_range : TYPE
        DESCRIPTION.
    Returns
    -------
    None.
    
    '''
    import openpyxl as px
    from openpyxl.workbook.defined_name import DefinedName
    wb = px.load_workbook(wbFullName)
    for sheet_name, range_adress in sheet_range.items():
        ws = wb[sheet_name]
        # добавлено 15.03.23
        ws.auto_filter.ref = range_adress[1].split("!")[1]
        # Добавить именованные диапазоны Глобальные 
        defn = DefinedName(range_adress[0], attr_text=range_adress[1])
        # wb.defined_names[range_adress[0]] = defn
        wb.defined_names.append(defn)
        defn = DefinedName(range_adress[2], attr_text=range_adress[3])
        # wb.defined_names[range_adress[2]] = defn
        wb.defined_names.append(defn)        
        
    wb.save(wbFullName)
    wb.close()
        

# # wbname_ - функции для работы с именованным диапазонами         


# https://openpyxl.readthedocs.io/en/stable/defined_names.html

#     def wb_is_open(self, wbFullName):
#         '''
#         Parameters
#         ----------
#         wbFullName : строка, полное имя раб книги
#         Returns
#         -------
#         true - если книга открыта.
#         '''
#         if wbFullName in [wb.fullname for wb in self._xlApp.books ]:
#             return True
#         return False
    

     
 

    

    

    
#     def wb_close(self, wbFullName, bIsSaveWb = False):
#         '''
#         Закрыть раб.книгу
#         Parameters
#         ----------
#         wbFullName : строка, полное имя или имя раб книги
#         bIsSaveWb  : = True, сохранить раб книгу перед закрытием
#         Returns
#         -------
#         возвращает True, если все Ок
#         '''
        
#         try:
#             wb = self.wb_get_wb(wbFullName)
#             if wb is not None:
#                 if bIsSaveWb:
#                     wb.save();
#                 wb.close() 
#             return True
            
#         except Exception as err:
#             msg = f"функция: wb_close\nПроизошла ошибка{err} "
#             print(msg)
#         return False  
    
#     def wb_create(self, wbFullName=None):
#         '''
#         Создать новую раб.книгу
#         Parameters
#         ----------
#         wbFullName : строка, полное имя для новой раб книги

#         Returns
#         -------
#         возвращает рабочую книгу или None если книга не создана
#         '''
#         try:
#             wb = self._xlApp.Book()
#             if wbFullName is not None:
#                 wb.save(wbFullName)
#             return wb
            
#         except Exception as err:
#             msg = f"функция: wb_create\nПроизошла ошибка{err} "
#             print(msg)
#         return None 

# # wst_ - функции для работы с woksheet
#     def wst_get_wb_wst(self, wbFullName, wstName):
#         '''
#         Parameters
#         ----------
#         wbFullName : строка, полное имя раб книги или имя
#         wstName : строка,  имя листа раб книги

#         Returns
#         -------
#         возвращает кортеж (wb - раб книгу и wst - лист рабочей книги)
#         wb = None если книга не открыта, wst = None нет листа в wb

#         '''
#         wb = self.wb_get_wb(wbFullName)
#         if wb is not None:
#             wstdict = {wst_.name:wst_ for wst_ in wb.sheets}
#             wst = wstdict.get(wstName)
#             msg =  f"функция: ws_get_wst\n wb:{wb}, wst{wst} "
#             print(msg)
#             return (wb, wst)


    # def rng_copy_paste_(self, rngCopy, wst, rngPasteNameStart, PasteAs="all", PasteOperation = None, 
    #                    PasteSkipBlanks=False, PasteTranspose=False):
    #     '''
    #     Скопировать диапазон rngCopy вставить в диапазон 

    #     Parameters
    #     ----------
    #     rngCopy : Range for Copy
    #     wst: Sheet for paste.
    #     rngPasteName : strings
    #         начало диапазона вставки лев край или имя именнованного диапазона (нотация A1, кортеж(r, c)).
    #     PasteAs : strings
    #         all_merging_conditional_formats, all, all_except_borders, all_using_source_theme, 
    #         column_widths, comments, formats, formulas, formulas_and_number_formats, validation, 
    #         values, values_and_number_formats. 
    #         The default is "all".
    #     PasteOperation : strings, 
    #         add”, “divide”, “multiply”, “subtract”.
    #         The default is None.
    #     PasteSkipBlanks : logical
    #         The default is False.
    #     PasteTranspose : logical
    #         The default is False.

    #     Returns
    #     -------
    #     True если Ок, False

    #     '''
    #     is_copy_paste = False
    #     try:
    #       rngCopy.copy()
    #       wst.range(rngPasteNameStart).paste(paste=PasteAs, operation=PasteOperation, 
    #                                     skip_blanks=PasteSkipBlanks, transpose=PasteTranspose )
    #       is_copy_paste = True
            
    #     except Exception as err:
    #         msg = f"функция: rng_copy_paste\nПроизошла ошибка\n{err}\n COPY wb:{rngCopy.sheet.book}, wst: {rngCopy.sheet}"
    #         msg += f"\nPASTE wb: {wst.book}, wst: {wst.name}, rngPasteNameStart: {rngPasteNameStart}"
    #         msg += f"\nPASTE paste: {PasteAs}, operation: {PasteOperation}, skip_blanks: {PasteSkipBlanks}, transpose: {PasteTranspose}"
    #         print(msg)
            
    #     finally:
    #         return is_copy_paste
 
    # def rng_set_df(self, wst, rngPasteNameStart, df, dfIsIndex = False, dfIsHeader = False, dfIsDropna = False ):
    #     '''
    #     Вставить df в диапазон wst.range(rngPasteNameStart) 
    #     Parameters
    #     ----------
    #     wst: Sheet for paste df.
    #     rngPasteNameStart : strings
    #         начало диапазона вставки лев край или имя именнованного диапазона (нотация A1, кортеж(r, c)).
    #     df : pd.DataFrame
    #         all_merging_conditional_formats, all, all_except_borders, all_using_source_theme, 

    #     dfIsIndex : logical, 
    #         The default is False.
    #     dfIsHeader : logical, 
    #         The default is False.
    #     dfIsDropna : logical, 
    #         The default is False.
    #     Returns
    #     -------
    #     True если Ок, False
        

    #     '''
         
    #     is_rng_set_df = False
    #     try:
    #       wst.range(rngPasteNameStart).options(index=dfIsIndex, header=dfIsHeader, dropna=dfIsDropna).value = df
    #       is_rng_set_df = True
            
    #     except Exception as err:
    #         msg = f"функция: rng_set_df\nПроизошла ошибка\n{err}\n PASTE wb:{wst.book}, wst: {wst.name}, wst: {wst.name}, rng:{rngPasteNameStart}"
    #         msg += f"\nDF df: {df}, index: {dfIsIndex}, header: {dfIsHeader}, dropna: {dfIsDropna}"
    #         print(msg)
            
    #     finally:
    #         return is_rng_set_df             
