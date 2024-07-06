# -*- coding: utf-8 -*-
"""
Created on Fri Aug 11 23:09:49 2023
# Оценка краткий отчет (общие сведения по ГК) кф 1С
Переделываем под новый отчет
@author: marks
"""

# %% import
import sys
import os
import numpy as np

import sys
import os
import numpy as np
from datetime import datetime
from calendar import monthrange
from dateutil.relativedelta import relativedelta

# %% Constans
T_SM_DRIVE = ("C", "D") # 0-"C", 1="D") ДЛЯ T_SM_DRIVE
# Кортеж поставщиков для обработки заявок на предоплату
T_VEN_ORDPAY=tuple(('ООО "Барион"','ООО "ФК "Балтимор"','ООО «Октава»', 'ООО "БионаФарм"'))
# ЩАБЛОН ФАЙЛА ДЛЯ ОТЧЕТА
FILE_TPL_NAME =  "TPL_SM_AUSALE_V06.xlsx"
# Число дней для оценки ГК просрочено или нет
DAYS_FORECAST = 15
# файл с аукционами
FILE_XLS_AU =  "ausale_data.xlsx"
SHEET_FILE_XLS_AU = "TDSheet"

# файл с заявками на оплату
FILE_XLS_PAY =  "pay_ord.xlsx"
SHEET_FILE_XLS_PAY = "TDSheet"

# %% Optionns + import SM Lib
#  вычисление числа дней в текущем месяце
_days_month = monthrange(datetime.now().year, datetime.now().month)[1]

# === !!! ЗАДАТЬ ПАРАМЕТРЫ
# 0-"C", 1="D") ДЛЯ T_SM_DRIVE
ID_DRIVE = 0
# !!! ЗАДАТЬ СДВИГ ОТ ТЕКУЩЕЙ ДАТЫ В БУДУЩЕЕ КОНТРОЛЬНУЮ ДАТУ
#DAYS_SHIFT_CONTROL = (_days_month-DAYS_FORECAST + 1)
DAYS_SHIFT_CONTROL = DAYS_FORECAST

# %%% import SM Lib
# C:\_SMLng\SMPYLib\smpd"  D:\_SMLng\SMPYLib\smpd"
smlib_path_pd = T_SM_DRIVE[ID_DRIVE]+r":\_SMLng\SMPYLib\smpd"
sys.path.insert(0, smlib_path_pd)  # путь к библиотеке
from smpd import pd  # убрать дублированик pandas
import smpd

smlib_path_smxl = T_SM_DRIVE[ID_DRIVE]+r":\_SMLng\SMPYLib\smxl\src" #C:\_SMLng\SMPYLib\smxl\src  D:\_SMLng\SMPYLib\smxl\sr
sys.path.insert(1, smlib_path_smxl) # путь к библиотеке

import smxl 

# !!! ЗАДАТЬ КОНТРОЛЬНУЮ ДАТУ
DATE_CHECK = pd.to_datetime(datetime.now().date() + relativedelta(days = DAYS_SHIFT_CONTROL))

print(DATE_CHECK)

#sys.exit()

# %%% Path for Work
# = Пути (устанавливаем путь к данным 1) =
smpy1c_path_data = "data_au" # папка с исходными данными
smpy1c_path_rpt = "rpt_au"   # папка с отчетами
#smpy1c_path = f"{os.path.split(os.getcwd())[0]}"
smpy1c_path = f"{os.getcwd()}"

sys.path.insert(2, os.path.join(smpy1c_path, smpy1c_path_data)) # путь папка с исходными данными
sys.path.insert(3, os.path.join(smpy1c_path, smpy1c_path_rpt))  # путь папка с отчетами
print(
    f"DATE_CHECK: {DATE_CHECK:%d-%m-%Y}\nsys.path[0]:  {sys.path[0]}\nsys.path[1]: {sys.path[1]}\nsys.path[2]: {sys.path[2]}\nsys.path[3]: {sys.path[3]}")

# %% LIB Function

def read_file_au(file_au=FILE_XLS_AU, file_au_sheet = SHEET_FILE_XLS_AU):
    '''
    Назначение: загрузить файл с данными из краткого отчета
    '''
    df_aus_file = os.path.join(os.path.join(sys.path[2], file_au))
    df_aus= smpd.pdxlsread(df_aus_file,file_au_sheet, header=[3,4]);
    # print(df_aus.columns)
#    превращаем названия столбцов и мультииндекса в индекс
    df_aus.columns=["_".join(col) for col in df_aus.columns.to_flat_index()]
    df_aus.columns = [ ss.split("_")[0] if "Unnamed" in ss  else ss for ss in df_aus.columns]
    return df_aus

#  функции для обработки краткого отчета
def df_aus_tidy(df_w=None):
    '''
    Назначение очистить df c кратким отчетом
    '''
    # Чистим отчет
# = 1) Заполняем Пустые строки
    list_fill = ['Ответственный','Номер аукциона', 'Номер гос.контракта','Заказчик','Грузополучатели']
    df_w[list_fill] = df_w[list_fill].fillna(method="ffill")
    df_w['Номер аукциона'] = "№" + df_w['Номер аукциона'] 
    
# = 2)  Переименовываем поля для удобства
    dict_rename = {'Номер аукциона':'Аукцион','Номер гос.контракта': 'ГК', 
                   'Регион контракта':'Регион',
                   "Информация о заявках на оплату: номер /сумма /сумма оплаты":"VEN_ORD_PAY"}
    df_w = df_w.rename(columns = dict_rename)
    
    dict_rename = {'Не отгружено всего_Количество':'Не_Отгр_УП', 'Сумма контракта':'Сумма_ГК', 'Не отгружено всего_Сумма':'Не_Отгр_СУММА','Менеджер поставщика':'Менеджер','Не отгружено на дату отчета_Количество':'Не_Отгр_УП_Дата', 'Не отгружено на дату отчета_Сумма':'Не_Отгр_СУММА_Дата'}
    df_w = df_w.rename(columns = dict_rename)
    df_w['Не_Отгр_Заявки_УП']=df_w['Не отгружено по заявкам_Количество']
    df_w['Не_Отгр_Заявки_СУММА']=df_w['Не отгружено по заявкам_Сумма']
# 3) Поля ['Не_Отгр_Заявки_УП', 'Не_Отгр_Заявки_СУММА', 'Не_Отгр_УП','Не_Отгр_СУММА'] проверяем , заполняем пустые преобразовываем
    list_fill = ['Не_Отгр_Заявки_УП', 'Не_Отгр_Заявки_СУММА', 'Не_Отгр_УП','Не_Отгр_СУММА','Привязано', 'В пути']
    df_w[list_fill] = df_w[list_fill].fillna(0)
    list_fill = ['Менеджер', 'Поставщик']
    df_w[list_fill] = df_w[list_fill].fillna("~~")
# отбираем поля 
    s1 = (df_w['Не_Отгр_Заявки_УП']).str.contains('/', na=False)
    df_w.loc[s1, 'Не_Отгр_Заявки_УП'] = ((df_w.loc[s1, 'Не_Отгр_Заявки_УП'].str.replace(' ','')).str.split('/', expand = True)).apply(pd.to_numeric).sum(axis=1) 
    s1 = (df_w['Не_Отгр_Заявки_СУММА']).str.contains('/', na=False)
    df_w.loc[s1, 'Не_Отгр_Заявки_СУММА'] = (df_w.loc[s1, 'Не_Отгр_Заявки_СУММА'].str.replace(' ','')).str.replace(',','.').str.split('/', expand = True).apply(pd.to_numeric).sum(axis=1)
    df_w[['Не_Отгр_Заявки_УП','Не_Отгр_Заявки_СУММА']] = df_w[['Не_Отгр_Заявки_УП','Не_Отгр_Заявки_СУММА']].apply(pd.to_numeric)
# = 4) По заявкам спец образом заполняем пустые поля
    col_list_gr = ['Аукцион', 'По заявкам', 'Дата подписания контракта','Дата окончания контракта', 'Дата расторжения контракта']
    list_colgrup = ['Аукцион']
    dict_colagg = {'По заявкам':'first', 'Дата подписания контракта':'first',  'Дата окончания контракта':'first',  'Дата расторжения контракта':'first'}
    dict_colrename = {'По заявкам':'GK_ORD', 'Дата подписания контракта':'Дата подписания контракта','Дата окончания контракта':'Дата окончания контракта', 'Дата расторжения контракта':'Дата расторжения контракта'}
    dfau_gkord=smpd.pdgar_mui(df_w[col_list_gr], list_colgrup, dict_colagg, dict_colrename, smdropna=False)
    s1=pd.isnull(dfau_gkord.loc[:,'GK_ORD'])
    dfau_gkord.loc[s1,'GK_ORD']="NV"
    df_w = pd.merge(df_w, dfau_gkord[['Аукцион', 'GK_ORD']], how='left', left_on =["Аукцион"], right_on =['Аукцион'])
    df_w = pd.merge(df_w, dfau_gkord, how='left', left_on =["Аукцион"], right_on =['Аукцион'], suffixes=('_DROP', '')).filter(regex='^(?!.*_DROP)')
# = 5) чистка df ставим правильные даты
    df_w['Дата поставки'] = df_w['Дата поставки'].str.split(';').str.get(0)
    df_w['Дата окончания'] = df_w['Дата окончания'].str.split(' ').str.get(0)
    #df_w['Дата поставки']= pd.to_datetime(dfausw['Дата поставки'])
    
    list_field = ['Дата окончания контракта', 'С','По', 'Дата поставки']
    list_field_new = ['GK_DATE_END', 'GK_SALE_DATE_BOF', 'GK_SALE_DATE_EOF', 'GK_SUP_DATE']
    ind = 0
    for i in list_field:
        df_w[list_field_new[ind]]= (pd.to_datetime(df_w[i], format = "%d.%m.%Y", dayfirst= True, errors = 'coerce'))   #.dt.date , errors = 'coerce'
        ind += 1

    df_w['Дата окончания']= pd.to_datetime(df_w['Дата окончания'], format = "%d.%m.%Y", dayfirst= True, errors = 'coerce')
    
#    df_w['Дата окончания']= pd.to_datetime(df_w['Дата окончания'], format = "%d.%m.%Y", dayfirst= True, errors = 'coerce')
    # df_w['Дата окончания']= (pd.to_datetime(df_w['Дата окончания'], format = "%d.%m.%Y", dayfirst= True, errors = 'coerce'))
    
    df_w['СМ_Комментарии'] = "~"
    
    return  df_w

def df_aus_check(df_w=None, date_check=None):
    '''
    назначение: сделать проверки
    date_sale - data на проверку просрочки поставки
    '''
    df_w['GK_SUP_DELAY'] = 'S0' # не просрочено
# dfausw['GK_SUP_DELAY'].astype('int16').dtypes
    s1 = (pd.notnull(df_w['GK_SALE_DATE_EOF'])) & (df_w['GK_SALE_DATE_EOF']<date_check )  & (df_w['GK_ORD']=="NV") & (df_w['Не_Отгр_УП']>0) & (df_w['Не_Отгр_УП'] >= df_w['Привязано'])
    df_w.loc[s1, 'GK_SUP_DELAY'] = 'S1' # просрочено

    s1 = (pd.notnull(df_w['GK_SALE_DATE_EOF'])) & (df_w['GK_SALE_DATE_EOF']<date_check ) & (df_w['GK_ORD']=="V") & (df_w['Не_Отгр_Заявки_УП'] >0) & (df_w['Не_Отгр_Заявки_УП'] >= df_w['Привязано'])
    df_w.loc[s1, 'GK_SUP_DELAY'] = 'S1'

# добавлено 30.05.23
    s1 = (pd.notnull(df_w['GK_SALE_DATE_BOF'])) & (df_w['GK_SALE_DATE_BOF']<date_check )  & (df_w['GK_ORD']=="NV") & (df_w['Не_Отгр_УП']>0) & (df_w['Не_Отгр_УП'] >= df_w['Привязано'])
    df_w.loc[s1, 'GK_SUP_DELAY'] = 'S1' # просрочено

    s1 = (pd.notnull(df_w['GK_SALE_DATE_BOF'])) & (df_w['GK_SALE_DATE_BOF']<date_check ) & (df_w['GK_ORD']=="V") & (df_w['Не_Отгр_Заявки_УП'] >0) & (df_w['Не_Отгр_Заявки_УП'] >= df_w['Привязано'])
    df_w.loc[s1, 'GK_SUP_DELAY'] = 'S1' 
    
# добавлено 13.11.23 проверка даты ГК
    s1 = df_w['GK_DATE_END']<date_check 
    df_w.loc[s1, 'GK_SUP_DELAY'] = 'S1' 
    
    
    df_w['TMP_CALC'] = 1 # для вычислений количеств
    
    return  df_w

def df_aus_get_ord_pay(df_w):
    '''
    Получить заяки на оплату по аукциону поставщику
    Parameters
    ----------
    df_w : DataFrame
        DESCRIPTION.

    Returns
    -------
    None.

    '''
    list_cols = ['Аукцион',  'Номенклатура', 'Этап', 'Поставщик', 'Менеджер', 'VEN_ORD_PAY' ]
    row_sel_old = ~pd.isna(df_w.loc[:,'VEN_ORD_PAY'])
    df = df_w.loc[row_sel_old,list_cols]
    # число полей VEN_ORD_PAY_ + i = числу \n + 1
    _max_n_count = np.max(df.loc[:,'VEN_ORD_PAY'].str.count('\n'))
    for i in range(_max_n_count+1):
        _field = "ORD_PAY_N" + str(i)
        df.loc[:, _field]= df.loc[:,'VEN_ORD_PAY'].str.split(pat='\n', expand = True).get(i).str.split(pat='/', expand = True).get(0)
        # df[_field].astype(str).astype(int)
        #df[_field].astype(np.int32, errors='ignore')
        if i > 0:
            df[_field] = df[_field].fillna(0)
        df[_field].astype(np.int32)
    
    # преобразовать в 1D
    
    lst_col = df.columns.to_list()
          
    id_fields = lst_col[ :len(lst_col) -_max_n_count -1]
    varfields = lst_col[ len(lst_col)-_max_n_count - 1: len(lst_col)+1 ]
    
    # print(lst_col, id_fields, varfields, sep='\n' )
    
    var_name = 'ORD_N'; value_name = 'ORD_PAY'
    df1D = smpd.pdmelt(df,id_fields, varfields, var_name, value_name)
    df1D['ORD_PAY'].astype(np.int32)
    return df, df1D

def df_aus_get_ord_pay_v01(df_w):
    '''
    Получить заяки на оплату по аукциону поставщику без препаратов
    Parameters
    ----------
    df_w : DataFrame
        DESCRIPTION.

    Returns
    -------
    None.

    '''
    list_cols = ['Аукцион',  'Этап', 'GK_SALE_DATE_EOF', 'Поставщик', 'Менеджер', 'VEN_ORD_PAY' ]
    row_sel_old = ~pd.isna(df_w.loc[:,'VEN_ORD_PAY'])
    df = df_w.loc[row_sel_old,list_cols]
    # число полей VEN_ORD_PAY_ + i = числу \n + 1
    _max_n_count = np.max(df.loc[:,'VEN_ORD_PAY'].str.count('\n'))
    for i in range(_max_n_count+1):
        _field = "ORD_PAY_N" + str(i)
        df.loc[:, _field]= df.loc[:,'VEN_ORD_PAY'].str.split(pat='\n', expand = True).get(i).str.split(pat='/', expand = True).get(0)
        # df[_field].astype(str).astype(int)
        #df[_field].astype(np.int32, errors='ignore')
        if i > 0:
            df[_field] = df[_field].fillna(0)
        # df[_field].astype(np.int32)
        # df.astype({_field: 'int32'}).dtypes
        df[_field].astype(str).astype(int)
    
    # преобразовать в 1D
    
    lst_col = df.columns.to_list()
          
    id_fields = lst_col[ :len(lst_col) -_max_n_count -1]
    varfields = lst_col[ len(lst_col)-_max_n_count - 1: len(lst_col)+1 ]
    
    # print(lst_col, id_fields, varfields, sep='\n' )
    
    var_name = 'ORD_N'; value_name = 'ORD_PAY'
    df1D = smpd.pdmelt(df,id_fields, varfields, var_name, value_name)
    df1D['ORD_PAY']=df1D['ORD_PAY'].astype(np.int32)
    # df1D.astype({'ORD_PAY': 'int32'}).dtypes
    #  Сгруппировать данные поля в df1D
    # 'Аукцион', 'Этап', 'GK_SALE_DATE_EOF', 'Поставщик', 'Менеджер', 'VEN_ORD_PAY', 'ORD_PAY', 'ORD_N']
    col_list_gr = ['Аукцион', 'Этап', 'GK_SALE_DATE_EOF', 'Поставщик', 'Менеджер', 'VEN_ORD_PAY', 'ORD_PAY', 'ORD_N']
    list_colgrup = ['Аукцион', 'Этап', 'GK_SALE_DATE_EOF', 'Поставщик', 'Менеджер', 'VEN_ORD_PAY', 'ORD_PAY']
    dict_colagg = {'ORD_N':'count'}
    dict_colrename = {'ORD_N':'ORD_N_COUNT'}
    s1 = df1D['ORD_PAY']>0
    df1D_G = smpd.pdgar_mui(df1D.loc[s1,col_list_gr], list_colgrup, dict_colagg, dict_colrename, smdropna=False) 
    
    
    return df, df1D_G

def df_aus_ord_pay_remain(df_av=None, df_ord=None):
    '''
    привязывает остаток по оплате к по аукциону поставщику по номеру заяки
    df_av  - df1D - из функции df_aus_get_ord_pay_v01
    df_ord  - заявки на оплату
    '''
    
    list_col_ord =['N_Заявки', 'Поставщик', 'Остаток оплаты']
    df_w = pd.merge(df_av, df_ord[list_col_ord], how='left', 
                    left_on =['Поставщик', 'ORD_PAY'], 
                    right_on =['Поставщик', 'N_Заявки'],
                    validate='many_to_one')
    df_w= df_w.sort_values(by=[ 'GK_SALE_DATE_EOF'], ascending=[True])
    list_col_w = ['Аукцион', 'Этап', 'GK_SALE_DATE_EOF', 'Менеджер', 'Поставщик', 'VEN_ORD_PAY',
           'ORD_PAY', 'Остаток оплаты']
    return df_w[list_col_w]

def df_aus_w_rpt(df_w):
    '''
    Возвращает DataFrame для AUSALE
        Parameters
    ----------
    df_w : DataFrame
        DESCRIPTION.

    Returns
    
    -------
   df_aus_w для отчета

    '''
    lst_v_df_aus_w=['№', 'Ответственный', 'Регион', 'Аукцион', 'ГК', 'Заказчик',
           'Грузополучатели', 'Номенклатура', 'Ед. изм. отчетов', 
           'По заявкам', 'GK_ORD',
           'Кол-во дней', 'Этап', 'С', 'По', 'GK_SALE_DATE_BOF',
           'GK_SALE_DATE_EOF', 'Информация о заявках по ГК','Не отгружено по заявкам_Количество',
           'Не отгружено по заявкам_Сумма', 'Не_Отгр_Заявки_УП', 'Не_Отгр_Заявки_СУММА',
           'Не_Отгр_УП', 'Не_Отгр_СУММА',
           'Не_Отгр_УП_Дата', 'Не_Отгр_СУММА_Дата', 'Привязано', 'В пути',
           'Дата поставки',  'GK_SUP_DATE', 'Поставщик', 'Менеджер', 'VEN_ORD_PAY',
           'Деф.', 'Дата окончания', 'Дата записи', 'Автор', 'Письмо о дефектуре',
           'Дата подписания контракта', 'Дата окончания контракта',
           'Дата расторжения контракта', 'GK_DATE_END', 'GK_SUP_DELAY','СМ_Комментарии']
    # return df_w.loc[:10,lst_v_df_aus_w]
    return df_w[lst_v_df_aus_w]
# функции для обрабоки поставщиков

def ven_get_ord_pay(df_w):
    '''
    получить свертку - поставщик, заявки на оплату на основе df из
    df_aus_get_ord_pay()
    
    !!!! Привидение к 1D - smpd.pdmelt

    Parameters
    ----------
    df_w : TYPE
        DESCRIPTION.
    Returns
    -------
    None.

    '''
    col_list_gr = ['Поставщик', 'VEN_ORD_PAY_0', 'VEN_ORD_PAY_1', 'Аукцион']
    list_colgrup = ['Поставщик', 'VEN_ORD_PAY_0', 'VEN_ORD_PAY_1']
    dict_colagg = {'Аукцион':'count'}
    dict_colrename = {'Аукцион':'AU_COUNT'}
    df = smpd.pdgar_mui(df_w.loc[:,col_list_gr], list_colgrup, dict_colagg, dict_colrename, smdropna=False)
    return  df

def df_aus_ven_pt(df_w=None):
    # получить pivot_table свертку по поставщику, менеджеру просрочено, непросрочено
    # версия 22.04.23, для листа AUS_VEN
    #
    pt_index = ['Менеджер', 'Поставщик']
    pt_values = ['TMP_CALC', 'Не_Отгр_Заявки_СУММА', 'Не_Отгр_СУММА', 'Не_Отгр_СУММА_Дата']
    pt_columns = ['GK_SUP_DELAY']
    pt_aggfunc = {'TMP_CALC':len, 'Не_Отгр_Заявки_СУММА':np.sum, 'Не_Отгр_СУММА':np.sum, 'Не_Отгр_СУММА_Дата':np.sum,}
    df_w_ven_pt = pd.pivot_table(df_w, index=pt_index, values =pt_values, 
                                 columns=pt_columns, aggfunc=pt_aggfunc, 
                                 fill_value=0).swaplevel(0, 1, axis=1)
    
    df_w_ven_pt.columns = ["_".join(col) for col in df_w_ven_pt.columns.to_flat_index()]
    df_w_ven_pt = df_w_ven_pt.reset_index()
    df_w_ven_pt["Pay"]=0.0; df_w_ven_pt["Sup_Pay"]=0.0;   df_w_ven_pt["Cmnt"]="~"; df_w_ven_pt['Отбор'] = 0;
    
    df_w_ven_pt = df_w_ven_pt.sort_values(by=["S1_Не_Отгр_СУММА_Дата", 'Поставщик'], ascending=[True, True])
    col_list = ['Менеджер', 'Поставщик', 
                'S1_TMP_CALC', 'S1_Не_Отгр_Заявки_СУММА', 'S1_Не_Отгр_СУММА', 'S1_Не_Отгр_СУММА_Дата', 
                'S0_TMP_CALC', 'S0_Не_Отгр_Заявки_СУММА' , 'S0_Не_Отгр_СУММА', 'S0_Не_Отгр_СУММА_Дата', 
                'Pay', 'Sup_Pay', 'Cmnt', 'Отбор']
    
    return  df_w_ven_pt[col_list]


def df_pay_load_tidy(filepay=None, filesheet=None):
    '''
    df - заявки на оплату поставщикам
    версия 22.04.23, для листа ORDPAY

    '''
    df = smpd.pdxlsread(filepay,filesheet, header=6);
    if len(df.filter(regex="Unnamed:").columns.to_list()):
        df.drop(df.filter(regex="Unnamed:").columns.to_list(), axis=1, inplace=True)
    assert len(df.filter(regex="Unnamed:").columns.to_list())==0, "в df есть столбцы Unname" 
    list_field = ['Дата платежа']
    for col in list_field:
        df[col]= (pd.to_datetime(df[col], format = "%d.%m.%Y", dayfirst= True, errors = 'coerce'))   #.dt.date , errors = 'coerce'
    
    dict_rename = {'Контрагент':'Поставщик', 'Номер':'N_Заявки'}
    df = df.rename(columns = dict_rename)
    df['Отбор'] = 0; df['Отбор'].astype(np.int16)
    df['Очередность'] = 0; df['Очередность'].astype(np.int16)
    df['Комментарии']='~'

    df = df.sort_values(["Поставщик", 'Дата платежа'], ascending=[True, True])
    s_row = df['Дата платежа'] >= pd.to_datetime("2022-01-01", errors="coerce")
    #df['N_Заявки'].astype(np.int32)
    df.astype({'N_Заявки': 'int32'}).dtypes
    list_col_df_pay = ['Автор', 'Дата', 'N_Заявки', 'Дата платежа', 'Поставщик',   'Состояние', 'Валюта документа', 
       'Остаток оплаты', 'Сумма документа', 'Кредитная нота', 'Согласовано к оплате', 
       'Оплачено',  'Описание', 'Отбор', 'Очередность', 'Комментарии']
    
    # дополнительно постащики из заявок на оплату
    
    col_list_gr = ['Поставщик', 'Валюта документа', 'Остаток оплаты']
    list_colgrup = ['Поставщик', 'Валюта документа']
    dict_colagg = {'Остаток оплаты':'sum'}
    dict_colrename = {'Остаток оплаты':'Сумма_задолженности'}
    df1= smpd.pdgar_mui(df.loc[:,col_list_gr], list_colgrup, dict_colagg, dict_colrename, smdropna=False)
    df1['Менеджер']='~'
    
    return df.loc[s_row, list_col_df_pay], df1

def df_aus_mr(df_w=None):
    '''
    Получить аукционы МР, ООО «Вейсфарм» ( добавлено 03.11.23)
     Parameters
    ----------
    df_w :
        
    Returns
    -------
    df_w_mr - аукционы МР

    '''
    col_list_gr = ['Аукцион', 'Регион', 'Номенклатура']
    list_colgrup = ['Аукцион', 'Регион']
    dict_colagg = {'Номенклатура':'count'}
    dict_colrename = {'Номенклатура':'MR'}
    s1 = df_w["Заказчик"].str.contains("МЕДРЕСУРС")
    df_w_mr = smpd.pdgar_mui(df_w.loc[s1,col_list_gr], list_colgrup, dict_colagg, dict_colrename, smdropna=False) 
    df_w_mr["MR"]=1
    s1 = df_w["Заказчик"].str.contains("«Вейсфарм»")
    df_w_mr = smpd.pdgar_mui(df_w.loc[s1,col_list_gr], list_colgrup, dict_colagg, dict_colrename, smdropna=False) 
    df_w_mr["MR"]=2   
    
    return  df_w_mr

def df_aus_au_ven_pt(df_w=None):
    '''
    получить pivot_table свертку по аукциону поставщику, менеджеру, номенклатура свернута, просрочено, непросрочено
    лист AUS_AU_VEN
    ''' 
    # не работает правильно pivot
    df_w_copy = df_w.copy(deep=True)
    row_sel_old = pd.isna(df_w_copy.loc[:,'VEN_ORD_PAY'])
    df_w_copy.loc[row_sel_old,'VEN_ORD_PAY']="~"
    
    pt_index = ['Аукцион', 'Регион', 'Этап', 'GK_SALE_DATE_EOF', 'Менеджер', 'Поставщик', 'VEN_ORD_PAY' ]
    pt_values = ['TMP_CALC', 'Сумма_ГК', 'Не_Отгр_Заявки_СУММА', 'Не_Отгр_СУММА', 'Не_Отгр_СУММА_Дата']
    pt_columns = ['GK_SUP_DELAY']
    pt_aggfunc = {'TMP_CALC':len, 'Сумма_ГК': np.mean, 'Не_Отгр_Заявки_СУММА':np.sum, 'Не_Отгр_СУММА':np.sum, 'Не_Отгр_СУММА_Дата':np.sum,}
    df_w_ven_pt = pd.pivot_table(df_w_copy, index=pt_index, values =pt_values, 
                                 columns=pt_columns, aggfunc=pt_aggfunc, 
                                 fill_value=0).swaplevel(0, 1, axis=1)
    
    df_w_ven_pt.columns = ["_".join(col) for col in df_w_ven_pt.columns.to_flat_index()]
    df_w_ven_pt = df_w_ven_pt.reset_index()
#    df_w_ven_pt['AU_PAY']=df_w_ven_pt['Аукцион'].str[-4:]
    df_w_ven_pt['AU_PAY']=(df_w_ven_pt['Аукцион'].str.split('_| ',expand=True)[0].str[-4:]).astype(str)
 
   # добавить свертку по номенклатуре
    dft=df_w.groupby(['Аукцион', 'Этап', 'Поставщик'])['Номенклатура']. agg( set ). reset_index(name='Номенклатура')
    df_w_ven_pt = pd.merge(df_w_ven_pt, dft, how='left', left_on =['Аукцион', 'Этап', 'Поставщик'], right_on = ['Аукцион', 'Этап', 'Поставщик'])
    df_w_ven_pt['Номенклатура'] = df_w_ven_pt['Номенклатура'].apply(', '.join).astype(str)
    
    # добавить МР
    df_mr = df_aus_mr(df_w)
    df_w_ven_pt = pd.merge(df_w_ven_pt, df_mr.loc[:,['Аукцион','MR']], how='left', left_on =['Аукцион'], right_on = ['Аукцион'])
    list_fill = ['MR']
    df_w_ven_pt[list_fill] = df_w_ven_pt[list_fill].fillna(0)
    df_w_ven_pt['Cmnt']="~"
    
    
   # отбор столбцов и сортировка
    col_list = ['Аукцион', 'Регион', 'Этап', 'GK_SALE_DATE_EOF', 'MR','Менеджер', 'Поставщик', 'Номенклатура',
                'S1_TMP_CALC', 'S1_Сумма_ГК', 
                'S1_Не_Отгр_Заявки_СУММА', 'S1_Не_Отгр_СУММА', 'S1_Не_Отгр_СУММА_Дата', 
                'S0_TMP_CALC', 'S0_Сумма_ГК', 
                'S0_Не_Отгр_Заявки_СУММА' , 'S0_Не_Отгр_СУММА', 'S0_Не_Отгр_СУММА_Дата',
                'AU_PAY', 'VEN_ORD_PAY','Cmnt']
    df_w_ven_pt=df_w_ven_pt[col_list] 
    dict_sort = {'GK_SALE_DATE_EOF':True ,'Аукцион':True, 'Поставщик':True}
    columns_sort = list(dict_sort.keys()); ascending_sort = list(dict_sort.values())
    df_w_ven_pt= df_w_ven_pt.sort_values(by=columns_sort, ascending=ascending_sort)
    
    return  df_w_ven_pt

def df_aus_au_pt(df_w=None, df_av_w = None):
    '''
    получить pivot_table свертку по аукциону   просрочено, непросрочено, поставщики свернуты
    для листа AUS_AU
    ''' 
    
    pt_index = ['Аукцион', 'Регион', 'Этап', 'GK_SALE_DATE_EOF' ]
    pt_values = ['TMP_CALC', 'Сумма_ГК', 'Не_Отгр_Заявки_СУММА', 'Не_Отгр_СУММА', 'Не_Отгр_СУММА_Дата']
    pt_columns = ['GK_SUP_DELAY']
    pt_aggfunc = {'TMP_CALC':len, 'Сумма_ГК': np.mean, 'Не_Отгр_Заявки_СУММА':np.sum, 'Не_Отгр_СУММА':np.sum, 'Не_Отгр_СУММА_Дата':np.sum,}
    df_w_ven_pt = pd.pivot_table(df_w, index=pt_index, values =pt_values, 
                                 columns=pt_columns, aggfunc=pt_aggfunc, 
                                 fill_value=0).swaplevel(0, 1, axis=1)
    
    df_w_ven_pt.columns = ["_".join(col) for col in df_w_ven_pt.columns.to_flat_index()]
    df_w_ven_pt = df_w_ven_pt.reset_index()

    
    dft=df_av_w.groupby(['Аукцион', 'Этап'])['Поставщик']. agg( set ). reset_index(name='Поставщик')
    dft['Поставщики_список'] = dft['Поставщик'].apply(', '.join).astype(str)
        
    
    df_w_ven_pt = pd.merge(df_w_ven_pt, dft, how='left', left_on =['Аукцион', 'Этап'], right_on = ['Аукцион', 'Этап'])
    df_w_ven_pt['Cmnt']='~'
    

    # добавить МР
    
    df_mr = df_aus_mr(df_w)
    df_w_ven_pt = pd.merge(df_w_ven_pt, df_mr.loc[:,['Аукцион','MR']], how='left', left_on =['Аукцион'], right_on = ['Аукцион'])
    list_fill = ['MR']
    df_w_ven_pt[list_fill] = df_w_ven_pt[list_fill].fillna(0)
    
    col_list = ['Аукцион', 'Регион', 'Этап', 'GK_SALE_DATE_EOF', 'MR',
                'S1_TMP_CALC', 'S1_Сумма_ГК', 'S1_Не_Отгр_Заявки_СУММА', 'S1_Не_Отгр_СУММА', 'S1_Не_Отгр_СУММА_Дата', 
                'S0_TMP_CALC', 'S0_Сумма_ГК', 'S0_Не_Отгр_Заявки_СУММА' , 'S0_Не_Отгр_СУММА', 'S0_Не_Отгр_СУММА_Дата',
                'Поставщики_список', 'Cmnt']
    df_w_ven_pt=df_w_ven_pt[col_list]
    dict_sort = {'GK_SALE_DATE_EOF':True ,'Аукцион':True}
    columns_sort = list(dict_sort.keys()); ascending_sort = list(dict_sort.values())
   
    df_w_ven_pt=df_w_ven_pt.sort_values(by=columns_sort, ascending=ascending_sort)
 
    return df_w_ven_pt
    # return  df_w_ven_pt

# %% РАСЧЕТ
# %%% 1) Загрузка данных
df_aus = read_file_au()
df_aus_w = df_aus.copy(deep=True) # рабочий df

# %%% 2) Чистка и проверка df_aus_w
df_aus_w = df_aus_tidy(df_w = df_aus_w)
df_aus_w = df_aus_check(df_w = df_aus_w, date_check=DATE_CHECK)
df_aus_pay, df_aus_pay_1D = df_aus_get_ord_pay(df_w=df_aus_w ) # заяки на оплату из ПЗ
df_av_pay, df_av_pay_1D = df_aus_get_ord_pay_v01(df_w=df_aus_w ) 

# %%% 3) ПОСТАЩИКИ
# df_ven_pay = ven_get_ord_pay(df_aus_pay)
df_aus_v = df_aus_ven_pt(df_w=df_aus_w)

# %%% 4 ) Заявки на оплату
file_pay = os.path.join(os.path.join(sys.path[2], FILE_XLS_PAY )) # файл с старым отчетом
df_pay, df_ordpay_ven = df_pay_load_tidy(file_pay, SHEET_FILE_XLS_PAY)

# %%% 5 ) Аукционы, Поставщики
df_aus_a_v = df_aus_au_ven_pt(df_w=df_aus_w)
df_av_pay_1D_rpt = df_aus_ord_pay_remain(df_av_pay_1D, df_pay)

# %%% 6 ) Аукцины список поставщиков
# свертка по аукционам и списку поставщиков
df_aus_a = df_aus_au_pt(df_w=df_aus_w, df_av_w=df_aus_a_v)
# sys.exit()

# %% ОТЧЕТ в XLS
# %%% 1) Константы
# 2) Константы для отчета
FILE_TPL = os.path.join(os.path.join(sys.path[3], FILE_TPL_NAME)) # файл с шаблоном
SHEET_FILE_TPL = "AUSALE" # лист с шаблоном краткий отчет
SHEET_FILE_TPL_V = "АUS_VEN" # лист с шаблоном поставщики
SHEET_FILE_TPL_A_V = "AUS_АU_VEN" # лист с шаблоном аукционы поставщики
SHEET_FILE_TPL_A = "AUS_АU" # лист с шаблоном аукционы 
SHEET_FILE_TPL_I = "AUS_ITEM" # лист с шаблоном препараты
SHEET_FILE_TPL_P = "ORDPAY" # лист с шаблоном  заявки на оплату
SHEET_FILE_TPL_AP = "AUS_PAY" # лист с шаблоном  заявки на оплату
ROW_TPL_SHEET_FILE_TPL = 6 # строка с шаблоном форматирования
FILE_RPT= os.path.join(os.path.join(sys.path[3], "SM_AUSALE_KF_RPT" + "_NV01.xlsx")) # файл с отчетом
# %%% 2) Даты
# 2.1) Сформировать даты для отчета
w_str_now_date = datetime.now().strftime("%d-%m-%Y")
w_str_date_check = DATE_CHECK.strftime("%d-%m-%Y")
# %%% 3) Формирование отчета

# 3) = СТАРТ ФОРМИРОВАНИЯ ОТЧЕТ в xls
appxl = smxl.CSMXl()
appxl.xl_open()
wb = appxl.wb_open(FILE_TPL)
appxl.wb_save_as(wb, FILE_RPT)

# %%%% 3.1 AUSALE 
# 3.1)  КРАТКИЙ ОТЧЕТ
df_rpt = df_aus_w_rpt(df_aus_w)
w_str_text_info = "Краткий отчет на " + w_str_now_date + ", проверка на " + w_str_date_check 
sheet_rng = appxl.df_to_rng_tpl(wb, SHEET_FILE_TPL, ROW_TPL_SHEET_FILE_TPL, 
            intColPasteStart=1, strRangeName='rAUSALE', strRangeNameFlt='rtAUSALE',
            df=df_rpt, strTextInfo=w_str_text_info )
dict_sheet_rng = {sheet_rng[0]: sheet_rng [1] }

# %%%% 3.2 AUS_VEN
# 3.2) Свертка по поставщикам rAUS_VEN
w_str_text_info = "Свертка по поставщикам на " + w_str_now_date  + ", проверка на " + w_str_date_check 
sheet_rng = appxl.df_to_rng_tpl(wb, SHEET_FILE_TPL_V, ROW_TPL_SHEET_FILE_TPL, 
            intColPasteStart=1, strRangeName='rAUS_VEN', strRangeNameFlt='rtAUS_VEN',
            df=df_aus_v, strTextInfo=w_str_text_info )
dict_sheet_rng[sheet_rng[0]] =  sheet_rng [1] 

# %%%% 3.3 ORDPAY
# 3.3) заявки на оплату rORDPAY
w_str_text_info = "Заявки на оплату на  " + w_str_now_date 
sheet_rng = appxl.df_to_rng_tpl(wb, SHEET_FILE_TPL_P, ROW_TPL_SHEET_FILE_TPL, 
            intColPasteStart=1, strRangeName='rORDPAY', strRangeNameFlt='rtORDPAY',
            df=df_pay, strTextInfo=w_str_text_info )
dict_sheet_rng[sheet_rng[0]] =  sheet_rng [1]

# %%%% 3.4 AUS_АU_VEN
# 3.4) Свертка по аукционам и поставщикам rAUS_АU_VEN
w_str_text_info = "Свертка по аукционам и поставщикам на   " + w_str_now_date + ", проверка на " + w_str_date_check  
sheet_rng = appxl.df_to_rng_tpl(wb, SHEET_FILE_TPL_A_V, ROW_TPL_SHEET_FILE_TPL, 
            intColPasteStart=1, strRangeName='rAUS_AU_VEN', strRangeNameFlt='rtAUS_AU_VEN',
            df=df_aus_a_v, strTextInfo=w_str_text_info )
dict_sheet_rng[sheet_rng[0]] =  sheet_rng [1] 

# %%%% 3.5 AUS_PAY
# 3.5) Аукцион, Поставщик, Заяки на оплату
w_str_text_info = "Аукцион, Поставщик, Заявки на оплату на   " + w_str_now_date + ", проверка на " + w_str_date_check 
sheet_rng = appxl.df_to_rng_tpl(wb, SHEET_FILE_TPL_AP, ROW_TPL_SHEET_FILE_TPL, 
            intColPasteStart=1, strRangeName='rAUS_PAY', strRangeNameFlt='rtAUS_PAY',
            df=df_av_pay_1D_rpt, strTextInfo=w_str_text_info )
dict_sheet_rng[sheet_rng[0]] =  sheet_rng [1] 

# %%%% 3.6 AUS_AU
# = 6.2.3) Свертка по аукционам rAUS_AU
w_str_text_info = "Свертка по аукционам на " + w_str_now_date  + ", проверка на " + w_str_now_date 
sheet_rng = appxl.df_to_rng_tpl(wb, SHEET_FILE_TPL_A, ROW_TPL_SHEET_FILE_TPL, 
            intColPasteStart=1, strRangeName='rAUS_AU', strRangeNameFlt='rtAUS_AU',
            df=df_aus_a, strTextInfo=w_str_text_info )
dict_sheet_rng[sheet_rng[0]] =  sheet_rng [1] 
# %%%% 3.99 финишные процедуры
# 3.99) 
appxl.wb_save_as(wb)
appxl.xl_quit()

# 3.100) добавляем автофилmтр в листы
smxl.opxl_wb_shets_add_autofilter(FILE_RPT, **dict_sheet_rng)
