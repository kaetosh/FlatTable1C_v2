# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 14:32:00 2024

@author: a.karabedyan
"""

import os
from typing import List, Dict
import pandas as pd
from pathlib import Path
from Register1C import Register_1C, Table_storage
from config import new_names, osv_filds, turnover_filds, analisys_filds
from ErrorClasses import NoExcelFilesError
from ExcelFileConverter import ExcelFileConverter
from ExcelFilePreprocessor import ExcelFilePreprocessor




class IFileProcessor:
   
    def __init__(self, file_type) -> None:
        path_folder_excel_files: Path = Path(os.getcwd())
        files: List[Path] = list(path_folder_excel_files.iterdir())
        self.file_type = file_type
        self.dict_df: Dict[str, pd.DataFrame] = {}
        self.dict_df_check: Dict[str, pd.DataFrame] = {}
        self.empty_files: List[str] = []
        self.excel_files: List[Path] = [file for file in files if (str(file).endswith('.xlsx') 
                                                                   or str(file).endswith('.xls'))
                                                                  and '_СВОД_' not in str(file)]
        self.converter: ExcelFileConverter = ExcelFileConverter(self.excel_files)
        self.preprocessor: ExcelFilePreprocessor = ExcelFilePreprocessor(self.excel_files, self.file_type)
        
    def get_filds_register(self):
        match self.file_type:
            case "account_turnover":
                return turnover_filds
            case "account_analisys":
                return analisys_filds
            case "account_osv":
                return osv_filds
            case _:
                raise ValueError(f"Неизвестный тип файла: {self.file_type}")
    
    def process_start(self) -> None:
        if not self.excel_files:
            raise NoExcelFilesError('Нет доступных Excel файлов для обработки.')
        self.converter.save_as_xlsx_no_alert()
        self.preprocessor.preprocessor_openpyxl()
    
    def general_table_header(self) -> None:
        if not self.excel_files:
            raise NoExcelFilesError('Нет доступных Excel файлов для обработки.')
        for oFile in self.excel_files:
            df: pd.DataFrame = pd.read_excel(oFile)
            register: Register_1C = self.get_filds_register()
            target_values: set = {i for i in register}
    
            # Найдем первый индекс совпадения и значение
            match_index = None
            first_valid_value = None
    
            for idx, row in df.iterrows():
                # Ищем первое совпадение
                matched_values = row[row.isin(target_values)]
                if not matched_values.empty:
                    match_index = idx
                    first_valid_value = matched_values.iloc[0]
                    break
    
            if match_index is not None:

                # Устанавливаем заголовки и очищаем данные
                df.columns = df.iloc[match_index]
                # df = df.loc[:, df.columns.notna()]
                # df = df.loc[:, df.columns.notna() | df.apply(lambda x: x.astype(str).str.contains(r'\d').any())]
                # df.columns = df.columns.astype(str)
                df = df.drop(df.index[0:(match_index + 1)])
                df.dropna(axis=0, how='all', inplace=True)
                df.dropna(axis=1, how='all', inplace=True)
                
                # переименуем столбцы, в которых находятся наши уровни и признаки курсива
                df.columns.values[0] = 'Уровень'
                if self.file_type == 'account_analisys' and df.iloc[:, 1].isin([0, 1]).all():
                    df.columns.values[1] = 'Курсив'
                df['Исх.файл'] = oFile.name
                
                # запишем таблицу в словарь
                sign_1c = register.get_attribute_name_by_value(first_valid_value)
                self.dict_df[oFile.name] = Table_storage(table=df, register=register, sign_1C=sign_1c)
            else:
                self.empty_files.append(oFile.name)
            
    
    def special_table_header(self) -> None:
        pass
    
    def process_end(self) -> None:
        print('Закончили обработку')

class AccountTurnoverProcessor(IFileProcessor):
        
    def special_table_header(self) -> None:
        for file in self.dict_df:
            df = self.dict_df[file].table
            sign_1c = self.dict_df[file].sign_1C
            indices_to_rename: List[int] = []
            filds_account_turnover = getattr(turnover_filds, sign_1c)
            
            df = df.loc[:, df.columns.notna()]
            df.columns = df.columns.astype(str)
            
            for turnover_type in filds_account_turnover:
                try:
                    name_atribute = turnover_filds.get_inner_attribute_by_value(turnover_type)
                    index_turnover_type: int or False = df.columns.get_loc(turnover_type) if turnover_type in df.columns else False
                    self.dict_df[file].set_index_column(name_atribute, index_turnover_type)
                    if ('debit' in name_atribute) or ('credit' in name_atribute) and index_turnover_type:
                        indices_to_rename.append(index_turnover_type)
                except TypeError:
                    continue # Пропускаем, если turnover_type == None
                except StopIteration:
                    continue  # Если ничего не найдено

            if filds_account_turnover.debit_turnover in df.columns:
                # Определяем верхнюю границу для добавления префикса 'до'
                debit_turnover_index: int = getattr(self.dict_df[file], 'index_column_debit_turnover', None)
                credit_turnover_index: int = getattr(self.dict_df[file], 'index_column_credit_turnover', None)
                end_debit_balance_index: int = getattr(self.dict_df[file], 'index_column_end_debit_balance', None)
                end_credit_balance_index: int = getattr(self.dict_df[file], 'index_column_end_credit_balance', None)
                upper_bound_index: int = credit_turnover_index or end_debit_balance_index or end_credit_balance_index
            
                # Создаем новый список названий столбцов с префиксом 'до'
                list_do_columns: List[str] = []
                for idx, col in enumerate(df.columns):
                    # Если нашли индекс 'дебетового оборота', добавляем префикс 'до' при выполнении условий
                    if debit_turnover_index is not None and idx > debit_turnover_index and (upper_bound_index is None or idx < upper_bound_index):
                        list_do_columns.append(f'{col}_до')
                    else:
                        list_do_columns.append(col)
                # Обновляем названия столбцов в DataFrame
                df.columns = list_do_columns
            
            if filds_account_turnover.credit_turnover in df.columns:
                list_ko_columns: List[str] = []
                
                # Определяем границы для добавления префикса 'ко'
                end_balances_index: int = max(end_debit_balance_index or -1, end_credit_balance_index or -1)  # Определяем конец диапазона
                
                for idx, col in enumerate(df.columns):
                    # Добавляем префикс, если индекс в нужном диапазоне
                    if credit_turnover_index is not None and idx > credit_turnover_index and (end_balances_index == -1 or idx < end_balances_index):
                        list_ko_columns.append(f'{col}_ко')
                    else:
                        list_ko_columns.append(col)
                # Обновляем названия столбцов в DataFrame
                df.columns = list_ko_columns
            
            # Получаем текущие имена столбцов
            current_columns: List[str] = df.columns.tolist()
            
            # Создаем словарь с новыми именами для желаемых индексов
            rename_dict: Dict[str, str] = {current_columns[i]: new_names[j] for j, i in enumerate(indices_to_rename) if i}
            
            print('rename_dict', rename_dict)
        
            # Переименовываем столбцы
            df.rename(columns=rename_dict, inplace=True)
            
            # запишем таблицу в словарь
            self.dict_df[file].table = df
            
        # список проблемных файлов и проч удалить потом
        # for i in self.dict_df:
        #     self.dict_df[i].table.to_excel(f'{i}_обраб.xlsx')
        # print('empty_files', self.empty_files)

class AccountOSVProcessor(IFileProcessor):
    def special_table_header(self) -> None:
        for file in self.dict_df:
            df = self.dict_df[file].table
            sign_1c = self.dict_df[file].sign_1C
            indices_to_rename: List[int] = []
            filds_osv = getattr(osv_filds, sign_1c)
            
            # for i in df.columns:
            #     match i:
            #         case (filds_osv.start_debit_balance | filds_osv.debit_turnover | filds_osv.end_debit_balance):
            #             if df.iloc[0, i] == 'Дебет':
            #                 df.rename(columns={i: f'{i}_Дебет'}, inplace=True)
            #             elif df.iloc[0, i] == 'Кредит':
            #                 df.rename(columns={i: f'{i}_Кредит'}, inplace=True)
            #             else:
            #                 raise ValueError(f"Неизвестный тип файла: {file}")
            #         case 'nan':
            #             pass
        
        
        counters = {'Дебет': 0, 'Кредит': 0}
        def update_account_list(item):
            if item in counters:
                # Увеличиваем счетчик для 'Дебет' или 'Кредит'
                counters[item] += 1
                # Возвращаем обновленное значение элемента
                return f"{item}_{['начало', 'оборот', 'конец'][counters[item] - 1]}"
            return item
        
        updated_list = [update_account_list(item) for item in df.iloc[0]]
        
        name_col = df.columns.to_list()
        # Список значений для замены
        replacement_values = ['Дебет_начало', 'Кредит_начало', 'Дебет_оборот', 'Кредит_оборот', 'Дебет_конец', 'Кредит_конец']
        
        # Замена значений в list_1
        for index, value in enumerate(updated_list):
            if value in replacement_values:
                name_col[index] = value
        
        df.columns = name_col
        
        df = df.loc[:, df.columns.notna()]
        print(df.columns)
        print()
        print(df.head(3))
        df.columns = df.columns.astype(str)
        df.to_excel('1.xlsx')
        
        # список проблемных файлов и проч удалить потом
        # for i in self.dict_df:
        #     self.dict_df[i].table.to_excel(f'{i}_обраб.xlsx')
        
        # #sign_1c = sign_1c_upp
        # list_columns_necessary = ['Уровень', sign_1c, 'Вид связи КА за период', 'Дебет_начало', 'Кредит_начало', 'Дебет_оборот', 'Кредит_оборот', 'Дебет_конец', 'Кредит_конец'] # список необходимых столбцов
        # list_columns_necessary_error = ['Уровень', sign_1c, 'Дебет_начало', 'Кредит_начало', 'Дебет_оборот', 'Кредит_оборот', 'Дебет_конец', 'Кредит_конец'] # список необходимых столбцов
        # try:
        #     df = df[list_columns_necessary].copy()
        # except KeyError as e:
        #     if 'Субконто' in e.args[0]:
        #         sign_1c = sign_1c_not_upp
        #         list_columns_necessary_error = ['Уровень', sign_1c, 'Дебет_начало', 'Кредит_начало', 'Дебет_оборот', 'Кредит_оборот', 'Дебет_конец', 'Кредит_конец'] # список необходимых столбцов
        #     df = df[list_columns_necessary_error].copy()
        #     logger_with_spinner(f'{file_excel}: ОТСУТСТВУЕТ СТОЛБЕЦ Вид связи КА за период')
            
class AccountAnalisysProcessor(IFileProcessor):
    pass
