# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 14:32:00 2024

@author: a.karabedyan
"""

import os
import pandas as pd
from pathlib import Path
from config import name_account_balance_movements, sign_1c_upp, sign_1c_not_upp, new_names
from ErrorClasses import NoExcelFilesError
from ExcelFileConverter import ExcelFileConverter
from ExcelFilePreprocessor import ExcelFilePreprocessor


class IFileProcessor:
   
    def __init__(self):
        path_folder_excel_files = Path(os.getcwd())
        files = list(path_folder_excel_files.iterdir())
        self.dict_df = {}
        self.dict_df_check = {}
        self.empty_files = []
        self.excel_files = [file for file in files if (str(file).endswith('.xlsx') or str(file).endswith('.xls')) and '_СВОД_' not in str(file)]
        self.converter = ExcelFileConverter(self.excel_files)
        self.preprocessor = ExcelFilePreprocessor(self.excel_files)
        
    def process_start(self):
        if not self.excel_files:
            raise NoExcelFilesError('Нет доступных Excel файлов для обработки.')
        self.converter.save_as_xlsx_no_alert()
        self.preprocessor.preprocessor_openpyxl()
    
    def process_end(self):
        print('Закончили обработку')

class AccountTurnoverProcessor(IFileProcessor):
    def find_turnover_index(self, name_account_balance_movements, turnover_type, df):
        try:
            # Ищем первый существующий индекс
            turnover_index = next(df.columns.get_loc(i) for i in name_account_balance_movements[turnover_type] if i in df.columns)
            
            # Проверяем, совпадает ли с вторым элементом
            if turnover_index == name_account_balance_movements[turnover_type][1]:
                self.sign_1c = sign_1c_upp
            
            return turnover_index
        except StopIteration:
            return False  # Если ничего не найдено
    
    def table_header(self):
        if not self.excel_files:
            raise NoExcelFilesError('Нет доступных Excel файлов для обработки.')
        for oFile in self.excel_files:
            self.df = pd.read_excel(oFile)
            # Получаем индекс строки, содержащей target_value (значение)
            index_for_columns = None  # Инициализируем как None для удобства проверки
            target_values = {value for sublist in name_account_balance_movements.values() for value in sublist} # Извлекаем все значения
            indices = self.df.index[self.df.apply(lambda row: row.isin(target_values).any(), axis=1)]
            if not indices.empty:
                index_for_columns = indices[0]  # Получаем первый индекс
            
            # устанавливаем заголовки
            self.df.columns = self.df.iloc[index_for_columns].astype(str)
            self.df = self.df.loc[:, self.df.columns.notna()]
            # удаляем данные выше строки, содержащей имена столбцов таблицы (наименование отчета, период и т.д.)
            self.df = self.df.drop(self.df.index[0:(index_for_columns+1)])
            self.df.dropna(axis=0, how='all', inplace=True) # удаляем пустые строки
            self.df.dropna(axis=1, how='all', inplace=True)
        
            # получим наименование первого столбца, в котором находятся наши уровни
            # переименуем этот столбец
            self.df.columns.values[0] = 'Уровень'
            self.sign_1c = sign_1c_not_upp
            
            indices_to_rename = []
            
            for turnover_type in name_account_balance_movements.keys():
                index_name = f"{turnover_type}_index"
                setattr(self, f"{turnover_type}_index",
                        self.find_turnover_index(name_account_balance_movements,
                                                 turnover_type,
                                                 self.df, sign_1c_upp))
                indices_to_rename.append(index_name)
            
            if any(col in self.df.columns for col in name_account_balance_movements['debit_turnover']):
                # Определяем верхнюю границу для добавления префикса 'до'
                upper_bound_index = self.credit_turnover_index or self.end_debit_balance_index or self.end_credit_balance_index
            
                # Создаем новый список названий столбцов с префиксом 'до'
                list_do_columns = []
                for idx, col in enumerate(self.df.columns):
                    # Если нашли индекс 'дебетового оборота', добавляем префикс 'до' при выполнении условий
                    if self.debit_turnover_index is not None and idx > self.debit_turnover_index and (upper_bound_index is None or idx < upper_bound_index):
                        list_do_columns.append(f'{col}_до')
                    else:
                        list_do_columns.append(col)
            
                # Обновляем названия столбцов в DataFrame
                self.df.columns = list_do_columns
            
            if any(col in self.df.columns for col in name_account_balance_movements['credit_turnover']):
                list_ko_columns = []
                
                # Находим индекс 'КредитОборот'
                credit_turnover_index = self.df.columns.get_loc('КредитОборот') if 'КредитОборот' in self.df.columns else None
                
                # Определяем границы для добавления префикса 'ко'
                end_balances_index = max(self.end_debit_balance_index or -1, self.end_credit_balance_index or -1)  # Определяем конец диапазона
                
                for idx, col in enumerate(self.df.columns):
                    # Добавляем префикс, если индекс в нужном диапазоне
                    if credit_turnover_index is not None and idx > credit_turnover_index and (end_balances_index == -1 or idx < end_balances_index):
                        list_ko_columns.append(f'{col}_ко')
                    else:
                        list_ko_columns.append(col)
                
                self.df.columns = list_ko_columns
            
            # переименуем первые два столбца
            self.df.columns.values[0] = 'Уровень'
            #logger.info(f'{file_excel}: успешно обновили шапку таблицы, удалили строки выше шапки')
            
            # Получаем текущие имена столбцов
            current_columns = self.df.columns.tolist()
            
            # Создаем словарь с новыми именами для желаемых индексов
            rename_dict = {current_columns[i]: new_names[j] for j, i in enumerate(indices_to_rename) if i}
        
            # Переименовываем столбцы
            self.df.rename(columns=rename_dict, inplace=True)
            
            # удаляем пустые строки и столбцы
            self.df.dropna(axis=0, how='all', inplace=True)
            
        return True
