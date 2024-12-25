# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 14:32:00 2024

@author: a.karabedyan
"""

import os
from typing import List, Dict
import pandas as pd
from pathlib import Path
from Register1C import Register_1C
from config import name_account_balance_movements, sign_1c_upp, sign_1c_not_upp, new_names, osv_filds, turnover_filds, analisys_filds
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
                self.sign_1c = register.get_attribute_name_by_value(first_valid_value)
    
                # Устанавливаем заголовки и очищаем данные
                df.columns = df.iloc[match_index].astype(str)
                df = df.loc[:, df.columns.notna()]
                df = df.drop(df.index[0:(match_index + 1)])
                df.dropna(axis=0, how='all', inplace=True)
                df.dropna(axis=1, how='all', inplace=True)
                
                # переименуем столбцы, в которых находятся наши уровни и признаки курсива
                df.columns.values[0] = 'Уровень'
                if self.file_type == 'account_analisys' and df.iloc[:, 1].isin([0, 1]).all():
                    df.columns.values[1] = 'Курсив'
                df['Исх.файл'] = oFile.name
                
                # запишем таблицу в словарь
                self.dict_df[oFile.name] = df
            else:
                self.empty_files.append(oFile.name)
    
    def special_table_header(self) -> None:
        pass
    
    def process_end(self) -> None:
        print('Закончили обработку')

class AccountTurnoverProcessor(IFileProcessor):
    
    def find_turnover_index(self, turnover_type: str, df: pd.DataFrame) -> int or False:
        try:
            # Ищем первый существующий индекс
            turnover_index: int = next(df.columns.get_loc(i) for i in name_account_balance_movements[turnover_type] if i in df.columns)
            
            # Проверяем, является ли файл выгрузкой из 1С УПП
            if df.columns[turnover_index] == name_account_balance_movements[turnover_type][1]:
                self.sign_1c = sign_1c_upp
            return turnover_index
        except StopIteration:
            return False  # Если ничего не найдено
    
    def special_table_header(self) -> None:

        for i in self.dict_df:
            df = self.dict_df[i]
            
            indices_to_rename: List[int] = []
            
            for turnover_type in name_account_balance_movements.keys():
                index_turnover_type: int or False = self.find_turnover_index(turnover_type, df)
                setattr(self, f"{turnover_type}_index",index_turnover_type)
                indices_to_rename.append(index_turnover_type)
            
            if any(col in df.columns for col in name_account_balance_movements['debit_turnover']):
                # Определяем верхнюю границу для добавления префикса 'до'
                upper_bound_index: int = self.credit_turnover_index or self.end_debit_balance_index or self.end_credit_balance_index
            
                # Создаем новый список названий столбцов с префиксом 'до'
                list_do_columns: List[str] = []
                for idx, col in enumerate(df.columns):
                    # Если нашли индекс 'дебетового оборота', добавляем префикс 'до' при выполнении условий
                    if self.debit_turnover_index is not None and idx > self.debit_turnover_index and (upper_bound_index is None or idx < upper_bound_index):
                        list_do_columns.append(f'{col}_до')
                    else:
                        list_do_columns.append(col)
                # Обновляем названия столбцов в DataFrame
                df.columns = list_do_columns
            
            if any(col in df.columns for col in name_account_balance_movements['credit_turnover']):
                list_ko_columns: List[str] = []
                
                # Находим индекс 'КредитОборот'
                credit_turnover_index: int = df.columns.get_loc('КредитОборот') if 'КредитОборот' in df.columns else None
                
                # Определяем границы для добавления префикса 'ко'
                end_balances_index: int = max(self.end_debit_balance_index or -1, self.end_credit_balance_index or -1)  # Определяем конец диапазона
                
                for idx, col in enumerate(df.columns):
                    # Добавляем префикс, если индекс в нужном диапазоне
                    if credit_turnover_index is not None and idx > credit_turnover_index and (end_balances_index == -1 or idx < end_balances_index):
                        list_ko_columns.append(f'{col}_ко')
                    else:
                        list_ko_columns.append(col)
                # Обновляем названия столбцов в DataFrame
                df.columns = list_ko_columns
            
            # переименуем первые два столбца
            df.columns.values[0] = 'Уровень'
            
            # Получаем текущие имена столбцов
            current_columns: List[str] = df.columns.tolist()
            
            # Создаем словарь с новыми именами для желаемых индексов
            rename_dict: Dict[str, str] = {current_columns[i]: new_names[j] for j, i in enumerate(indices_to_rename) if i}
        
            # Переименовываем столбцы
            df.rename(columns=rename_dict, inplace=True)
            
            # название файла
            df['Исх.файл'] = oFile.name
            
            # запишем таблицу в словарь
            self.dict_df[oFile.name] = df
            
        # список проблемных файлов и проч удалить потом
        # for i in self.dict_df:
        #     self.dict_df[i].to_excel(f'{i}_обраб.xlsx')
        # print('empty_files', self.empty_files)

class AccountOSVProcessor(IFileProcessor):
    def table_header_OSV(self) -> None:
        if not self.excel_files:
            raise NoExcelFilesError('Нет доступных Excel файлов для обработки.')
        for oFile in self.excel_files:
            df: pd.DataFrame = pd.read_excel(oFile)
            
            # Получаем индекс строки, содержащей target_value (значение)
            index_for_columns: int or None = None
            target_values: set = {value for sublist in name_account_balance_movements.values() for value in sublist} # Извлекаем все значения
            indices: pd.core.indexes.base.Index = df.index[df.apply(lambda row: row.isin(target_values).any(), axis=1)]
            if not indices.empty:
                index_for_columns = indices[0]  # Получаем первый индекс
            else:
                self.empty_files.append(oFile.name)
                continue
            
class AccountAnalisysProcessor(IFileProcessor):
    pass
