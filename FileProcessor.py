# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 14:32:00 2024

@author: a.karabedyan
"""

import os
from typing import List, Dict
import pandas as pd
pd.options.mode.copy_on_write = False
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
import numpy as np
from pathlib import Path
from Register1C import Register1c, TableStorage
from config import new_names, osv_fields, turnover_fields, analysis_fields, exclude_values, accounts_without_subaccount
from ErrorClasses import NoExcelFilesError
from ExcelFileConverter import ExcelFileConverter
from ExcelFilePreprocessor import ExcelFilePreprocessor




class IFileProcessor:
   
    def __init__(self, file_type) -> None:
        self.pivot_table_check: pd.DataFrame = pd.DataFrame()
        self.pivot_table: pd.DataFrame = pd.DataFrame()
        self.excel_files: List[Path] =[]
        self.file_type = file_type
        self.dict_df: Dict[str, TableStorage] = {}
        self.dict_df_check: Dict[str, pd.DataFrame] = {}
        self.empty_files: List[str] = []
        self.converter: ExcelFileConverter = ExcelFileConverter()
        self.preprocessor: ExcelFilePreprocessor = ExcelFilePreprocessor()
        self.register: Register1c = self.get_fields_register()
        
    def get_fields_register(self):
        match self.file_type:
            case "account_turnover":
                return turnover_fields
            case "account_analysis":
                return analysis_fields
            case "account_osv":
                return osv_fields
            case _:
                raise ValueError(f"Неизвестный тип файла: {self.file_type}")

    @staticmethod
    def is_accounting_code(value):
        if value:
            # Проверка на значение "000"
            if str(value) == "000":
                return True
            try:
                parts = str(value).split('.')
                has_digit = any(part.isdigit() for part in parts)
                # Проверка, состоит ли каждая часть только из цифр (длиной 1 или 2) или (если длина меньше 3) только из букв
                return has_digit and all(
                    (part.isdigit() and len(part) <= 2) or (len(part) < 3 and part.isalpha()) for part in parts)
            except TypeError:
                return False
        else:
            return False

    @staticmethod
    def fill_level(row, prev_value, level, sign_1c) -> float:
        if row['Уровень'] == level:
            return row[sign_1c]
        else:
            return prev_value

    @staticmethod
    def get_path_excel_files()-> List[Path]:
        path_folder_excel_files: Path = Path(os.getcwd())
        files = list(path_folder_excel_files.iterdir())
        path_excel_files: List[Path] = [file for file in files if (str(file).endswith('.xlsx')
                                                                   or str(file).endswith('.xls'))
                                        and '_Pivot_' not in str(file)]
        if not path_excel_files:
            raise NoExcelFilesError('Нет доступных Excel файлов для обработки.')
        return path_excel_files

    def conversion_preprocessing(self) -> None:
        self.converter.save_as_xlsx_no_alert(self.get_path_excel_files())
        self.excel_files = self.get_path_excel_files()
        self.preprocessor.preprocessor_openpyxl(self.excel_files, self.file_type)
    
    # определяет родительские счета
    @staticmethod
    def get_parent_accounts(account) -> List[str]:
        parent_accounts = []
        for i in range(1, account.count('.') + 1):
            parent = '.'.join(account.split('.')[:-i])
            if parent not in parent_accounts:
                parent_accounts.append(parent)
        return parent_accounts
    
    # определяет счета, у которых нет субсчетов
    @staticmethod
    def accounting_code_without_subaccount(accounting_codes):
        accounting_codes_xx = [i[:2] for i in accounting_codes]
        count_dict = {}
        for item in accounting_codes_xx:
            if item in count_dict:
                count_dict[item] += 1
            else:
                count_dict[item] = 1
        result = [key for key, value in count_dict.items() if value == 1]
        result.append('00')
        result.append('000')
        return result
    
    # Функция для проверки того, является ли счет без субсчетов
    @staticmethod
    def is_parent(account, accounts):
        for acc in accounts:
            if acc.startswith(account + '.') and acc != account:
                return True
        return False
        
    def general_table_header(self) -> None:
        if not self.excel_files:
            print('1self.excel_files', self.excel_files)
            raise NoExcelFilesError('Нет доступных Excel файлов для обработки.')
        for oFile in self.excel_files:
            df: pd.DataFrame = pd.read_excel(oFile)
            target_values: set = {i for i in self.register}
    
            # Найдем первый индекс совпадения и значение
            match_index = 0
            first_valid_value = None
    
            for idx, row in df.iterrows():
                matched_values = row[row.isin(target_values)]
                if not matched_values.empty:
                    match_index = idx
                    first_valid_value = matched_values.iloc[0]
                    if self.register is not analysis_fields:
                        break
                    else:
                        for i in matched_values:
                            if i in [analysis_fields.upp.corresponding_account,
                                     analysis_fields.notupp.corresponding_account]:
                                first_valid_value = i
                                break
                        break

            if match_index is not None:

                # Устанавливаем заголовки и очищаем данные
                df.columns = df.iloc[match_index]
                df = df.drop(df.index[0:(match_index + 1)])
                df.dropna(axis=0, how='all', inplace=True)
                
                # переименуем столбцы, в которых находятся наши уровни и признаки курсива
                df.columns.values[0] = 'Уровень'
                if self.file_type == 'account_analysis' and df.iloc[:, 1].isin([0, 1]).all():
                    df.columns.values[1] = 'Курсив'
                df['Исх.файл'] = oFile.name
                
                # запишем таблицу в словарь
                sign_1c = self.register.get_outer_attribute_name_by_value(first_valid_value)
                self.dict_df[oFile.name] = TableStorage(table=df, register=self.register, sign_1C=sign_1c)
            
            else:
                self.empty_files.append(oFile.name)
    
    def special_table_header(self) -> None:
        for file in self.dict_df:
            df = self.dict_df[file].table
            df = df.loc[:, df.columns.notna()]
            df.columns = df.columns.astype(str)
            # запишем таблицу в словарь
            self.dict_df[file].table = df

    def handle_missing_values(self):
        for file in self.dict_df:
            df = self.dict_df[file].table
            sign_1c = self.dict_df[file].sign_1C
            register = self.dict_df[file].register
            register_fields = getattr(register, sign_1c)


            print('sign_1c', sign_1c)
            for i in register_fields:
                print(i)

    
            if register_fields.quantity in df.columns:
                mask = df[register_fields.quantity].str.contains('Кол.', na=False)
                df.loc[~mask, register_fields.analytics] = df.loc[~mask, register_fields.analytics].fillna('Не_заполнено')
                df[register_fields.analytics] = df[register_fields.analytics].ffill()
            else:
                # Проставляем значение "Количество" (для ОСВ, так как строки с количеством не обозначены)
                df[register_fields.analytics] = np.where(
                                            df[register_fields.analytics].isna() & df['Уровень'].eq(df['Уровень'].shift(1)),
                                            'Количество',
                                            df[register_fields.analytics]
                                        )

                df[register_fields.analytics] = df[register_fields.analytics].fillna('Не_заполнено')

    
            # Преобразование в строки и добавление ведущего нуля при необходимости
            df[register_fields.analytics] = df[register_fields.analytics].astype(str).apply(
                lambda x: f'0{x}' if len(x) == 1 and self.is_accounting_code(x) else x)
            
            # запишем таблицу в словарь
            self.dict_df[file].table = df
    
    def horizontal_structure(self) -> None:
        for file in self.dict_df:
            df = self.dict_df[file].table
            sign_1c = self.dict_df[file].sign_1C
            register = self.dict_df[file].register
            register_fields = getattr(register, sign_1c)
           
            # Инициализация переменной для хранения предыдущего значения
            prev_value = None
        
            # получим максимальный уровень иерархии
            max_level = df['Уровень'].max()
        
            # разнесем уровни в горизонтальную ориентацию в цикле
            for i in range(max_level + 1):
                df[f'Level_{i}'] = df.apply(lambda x: self.fill_level(x, prev_value, i, register_fields.analytics), axis=1)
                for j, row in df.iterrows():
                    previous_index = df.index[df.index.get_loc(j) - 1]
                    if row['Уровень'] == i:
                        prev_value = row[register_fields.analytics]
                        if prev_value == 'Количество':
                            prev_value = df.loc[previous_index, register_fields.analytics]
                    df.at[j, f'Level_{i}'] = prev_value

            # запишем таблицу в словарь
            self.dict_df[file].table = df
        
            
    def corr_account_col(self) -> None:
        pass

    def revolutions_before_processing(self) -> None:
        for file in self.dict_df:
            df = self.dict_df[file].table
            sign_1c = self.dict_df[file].sign_1C
            register = self.dict_df[file].register
            register_fields = getattr(register, sign_1c)

            existing_columns = [i for i in df.columns if i in new_names]

            if df[df[register_fields.analytics] == 'Итого'][existing_columns].empty:
                raise NoExcelFilesError
            else:
                df_for_check = df[df[register_fields.analytics] == 'Итого'][[register_fields.analytics] + existing_columns].copy().tail(2).iloc[[0]]

                with pd.option_context("future.no_silent_downcasting", True):
                    df_for_check.loc[:, :] = df_for_check.fillna(0).infer_objects(copy=False)
                df_for_check['Сальдо_начало_до_обработки'] = df_for_check[new_names[0]] - df_for_check[new_names[1]]
                df_for_check['Сальдо_конец_до_обработки'] = df_for_check[new_names[4]] - df_for_check[new_names[5]]
                df_for_check['Оборот_до_обработки'] = df_for_check[new_names[2]] - df_for_check[new_names[3]]

                df_for_check = df_for_check[[
                    'Сальдо_начало_до_обработки',
                    'Оборот_до_обработки',
                    'Сальдо_конец_до_обработки']].copy()
                df_for_check = df_for_check.reset_index(drop=True)

                # запишем таблицу в словарь
                self.dict_df[file].table_for_check = df_for_check

    def lines_delete(self):
        
        for file in self.dict_df:
            df = self.dict_df[file].table
            sign_1c = self.dict_df[file].sign_1C
            register = self.dict_df[file].register
            register_fields = getattr(register, sign_1c)

            df[register_fields.analytics] = df[register_fields.analytics].astype(str)
            
            # Определяем желаемый порядок столбцов
            desired_order = [
                'Дебет_начало',
                'Кредит_начало',
                'Дебет_оборот',
                'Кредит_оборот',
                'Дебет_конец',
                'Кредит_конец'
            ]
            
            # Находим столбцы, заканчивающиеся на '_до' и '_ко'
            do_columns = df.filter(regex='_до$').columns.tolist()
            ko_columns = df.filter(regex='_ко$').columns.tolist()
            
            do_columns.sort()
            ko_columns.sort()
            
            # Добавляем найденные столбцы к желаемому порядку
            desired_order.extend(do_columns)
            desired_order.append('Кредит_оборот')
            desired_order.extend(ko_columns)
            desired_order.append('Дебет_конец')
            desired_order.append('Кредит_конец')
            desired_order = [col for col in desired_order if col in df.columns]
        
            if sign_1c == 'upp' and df[register_fields.analytics].isin(['Количество']).any():
                for i in desired_order:
                    df[f'Количество_{i}'] = df[i].shift(-1)
            elif sign_1c == 'notupp' and register_fields.quantity in df.columns:
                for i in desired_order:
                    df[f'Количество_{i}'] = df[i].shift(-1)
        
            max_level = df['Уровень'].max()
            
            df = df[~df[register_fields.analytics].str.contains('Итого')]
            df = df[~df[register_fields.analytics].str.contains('Количество')]
            if register_fields.quantity in df.columns:
                df = df[~df[register_fields.quantity].str.contains('Кол.', na=False)]
                df = df.drop([register_fields.quantity], axis=1)
            
            for i in range(max_level):
                df = df[~((df['Уровень']==i) & (df[register_fields.analytics] == df[f'Level_{i}']) & (i<df['Уровень'].shift(-1)))]
        
            df = df[~df[register_fields.analytics].isin(exclude_values)].copy()
            df[register_fields.analytics] = df[register_fields.analytics].astype(str)
        
            # Список необходимых столбцов
            required_columns = ['Дебет_начало', 'Кредит_начало', 'Дебет_оборот', 'Кредит_оборот', 'Дебет_конец', 'Кредит_конец']
            
            # Отбор существующих столбцов
            existing_columns = [col for col in required_columns if col in df.columns]
            
            df = df[df[existing_columns].notna().any(axis=1)]
            df = df.rename(columns={'Счет': 'Субконто'})
            df.drop('Уровень', axis=1, inplace=True)
            
            # запишем таблицу в словарь
            self.dict_df[file].table = df

    def revolutions_after_processing(self) -> None:
        for file in self.dict_df:
            df = self.dict_df[file].table
            df_for_check = self.dict_df[file].table_for_check

            df_for_check_2 = pd.DataFrame()
            df_for_check_2['Сальдо_начало_после_обработки'] = [df['Дебет_начало'].sum() - df['Кредит_начало'].sum()]
            df_for_check_2['Оборот_после_обработки'] = [df['Дебет_оборот'].sum() - df['Кредит_оборот'].sum()]
            df_for_check_2['Сальдо_конец_после_обработки'] = [df['Дебет_конец'].sum() - df['Кредит_конец'].sum()]
            df_for_check_2 = df_for_check_2.reset_index(drop=True)

            # Объединение DataFrame с использованием внешнего соединения
            merged_df = pd.concat([df_for_check, df_for_check_2], axis=1)

            # Заполнение отсутствующих значений нулями
            merged_df = merged_df.infer_objects().fillna(0)

            # Вычисление разницы
            merged_df['Разница_сальдо_нач'] = merged_df['Сальдо_начало_до_обработки'] - merged_df[
                'Сальдо_начало_после_обработки']
            merged_df['Разница_оборот'] = merged_df['Оборот_до_обработки'] - merged_df['Оборот_после_обработки']
            merged_df['Разница_сальдо_кон'] = merged_df['Сальдо_конец_до_обработки'] - merged_df[
                'Сальдо_конец_после_обработки']

            merged_df['Разница_сальдо_нач'] = merged_df['Разница_сальдо_нач'].apply(lambda x: round(x))
            merged_df['Разница_оборот'] = merged_df['Разница_оборот'].apply(lambda x: round(x))
            merged_df['Разница_сальдо_кон'] = merged_df['Разница_сальдо_кон'].apply(lambda x: round(x))

            merged_df['Исх.файл'] = file

            # запишем таблицу в словарь
            self.dict_df[file].table_for_check = merged_df

    def joining_tables(self) -> None:
        list_tables_for_joining = [self.dict_df[i].table for i in self.dict_df]
        list_tables_check_for_joining = [self.dict_df[i].table_for_check for i in self.dict_df]
        self.pivot_table = pd.concat(list_tables_for_joining)
        self.pivot_table_check = pd.concat(list_tables_check_for_joining)
        
    def shiftable_level(self) -> None:
        for j in range(5):
            list_lev = [i for i in self.pivot_table.columns.to_list() if 'Level' in i]
            for i in list_lev:
                # если в столбце есть и субсчет и субконто, нужно выравнивать столбцы
                if self.pivot_table[i].apply(self.is_accounting_code).nunique() == 2:
                    shift_level = i  # получили столбец, в котором есть и субсчет и субконто, например Level_2
                    lm = int(shift_level.split('_')[-1])  # получим его хвостик, например 2
                    # получим перечень столбцов, которые бум двигать (первый - это столбец, где есть и субсчет и субконто)
                    new_list_lev = list_lev[lm:]
                    # сдвигаем:
                    self.pivot_table[new_list_lev] = self.pivot_table.apply(
                        lambda x: pd.Series([x[i] for i in new_list_lev]) if self.is_accounting_code(
                            x[new_list_lev[0]]) else pd.Series([x[i] for i in list_lev[lm - 1:-1]]), axis=1)
                    break
                
    def rename_columns(self) -> None:

        # Разделяем столбцы на две группы
        level_columns = [col for col in self.pivot_table.columns if 'Level_' in col]
        
        # Сортируем столбцы с Level_ по числовому значению в их названиях
        level_columns.sort(key=lambda x: int(x.split('_')[1]))
        
        new_names_for_upp = {self.register.upp.analytics: 'Аналитика',
                             self.register.upp.start_debit_balance: 'Дебет_начало',
                             self.register.upp.start_credit_balance: 'Кредит_начало',
                             self.register.upp.debit_turnover: 'Дебет_оборот',
                             self.register.upp.credit_turnover: 'Кредит_оборот',
                             self.register.upp.end_debit_balance: 'Дебет_конец',
                             self.register.upp.end_credit_balance: 'Кредит_конец'}
        new_names_for_notupp = {self.register.notupp.analytics: 'Аналитика',
                                self.register.notupp.start_debit_balance: 'Дебет_начало',
                                self.register.notupp.start_credit_balance: 'Кредит_начало',
                                self.register.notupp.debit_turnover: 'Дебет_оборот',
                                self.register.notupp.credit_turnover: 'Кредит_оборот',
                                self.register.notupp.end_debit_balance: 'Дебет_конец',
                                self.register.notupp.end_credit_balance: 'Кредит_конец'}
        
        self.pivot_table = self.pivot_table.rename(columns=new_names_for_upp, errors='ignore')
        self.pivot_table = self.pivot_table.rename(columns=new_names_for_notupp, errors='ignore')
            
        # Определяем желаемый порядок столбцов
        desired_order = [
            'Исх.файл',
            'Аналитика',
            'Дебет_начало',
            'Количество_Дебет_начало',
            'Кредит_начало',
            'Количество_Кредит_начало',
            'Дебет_оборот',
            'Количество_Дебет_оборот'
        ]
        
        # Находим столбцы, заканчивающиеся на '_до' и '_ко'
        do_columns = self.pivot_table.filter(regex='_до$').columns.tolist()
        ko_columns = self.pivot_table.filter(regex='_ко$').columns.tolist()
        
        do_columns.sort()
        ko_columns.sort()
        
        # Добавляем найденные столбцы к желаемому порядку
        desired_order.extend(do_columns)
        desired_order.append('Кредит_оборот')
        desired_order.append('Количество_Кредит_оборот')
        desired_order.extend(ko_columns)
        desired_order.append('Дебет_конец')
        desired_order.append('Количество_Дебет_конец')
        desired_order.append('Кредит_конец')
        desired_order.append('Количество_Кредит_конец')
        
        # Отбор существующих столбцов
        existing_columns = [col for col in desired_order if col in self.pivot_table.columns]
        
        # Используем reindex для сортировки DataFrame
        self.pivot_table = self.pivot_table.reindex(columns=(existing_columns + level_columns)).copy()
        
    def unloading_pivot_table(self) -> None:
        folder_path_summary_files = f"_Pivot_{self.file_type}.xlsx"
        with pd.ExcelWriter(folder_path_summary_files) as writer:
            self.pivot_table.to_excel(writer, sheet_name='Свод', index=False)
            self.pivot_table_check.to_excel(writer, sheet_name='Сверка', index=False)
            
    @staticmethod
    def process_end() -> None:
        print('Закончили обработку')

class AccountTurnoverProcessor(IFileProcessor):
        
    def special_table_header(self) -> None:
        for file in self.dict_df:
            df = self.dict_df[file].table
            sign_1c = self.dict_df[file].sign_1C
            indices_to_rename: List[int] = []
            fields_account_turnover = getattr(turnover_fields, sign_1c)
            df = df.loc[:, df.columns.notna()]
            df.columns = df.columns.astype(str)
            
            for turnover_type in fields_account_turnover:
                try:
                    name_attribute = turnover_fields.get_inner_attribute_name_by_value(turnover_type)
                    index_turnover_type: int or False = df.columns.get_loc(turnover_type) if turnover_type in df.columns else False
                    self.dict_df[file].set_index_column(name_attribute, index_turnover_type)
                    if ('debit' in name_attribute) or ('credit' in name_attribute) and index_turnover_type:
                        indices_to_rename.append(index_turnover_type)
                except TypeError:
                    continue # Пропускаем, если turnover_type == None
                except StopIteration:
                    continue  # Если ничего не найдено

            match fields_account_turnover.analytics:
                case 'Субконто':
                    end_debit_balance_index: int = 8 if fields_account_turnover.type_connection else 5
                    end_credit_balance_index: int = 9 if fields_account_turnover.type_connection else 6
                    credit_turnover_index: int = 7 if fields_account_turnover.type_connection else 4
                case 'Счет':
                    end_debit_balance_index: int = 6 if fields_account_turnover.quantity else 5
                    end_credit_balance_index: int = 7 if fields_account_turnover.quantity else 6
                    credit_turnover_index: int = 5 if fields_account_turnover.quantity else 4
                case _:
                    raise NoExcelFilesError

            if fields_account_turnover.debit_turnover in df.columns:
                # Определяем верхнюю границу для добавления префикса 'до'
                debit_turnover_index: int = getattr(self.dict_df[file], 'index_column_debit_turnover', None)
                credit_turnover_index = getattr(self.dict_df[file], 'index_column_credit_turnover', None)
                end_debit_balance_index = getattr(self.dict_df[file], 'index_column_end_debit_balance', None)
                end_credit_balance_index = getattr(self.dict_df[file], 'index_column_end_credit_balance', None)
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
            
            if fields_account_turnover.credit_turnover in df.columns:
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
        
            # Переименовываем столбцы
            df = df.rename(columns=rename_dict)
            
            # запишем таблицу в словарь
            self.dict_df[file].table = df
            

class AccountOSVProcessor(IFileProcessor):
    def special_table_header(self) -> None:
        for file in self.dict_df:
            df = self.dict_df[file].table
            
            # счетчик того, сколько столбцов Дебет и Кредит
            counters = {'Дебет': 0, 'Кредит': 0}
            
            '''
            в ОСВ наименования сальдо/оборотов и дебет/кредит в разных строках,
            поэтому добавляем к дебет/кредит 'начало', 'оборот', 'конец'
            '''
            def update_account_list(item):
                if item in counters:
                    counters[item] += 1
                    return f"{item}_{['начало', 'оборот', 'конец'][counters[item] - 1]}"
                return item
            
            # берем строку, где есть дебет/кредит (первая, сразу после шапки)
            # и дополняем к этим значениям 'начало', 'оборот', 'конец'
            updated_list = [update_account_list(item) for item in df.iloc[0]]
            name_col = df.columns.to_list()

            replacement_values = ['Дебет_начало', 'Кредит_начало', 'Дебет_оборот', 'Кредит_оборот', 'Дебет_конец', 'Кредит_конец']
            
            # обновляем шапку таблицы 
            for index, value in enumerate(updated_list):
                if value in replacement_values:
                    name_col[index] = value
            df.columns = name_col
            
            df = df.loc[:, df.columns.notna()]
            df.columns = df.columns.astype(str)
            df = df.iloc[1:]
            
            # запишем таблицу в словарь
            self.dict_df[file].table = df
        
            
class AccountAnalysisProcessor(IFileProcessor):
    def handle_missing_values(self):
        for file in self.dict_df:
            df = self.dict_df[file].table
            sign_1c = self.dict_df[file].sign_1C
            register_fields = getattr(analysis_fields, sign_1c)
            
            # сохраним столбец "Вид связи КА" в отдельный фрейм
            # чтобы в методе lines_delete проставить пропущенные значения "Вид связи КА"
            if register_fields.type_connection in df.columns:
                df_type_connection = (
                        df
                        .drop_duplicates(subset=[register_fields.analytics, register_fields.type_connection])
                        .dropna(subset=[register_fields.analytics, register_fields.type_connection])  # Удаляем строки с NaN в указанных столбцах
                        .loc[:, [register_fields.analytics, register_fields.type_connection]]
                    )
                self.dict_df[file].table_type_connection = df_type_connection

            # Проверка на пропуски и условия для заполнения
            mask = (
                df[register_fields.analytics].isna() &
                ~df[register_fields.corresponding_account].apply(self.is_accounting_code) &
                ~df[register_fields.corresponding_account].isin(['Кол-во:']) &
                df[register_fields.corresponding_account].isin(exclude_values))
            
            # Заполнение пропусков
            df[register_fields.analytics] = np.where(mask, 'Не_заполнено', df[register_fields.analytics])
            
            # Заполнение последними непустыми значениями
            df[register_fields.analytics] = df[register_fields.analytics].ffill()
            
            # Приведение к строковому типу
            df[register_fields.analytics] = df[register_fields.analytics].astype(str)
            
            # Добавление '0' к счетам до 10
            df[register_fields.analytics] = df[register_fields.analytics].apply(
                lambda x: f'0{x}' if (len(x) == 1 and self.is_accounting_code(x)) else x)

            # Запишем таблицу в словарь
            self.dict_df[file].table = df
    
    def corr_account_col(self) -> None:
        for file in self.dict_df:
            df = self.dict_df[file].table
            sign_1c = self.dict_df[file].sign_1C
            register_fields = getattr(analysis_fields, sign_1c)
        
            # добавим столбец корр.счет, взяв его из основного столбца, при условии, что значение - бухгалтерских счет (функция is_accounting_code)
            df['Корр_счет'] = df[register_fields.corresponding_account].apply(lambda x: str(x) if (self.is_accounting_code(x) or str(x) == '0') else None)
            
            # добавим нолик, если счет до 10, чтобы было 01 02 04 05 07 08 09
            df['Корр_счет'] = df['Корр_счет'].apply(lambda x: f'0{x}' if len(str(x)) == 1 else x)
            
            # добавим нолик к счетам и в основном столбце
            df['Корр_счет'] = df['Корр_счет'].apply(lambda x: f'0{x}' if len(str(x)) == 1 else x)
        
            # Заполнение пропущенных значений в столбце значениями из предыдущей строки
            df['Корр_счет'] = df['Корр_счет'].ffill()
            
            # Запишем таблицу в словарь
            self.dict_df[file].table = df

    def revolutions_before_processing(self) -> None:
        for file in self.dict_df:
            df = self.dict_df[file].table
            sign_1c = self.dict_df[file].sign_1C
            register = self.dict_df[file].register
            register_fields = getattr(register, sign_1c)
            df_for_check = df[[register_fields.corresponding_account,
                               register_fields.debit_turnover,
                               register_fields.credit_turnover]].copy()
            df_for_check['Кор.счет_ЧЕК'] = df_for_check[register_fields.corresponding_account].apply(
                lambda x: str(x) if self.is_accounting_code(x) else None).copy()
            df_for_check = df_for_check.dropna(subset=['Кор.счет_ЧЕК'])
            df_for_check['Кор.счет_ЧЕК'] = df_for_check['Кор.счет_ЧЕК'].fillna('')
            df_for_check['Кор.счет_ЧЕК'] = df_for_check['Кор.счет_ЧЕК'].astype(str)
            df_for_check['Кор.счет_ЧЕК'] = df_for_check['Кор.счет_ЧЕК'].apply(lambda x: f'0{x}' if len(x) == 1 else x)

            if '94.Н' in df_for_check['Кор.счет_ЧЕК'].values:
                df_for_check = df_for_check[
                    (df_for_check['Кор.счет_ЧЕК'] == '94.Н') |
                    (df_for_check['Кор.счет_ЧЕК'].str.match(r'^\d{2}$') &
                     ~df_for_check['Кор.счет_ЧЕК'].isin([str(x) for x in range(94, 95)]))
                    ].copy()

            else:
                df_for_check = df_for_check[df_for_check['Кор.счет_ЧЕК'].str.match(r'^(\d{2}|000)$')].copy()

            df_for_check['Кор.счет_ЧЕК'] = df_for_check['Кор.счет_ЧЕК'].replace('94.Н', '94')
            df_for_check = df_for_check.groupby('Кор.счет_ЧЕК')[[register_fields.debit_turnover,
                                                                 register_fields.credit_turnover]].sum().copy()
            df_for_check = df_for_check.reset_index()

            # запишем таблицу в словарь
            self.dict_df[file].table_for_check = df_for_check

        
    def lines_delete(self):
        
        for file in self.dict_df:
            df = self.dict_df[file].table
            sign_1c = self.dict_df[file].sign_1C
            register_fields = getattr(analysis_fields, sign_1c)
            df_delete = df[~df[register_fields.corresponding_account].isin(exclude_values)]
            df_delete = df_delete.dropna(subset=[register_fields.corresponding_account]).copy()
            df_delete = df_delete[df_delete['Курсив'] == 0][[register_fields.corresponding_account, 'Корр_счет']]
            unique_df = df_delete.drop_duplicates(subset=[register_fields.corresponding_account, 'Корр_счет'])
            unique_df = unique_df[~unique_df['Корр_счет'].isin([None])]
        
            all_acc_dict = {}
            for item in list(unique_df['Корр_счет']):
                if item in all_acc_dict:
                    all_acc_dict[item] += 1
                else:
                    all_acc_dict[item] = 1
            
            # счета с субсчетами
            acc_with_sub = [i for i in all_acc_dict if self.is_parent(i, all_acc_dict)]
        
            clean_acc = [i for i in all_acc_dict if i not in acc_with_sub]
            clean_acc = [i for i in clean_acc if all_acc_dict[i] == 1]
            del_acc = [i for i in all_acc_dict if i not in clean_acc]
            
            # список из 94 счетов, т.к основной счет 94.Н в серых 1с
            # к нему открыты субсчета 94, 94.01, 94.04
            # поэтому для серой 1с оставляем только 94.Н
            # в желтых 1с и так 94 счет без субсчетов
            acc_with_94 = [i for i in all_acc_dict if '94' in i]
            del_acc_with_94 = []
            if '94.Н' in acc_with_94:
                del_acc_with_94 = [i for i in acc_with_94 if i !='94.Н']
            del_acc = list(set(del_acc + del_acc_with_94))
            
            for i in accounts_without_subaccount:
                unwanted_subaccounts = [n for n in all_acc_dict if i in n]
                del_unwanted_subaccounts = [n for n in unwanted_subaccounts if n !=i]
                del_acc = list(set(del_acc + del_unwanted_subaccounts))
        
            for i in accounts_without_subaccount:
                if i in del_acc:
                    del_acc.remove(i)
            
            df[register_fields.corresponding_account] = df[register_fields.corresponding_account].apply(lambda x: str(x))
    
        
            values_with_quantity = False
            if (df[register_fields.corresponding_account].isin(['Кол-во:']).any()
                or register_fields.quantity in df.columns):

                df['С кред. счетов_КОЛ'] = df[register_fields.debit_turnover].shift(-1)
                df['В дебет счетов_КОЛ'] = df[register_fields.credit_turnover].shift(-1)
                if register_fields.quantity in df.columns:
                    df = df[df[register_fields.quantity] != 'Кол.'].copy()
                values_with_quantity = True
            
            # Заполняем пропущенные значения в столбце Вид_связи    
            if register_fields.type_connection in df.columns:
                merged = df.merge(self.dict_df[file].table_type_connection, on=register_fields.analytics, how='left', suffixes=('', '_B'))
                df[register_fields.type_connection] = df[register_fields.type_connection].fillna(merged[f'{register_fields.type_connection}_B'])

            df = df[
                ~df[register_fields.corresponding_account].isin(exclude_values) &  # Исключение определенных значений (Сальдо, Оборот и т.д.)
                ~df[register_fields.corresponding_account].isin(del_acc) # Исключение счетов, по которым есть расшифровка субконто (60, 60.01 и т.д.)
                ].copy()
            
            df = df[df['Курсив'] == 0].copy()
            df[register_fields.corresponding_account] = df[register_fields.corresponding_account].astype(str)
        
        
            shiftable_level = 'Level_0'
            list_level_col = [i for i in df.columns.to_list() if i.startswith('Level')]
            for i in list_level_col[::-1]:
                if all(df[i].apply(self.is_accounting_code)):
                    shiftable_level = i
                    break
                
            df['Субсчет'] = df.apply(
                lambda row: row[shiftable_level] if (str(row[shiftable_level])!= '7') else f"0{row[shiftable_level]}",
                axis=1)  # 07 без субсчетов
            df['Субсчет'] = df.apply(
                lambda row: 'Без_субсчетов' if not self.is_accounting_code(row['Субсчет']) else row['Субсчет'], axis=1)
            
            df = df.rename(columns={register_fields.corresponding_account: 'Субконто_корр_счета',
                                    register_fields.analytics: 'Аналитика',
                                    register_fields.debit_turnover: 'С кред. счетов',
                                    register_fields.credit_turnover: 'В дебет счетов'})
        
            # Указываем желаемый порядок для известных столбцов
            desired_order = ['Исх.файл',
                             'Субсчет',
                             'Аналитика',
                             'Вид связи КА за период',
                             'Корр_счет',
                             'Субконто_корр_счета',
                             'С кред. счетов',
                             'В дебет счетов']
            
            if values_with_quantity:
                desired_order = ['Исх.файл',
                                'Субсчет',
                                'Аналитика',
                                'Вид связи КА за период',
                                'Корр_счет',
                                'Субконто_корр_счета',
                                'С кред. счетов',
                                'С кред. счетов_КОЛ',
                                'В дебет счетов',
                                'В дебет счетов_КОЛ']
        
            desired_order = [item for item in desired_order if item in df.columns.to_list()]
            
            # Находим все столбцы, содержащие 'Level_'
            level_columns = [col for col in df.columns.to_list() if 'Level_' in col]
        
            # Объединяем известные столбцы с найденными столбцами 'Level_'
            new_order = desired_order + level_columns
        
            # Переупорядочиваем столбцы в DataFrame
            df = df[new_order]
            
            df.loc[:, 'Субконто_корр_счета'] = df['Субконто_корр_счета'].apply(
                lambda x: 'Не расшифровано' if self.is_accounting_code(x) else x)
        
            df = df.dropna(subset=['С кред. счетов', 'В дебет счетов'], how='all')
            df = df[(df['С кред. счетов'] != 0) | (df['С кред. счетов'].notna())]
            df = df[(df['В дебет счетов'] != 0) | (df['В дебет счетов'].notna())]
            
            # Запишем таблицу в словарь
            self.dict_df[file].table = df

    def revolutions_after_processing(self) -> None:
        for file in self.dict_df:
            df = self.dict_df[file].table
            df_for_check = self.dict_df[file].table_for_check
            df_for_check_2 = df[['Корр_счет',
                                 'С кред. счетов',
                                 'В дебет счетов']].copy()
            df_for_check_2['Корр_счет'] = df_for_check_2['Корр_счет'].astype(str).copy()
            df_for_check_2['Кор.счет_ЧЕК'] = df_for_check_2['Корр_счет'].apply(
                lambda x: str(x[:2]) if (len(x) >= 2 and x != '000') else str(x)).copy()
            df_for_check_2 = df_for_check_2.groupby('Кор.счет_ЧЕК')[['С кред. счетов',
                                                                     'В дебет счетов']].sum().copy()
            df_for_check_2 = df_for_check_2.reset_index()

            # Объединение DataFrame с использованием внешнего соединения
            merged_df = df_for_check.merge(df_for_check_2, on='Кор.счет_ЧЕК', how='outer',
                                           suffixes=('_df_for_check', '_df_for_check_2'))

            # Заполнение отсутствующих значений нулями
            merged_df = merged_df.infer_objects().fillna(0)

            turnover_deb = merged_df['С кред. счетов_df_for_check'] if 'С кред. счетов_df_for_check' in merged_df.columns else 0
            turnover_cre = merged_df['В дебет счетов_df_for_check'] if 'В дебет счетов_df_for_check' in merged_df.columns else 0

            turnover_deb_2 = merged_df[
                'С кред. счетов_df_for_check_2'] if 'С кред. счетов_df_for_check_2' in merged_df.columns else 0
            turnover_cre_2 = merged_df[
                'В дебет счетов_df_for_check_2'] if 'В дебет счетов_df_for_check_2' in merged_df.columns else 0

            # Вычисление разницы
            merged_df['Разница_С_кред'] = turnover_deb - turnover_deb_2
            merged_df['Разница_В_дебет'] = turnover_cre - turnover_cre_2
            merged_df['Разница_С_кред'] = merged_df['Разница_С_кред'].apply(lambda x: round(x))
            merged_df['Разница_В_дебет'] = merged_df['Разница_В_дебет'].apply(lambda x: round(x))
            merged_df['Исх.файл'] = file
            # запишем таблицу в словарь
            self.dict_df[file].table_for_check = merged_df

    def rename_columns(self) -> None:
        list_lev = [i for i in self.pivot_table.columns.to_list() if 'Level' in i]
        for n in list_lev[::-1]:
            if all(self.pivot_table[n].apply(self.is_accounting_code)):
                self.pivot_table['Субсчет'] = self.pivot_table[n].copy()
                break
        
        
        for p in list_lev:
            if not all(self.pivot_table[p].apply(self.is_accounting_code)):
                self.pivot_table['Аналитика'] = self.pivot_table['Аналитика'].where(self.pivot_table['Аналитика']!= 'Не_указано', self.pivot_table[p])
                break
     
        self.pivot_table['Субсчет'] = self.pivot_table.apply(
            lambda row: row['Субсчет'] if (str(row['Субсчет'])!= '7') else f"0{row['Субсчет']}",
            axis=1)
        

                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
