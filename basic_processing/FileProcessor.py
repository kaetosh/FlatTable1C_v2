# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 14:32:00 2024

@author: a.karabedyan
"""

import os
from typing import List, Dict, Literal
import pandas as pd
pd.options.mode.copy_on_write = False
import numpy as np
from pathlib import Path
from basic_processing.Register1C import Register1c, TableStorage
from config import new_names, osv_fields, turnover_fields, analysis_fields, exclude_values
from additional.ErrorClasses import NoExcelFilesError
from pre_processing.ExcelFileConverter import ExcelFileConverter
from pre_processing.ExcelFilePreprocessor import ExcelFilePreprocessor


name_file_register = str
accounting_account = str

class IFileProcessor:
   
    def __init__(self, file_type: Literal['account_turnover',
                                          'account_analysis',
                                          'account_osv']) -> None:
        self.pivot_table_check: pd.DataFrame = pd.DataFrame() # свод таблица с отклонениями по всем обработанным регистрам
        self.pivot_table: pd.DataFrame = pd.DataFrame() # свод таблица по всем обработанным регистрам
        self.excel_files: List[Path] =[] # список путей к обрабатываемым регистрам
        self.file_type = file_type # тип регистра (осв, анализ счета, обороты счета)
        self.dict_df: Dict[name_file_register, TableStorage] = {} # словарь со всеми обрабатываемыми регистрами и их характеристиками
        self.empty_files: List[name_file_register] = [] # список имен не обработанных по разным причинам файлов (регистров)
        self.converter: ExcelFileConverter = ExcelFileConverter() # класс для пересохранения файлов (регистров)
        self.preprocessor: ExcelFilePreprocessor = ExcelFilePreprocessor() # класс для предварительной обработки файлов (регистров)
        self.register: Register1c = self._get_fields_register() # класс с полями обрабатываемых регистров
        
    def _get_fields_register(self) -> Register1c:
        """
        На основе типа регистра (анализ счета, обороты счета или осв) возвращает
        соответствующий этому типу экземпляр класса регистра, содержащий поля
        (сальдо, обороты и т.д.), необходимые для формирования шапки обрабатываемых таблиц (регистров)
        """
        match self.file_type:
            case "account_turnover":
                return turnover_fields
            case "account_analysis":
                return analysis_fields
            case "account_osv":
                return osv_fields

    @staticmethod
    def _is_accounting_code(value: accounting_account) -> bool:
        """
        Проверяет, является ли значение бухгалтерским счетом. На вход - строка.
        """
        if value:
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
    def _fill_level(row: pd.Series,
                    prev_value: str,
                    level: int,
                    fild_register: str) -> str:
        """
        Используется для заполнения значений из верхних уровней на этапе превращения таблицы
        в плоский вид. Допустим регистр 1С имеет следующую структуру в столбце:
        счет-субсчет-контрагент-договор, эти значения расположены друг под другом. Например:
        |--------------------|
         Счет
        |--------------------|
         60
        |--------------------|
         60.01
        |--------------------|
         ООО фирма НеАйс
        |--------------------|
         Договор поставки №1
        |--------------------|
        С помощью этой функции скрипт по сути транспонирует данные из вертикальной ориентации в горизонтальную.
        Для каждой строки будут добавлены столбцы со значениями предыдущих уровней:
        |--------------------|--------------------|--------------------|--------------------|
         Счет                 Уровень 1            Уровень 2            Уровень 3
        |--------------------|--------------------|--------------------|--------------------|
         60                   60                   60                   60
        |--------------------|--------------------|--------------------|--------------------|
         60.01                60                   60                   60
        |--------------------|--------------------|--------------------|--------------------|
         ООО фирма НеАйс      60.01                60                   60
        |--------------------|--------------------|--------------------|--------------------|
         Договор поставки №1  ООО фирма НеАйс      60.01                60
        |--------------------|--------------------|--------------------|--------------------|
        Далее строки, кроме строк самого глубокого уровня (со значением Договор поставки №1) будут удалены.
        """
        if row['Уровень'] == level:
            return row[fild_register]
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
                lambda x: f'0{x}' if len(x) == 1 and self._is_accounting_code(x) else x)
            
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
                df[f'Level_{i}'] = df.apply(lambda x: self._fill_level(x, prev_value, i, register_fields.analytics), axis=1)
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
                if self.pivot_table[i].apply(self._is_accounting_code).nunique() == 2:
                    shift_level = i  # получили столбец, в котором есть и субсчет и субконто, например Level_2
                    lm = int(shift_level.split('_')[-1])  # получим его хвостик, например 2
                    # получим перечень столбцов, которые бум двигать (первый - это столбец, где есть и субсчет и субконто)
                    new_list_lev = list_lev[lm:]
                    # сдвигаем:
                    self.pivot_table[new_list_lev] = self.pivot_table.apply(
                        lambda x: pd.Series([x[i] for i in new_list_lev]) if self._is_accounting_code(
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

                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
