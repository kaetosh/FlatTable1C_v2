# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 14:32:00 2024

@author: a.karabedyan
"""

import os
from typing import List, Dict, Literal, Tuple
import pandas as pd
pd.options.mode.copy_on_write = False
import numpy as np
from pathlib import Path
from basic_processing.Register1C import Register1c, FieldsRegister, TableStorage
from config import osv_fields, turnover_fields, analysis_fields, exclude_values
from additional.ErrorClasses import NoExcelFilesError
from pre_processing.ExcelFileConverter import ExcelFileConverter
from pre_processing.ExcelFilePreprocessor import ExcelFilePreprocessor


name_file_register = str
name_file_table = str
accounting_account = str
table = pd.DataFrame
table_for_check = pd.DataFrame
table_type_connection = pd.DataFrame

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
            case _:
                raise NoExcelFilesError('Нет доступных Excel файлов для обработки.')

    @staticmethod
    def _get_data_from_table_storage(file: name_file_table,
                                     dict_df: [name_file_table, TableStorage]) -> Tuple[table,
                                                                                        Literal['upp', 'notupp'],
                                                                                        Register1c,
                                                                                        FieldsRegister,
                                                                                        table_type_connection,
                                                                                        table_for_check]:
        df = dict_df[file].table
        sign_1c = dict_df[file].sign_1C
        register = dict_df[file].register
        register_fields = getattr(register, sign_1c)
        connection_type = dict_df[file].table_type_connection
        check_table = dict_df[file].table_for_check
        return df, sign_1c, register, register_fields, connection_type, check_table

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
                    fild_register: Literal["Субконто", "Счет", "Кор.счет", "Кор. Счет"]) -> str: # значения экземпляров FieldsRegister.version_1c_id
        """
        Вспомогательная функция для метода horizontal_structure, который заполняет
        значения из верхних уровней на этапе превращения таблицы в плоский вид.
        """
        if row['Уровень'] == level:
            return row[fild_register]
        else:
            return prev_value

    @staticmethod
    def _get_path_excel_files()-> List[Path]:
        """
        Так как скрипт должен запускаться из папки с исходными файлами,
        функция получает список путей к файлам из этой папки
        """
        path_folder_excel_files: Path = Path(os.getcwd())
        files = list(path_folder_excel_files.iterdir())
        path_excel_files: List[Path] = [file for file in files if (str(file).endswith('.xlsx')
                                                                   or str(file).endswith('.xls'))
                                        and '_Pivot_' not in str(file)]
        if not path_excel_files:
            raise NoExcelFilesError('Нет доступных Excel файлов для обработки.')
        return path_excel_files

    @staticmethod
    def _is_parent(account: accounting_account, accounts: List[accounting_account]) -> bool:
        """
        Функция для проверки есть ли субсчета у счета.
        """
        for acc in accounts:
            if acc.startswith(account + '.') and acc != account:
                return True
        return False

    def conversion(self) -> None:
        """
        Пересохранение в актуальный формат .xlsx
        """
        self.excel_files = self._get_path_excel_files()
        self.converter.save_as_xlsx_no_alert(self._get_path_excel_files())

    def preprocessing(self) -> None:
        """
        Предварительная обработка файлов:
            - снятие объединения ячеек
            - добавления столбца с номерами группировок строк (используется для создания плоской таблицы)
            - добавление столбца с признаком курсивного шрифта (актуально для анализа счета в УПП, строки с курсивом
            это промежуточные итоги, для исключения в сводном файле)
        Обновление атрибута self.excel_files так как .xls пересохранены в .xlsx
        """
        self.excel_files = self._get_path_excel_files()
        self.preprocessor.preprocessor_openpyxl(self.excel_files, self.file_type)

    def general_table_header(self) -> None:
        """
        Функция в цикле проходит по каждому файлу, загружает его в pandas. DataFrame.
        Далее находит строку, в которой есть совпадение с хоть одним значением
        из self.register (класс FieldsRegister регистра 1С, содержащий названия полей)
        и использует эту строку как шапку таблицы.
        """
        if not self.excel_files:
            raise NoExcelFilesError('Нет доступных Excel файлов для обработки.')
        for oFile in self.excel_files:
            df = pd.read_excel(oFile)

            # Перечень полей регистра 1С
            target_values = {i.version_1c_id for i in [self.register.upp, self.register.notupp]}

            # Найдем первый индекс совпадения и значение
            match_index = 0
            first_valid_value = None

            for idx, row in df.iterrows():
                matched_values = row[row.isin(target_values)]
                if not matched_values.empty:
                    match_index = idx
                    first_valid_value = matched_values.iloc[0]
                    break

            if match_index is not None:

                # Устанавливаем заголовки и очищаем данные
                df.columns = df.iloc[match_index]
                df = df.drop(df.index[0:(match_index + 1)])
                df.dropna(axis=0, how='all', inplace=True)

                # переименуем столбцы, в которых находятся уровни и признак курсива
                df.columns.values[0] = 'Уровень'
                df.columns.values[1] = 'Курсив'
                
                sign_1c = self.register.get_outer_attribute_name_by_value(first_valid_value)
                register_fields = getattr(self.register, sign_1c)
                # Столбец с названием файла (названием компании)
                df[register_fields.file_name] = oFile.name
                # запишем таблицу в словарь
                self.dict_df[oFile.name] = TableStorage(table=df, register=self.register, sign_1C=sign_1c)

            else:
                # Названия пустых или проблемных файлов сохраним отдельно
                self.empty_files.append(oFile.name)

    def special_table_header(self) -> None:
        """
        Метод переопределяется в классах для ОСВ и Обороты счета, так как имена их столбцов отличаются от Анализа счета
        Метод удаляет столбцы с пустыми именами и приводит имена к строковому формату.
        Шапка таблицы дополняется в зависимости от типа регистра 1С. (в переопределенных методах)
        """
        for file in self.dict_df:
            df, *_ = self._get_data_from_table_storage(file, self.dict_df)
            df = df.loc[:, df.columns.notna()]
            df.columns = df.columns.astype(str)
            # запишем таблицу в словарь
            self.dict_df[file].table = df

    def handle_missing_values(self):
        """
        Выгруженные регистры зачастую могут содержать пропущенные значения,
        обычно Вид. Субконто, например для некоторых значений счета учета запасов могут не содержать значение
        "Вид номенклатуры", счета учета расчетов - Вид контрагента или Вид договора.
        Такие пропущенные значения заполняются "Не_заполнено".
        """
        for file in self.dict_df:
            df, sign_1c, register, register_fields, *_ = self._get_data_from_table_storage(file, self.dict_df)

            # Для выгрузок с полем "Количество"
            if register_fields.quantity in df.columns:
                mask = df[register_fields.quantity].str.contains('Кол.', na=False)
                df.loc[~mask, register_fields.analytics] = df.loc[~mask, register_fields.analytics].fillna('Не_заполнено')
                df[register_fields.analytics] = df[register_fields.analytics].ffill()
            else:
                # Проставляем значение "Количество" (для ОСВ, так как строки с количеством не обозначены)
                df[register_fields.analytics] = np.where(
                                            df[register_fields.analytics].isna() & df['Уровень'].eq(df['Уровень'].shift(1)),
                                            'Количество',
                                            df[register_fields.analytics])

                df[register_fields.analytics] = df[register_fields.analytics].fillna('Не_заполнено')

            # Преобразование в строки и добавление ведущего нуля для счетов до 10 (01, 02 и т.д.)
            df[register_fields.analytics] = df[register_fields.analytics].astype(str).apply(
                lambda x: f'0{x}' if len(x) == 1 and self._is_accounting_code(x) else x)

            # запишем таблицу в словарь
            self.dict_df[file].table = df

    def horizontal_structure(self) -> None:
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
        Далее строки, кроме строк самого глубокого уровня (со значением
        Договор поставки №1) будут удалены (метод lines_delete).
        """
        for file in self.dict_df:
            df, sign_1c, register, register_fields, *_ = self._get_data_from_table_storage(file, self.dict_df)

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
        """
        Сохраняет в хранилище таблиц данные по оборотам до обработки, чтобы в дальнейшем
        сравнить их с данными по оборотам после обработки с целью убедиться в корректности обработки
        """
        for file in self.dict_df:
            df, sign_1c, register, register_fields, *_ = self._get_data_from_table_storage(file, self.dict_df)
            existing_columns = [i for i in df.columns if i in register_fields.get_attributes_by_suffix('_for_rename')]

            if df[df[register_fields.version_1c_id] == 'Итого'][existing_columns].empty:
                raise NoExcelFilesError
            else:
                df_for_check = df[df[register_fields.version_1c_id] == 'Итого'][[register_fields.version_1c_id] + existing_columns].copy().tail(2).iloc[[0]]
                df_for_check[existing_columns] = df_for_check[existing_columns].astype(float).fillna(0)

                df_for_check[register_fields.start_balance_before_processing] = (df_for_check[register_fields.start_debit_balance_for_rename]
                                                              - df_for_check[register_fields.start_credit_balance_for_rename])

                df_for_check[register_fields.end_balance_before_processing] = (df_for_check[register_fields.end_debit_balance_for_rename]
                                                             - df_for_check[register_fields.end_credit_balance_for_rename])

                df_for_check[register_fields.turnover_before_processing] = (df_for_check[register_fields.debit_turnover_for_rename]
                                                       - df_for_check[register_fields.credit_turnover_for_rename])

                df_for_check = df_for_check[register_fields.get_attributes_by_suffix('_before_processing')].reset_index(drop=True)

                # запишем таблицу в словарь
                self.dict_df[file].table_for_check = df_for_check

    def lines_delete(self) -> None:
        """
        Метод для ОСВ и Оборотов счета. Для анализа сета метод переопределен.
        После разнесения строк в плоский вид, в таблице остаются строки с дублирующими оборотами.
        Например, итоговые обороты, итоги по субконто и т.д.
        Метод их удаляет.
        """
        for file in self.dict_df:
            df, sign_1c, register, register_fields, *_ = self._get_data_from_table_storage(file, self.dict_df)

            # Получим список столбцов с сальдо и оборотами и оставим только те, которые есть в таблице
            desired_order = [col for col in register_fields.get_attributes_by_suffix('_for_rename') if col in df.columns]

            # Находим столбцы в таблице, заканчивающиеся на '_до' и '_ко'
            do_ko_columns = df.filter(regex='(_до|_ко)$').columns.tolist()
            # do_columns = df.filter(regex='_до$').columns.tolist().sort()
            # ko_columns = df.filter(regex='_ко$').columns.tolist().sort()
            # Добавим столбцы, заканчивающиеся на '_до' и '_ко', в таблицу
            if do_ko_columns:
                desired_order += do_ko_columns

            # try:
            #     # Получим индекс столбца Дебет_оборот и вставим после него столбцы с деб. оборотами счетов (для Оборотов счета)
            #     if do_columns:
            #         ind_after_deb_turnover = desired_order.index(register_fields.debit_turnover_for_rename) + 1
            #         desired_order[ind_after_deb_turnover:ind_after_deb_turnover] = do_columns
            #
            #     # Получим индекс столбца Кредит_оборот и вставим после него столбцы с кр. оборотами счетов (для Оборотов счета)
            #     if ko_columns:
            #         ind_after_cre_turnover = desired_order.index(register_fields.credit_turnover_for_rename) + 1
            #         desired_order[ind_after_cre_turnover:ind_after_cre_turnover] = ko_columns
            # except ValueError:
            #     raise NoExcelFilesError('Нет доступных Excel файлов для обработки.')

            # Если таблица с количественными данными, дополним ее столбцами с количеством путем
            # сдвига соответствующего столбца на строку вверх, так как строки с количеством чередуются с денежными значениями
            if df[register_fields.analytics].isin(['Количество']).any() or register_fields.quantity in df.columns:
                for i in desired_order:
                    df[f'Количество_{i}'] = df[i].shift(-1)

            # Удалим строки с итоговыми значениями и количественными значениями (строки с кол-вом мы разнесли в столбцы)
            # df = df[~df[register_fields.analytics].str.contains('Итого|Количество')]
            if register_fields.quantity in df.columns:
                df = df[~df[register_fields.quantity].str.contains('Кол.', na=False)]
                df = df.drop([register_fields.quantity], axis=1)
            for i in range(df['Уровень'].max()):
                df = df[~((df['Уровень']==i) & (df[register_fields.analytics] == df[f'Level_{i}']) & (i<df['Уровень'].shift(-1)))]

            # Удаляем строки, содержащие значения из списка exclude_values
            df = df[~df[register_fields.analytics].isin(exclude_values)].copy()
            
            # УТОЧНИТЬ, НЕТ ЛИ ЭТОЙ ОПЕРАЦИИ НА ЭТАПЕ СПЕЦЗАГОЛОВКОВ
            # df = df.rename(columns={'Счет': 'Субконто'})
            df.drop('Уровень', axis=1, inplace=True)

            # отберем только те строки, в которых хотя бы в одном из столбцов, определенных в existing_columns, есть непропущенные значения (не NaN)
            df = df[df[desired_order].notna().any(axis=1)]

            # запишем таблицу в словарь
            self.dict_df[file].table = df

    def revolutions_after_processing(self) -> None:
        """
        Добавляет к таблице с оборотами до обработки, созданной методом revolutions_before_processing,
        данные по оборотам после обработки и отклонениями между ними.
        """
        for file in self.dict_df:
            df, sign_1c, register, register_fields, *_, df_check_before_process = self._get_data_from_table_storage(file, self.dict_df)
            
            # Вычисление итоговых значений - свернутые значения сальдо и оборотов - обработанных таблиц
            df_check_after_process = pd.DataFrame({
                register_fields.start_balance_after_processing: [df[register_fields.start_debit_balance_for_rename].sum() - df[register_fields.start_credit_balance_for_rename].sum()],
                register_fields.turnover_after_processing: [df[register_fields.debit_turnover_for_rename].sum() - df[register_fields.credit_turnover_for_rename].sum()],
                register_fields.end_balance_after_processing: [df[register_fields.end_debit_balance_for_rename].sum() - df[register_fields.end_credit_balance_for_rename].sum()]
            })
    
            # Объединение таблиц - обороты до и после обработки таблиц
            pivot_df_check = pd.concat([df_check_before_process, df_check_after_process], axis=1).fillna(0)
    
            # Вычисление отклонений в данных до и после обработки таблиц
            for field in register_fields.get_attributes_by_suffix('_deviation'):
                pivot_df_check[field] = (pivot_df_check[field.replace('_разница', '_до_обработки')] -
                                          pivot_df_check[field.replace('_разница', '_после_обработки')]).round()
    
            # Помечаем данные именем файла
            pivot_df_check[register_fields.file_name] = file
    
            # Запись таблицы в хранилище таблиц
            self.dict_df[file].table_for_check = pivot_df_check

    def joining_tables(self) -> None:
        """
        Объединяет все таблицы друг под другом.
        """
        list_tables_for_joining = [self.dict_df[i].table for i in self.dict_df]
        list_tables_for_joining = [df.reset_index(drop=True) for df in list_tables_for_joining]

        list_tables_check_for_joining = [self.dict_df[i].table_for_check for i in self.dict_df]
        list_tables_check_for_joining = [df.reset_index(drop=True) for df in list_tables_check_for_joining]

        self.pivot_table = pd.concat(list_tables_for_joining, ignore_index=True)
        self.pivot_table_check = pd.concat(list_tables_check_for_joining, ignore_index=True)

    def shiftable_level(self) -> None:
        """
        Выравнивает столбцы таким образом, чтобы бухгалтерские счета находились в одном столбце.
        Сортировка столбцов в нужном порядке.

        """
        for j in range(5):
            list_lev = [i for i in self.pivot_table.columns.to_list() if 'Level' in i]
            for i in list_lev:
                # если в столбце есть и субсчет и субконто, нужно выравнивать столбцы
                if self.pivot_table[i].apply(self._is_accounting_code).nunique() == 2:
                    lm = int(i.split('_')[-1])  # получим его хвостик столбца, в котором есть и субсчет и субконто, например Level_2, значит 2
                    # получим перечень столбцов, которые бум двигать (первый - это столбец, где есть и субсчет и субконто)
                    new_list_lev = list_lev[lm:]
                    # сдвигаем:
                    self.pivot_table[new_list_lev] = self.pivot_table.apply(
                        lambda x: pd.Series([x[i] for i in new_list_lev]) if self._is_accounting_code(
                            x[new_list_lev[0]]) else pd.Series([x[i] for i in list_lev[lm - 1:-1]]), axis=1)
                    break

        # Получим список столбцов с сальдо и оборотами
        register_fields = self._get_fields_register().upp # upp или notupp без разницы, поля одинаковые
        desired_order = register_fields.get_attributes_by_suffix('_for_rename')

        # Находим столбцы в таблице, заканчивающиеся на '_до' и '_ко'
        do_columns = sorted(self.pivot_table.filter(regex='_до$').columns.tolist())
        ko_columns = sorted(self.pivot_table.filter(regex='_ко$').columns.tolist())

        # Функция для вставки столбцов
        def insert_columns(columns, reference_column):
            if columns:
                try:
                    ind_after = desired_order.index(reference_column) + 1
                    desired_order[ind_after:ind_after] = columns
                except ValueError:
                    raise NoExcelFilesError(f'Column {reference_column} not found in desired_order.')

        # Вставляем столбцы с lt, и кред оборотами в список заголовков
        insert_columns(do_columns, register_fields.debit_quantity_turnover_for_rename)
        insert_columns(ko_columns, register_fields.credit_quantity_turnover_for_rename)

        # Отберем столбцы c "Level_"
        level_columns = [col for col in self.pivot_table.columns if 'Level_' in col]
        # Сортируем столбцы с Level_ по числовому значению в их названиях
        level_columns.sort(key=lambda x: int(x.split('_')[1]))
        # Формируем итоговый порядок необходимых столбцов
        desired_order = [register_fields.file_name, register_fields.analytics] + desired_order + level_columns
        # Отбор существующих столбцов
        desired_order = [col for col in desired_order if col in self.pivot_table.columns]
        # Используем reindex для сортировки DataFrame
        self.pivot_table = self.pivot_table.reindex(columns=desired_order).copy()


    def unloading_pivot_table(self) -> None:
        """
        Выгружает финальный файл в Excel: сводная таблица и таблица с отклонениями на отдельных листах.
        """
        folder_path_summary_files = f"_Pivot_{self.file_type}.xlsx"
        with pd.ExcelWriter(folder_path_summary_files) as writer:
            self.pivot_table.to_excel(writer, sheet_name='Свод', index=False)
            self.pivot_table_check.to_excel(writer, sheet_name='Сверка', index=False)

    @staticmethod
    def process_end() -> None:
        print('Закончили обработку')
