from typing import List, Dict
import pandas as pd

from additional.ErrorClasses import NoExcelFilesError
from basic_processing.FileProcessor import IFileProcessor
pd.options.mode.copy_on_write = False
from config import new_names
from additional.decorators import catch_and_log_exceptions, logger

class AccountTurnoverProcessor(IFileProcessor):
    @catch_and_log_exceptions(prefix='Установка специальных заголовков в таблицах')
    def special_table_header(self) -> None:
        """
        В регистре Обороты счета наименования столбцов с оборотами с корреспондирующими счетами не имеют признака
        дебетового или кредитового оборота.
        Задача данного метода найти такие столбцы и разделить их на дебетовые и кредитовые, добавив к наименованию
        'до' и 'ко' соответственно.
        """
        # Выгрузим обрабатываемую таблицу из хранилища таблиц
        df, sign_1c, register, register_fields, *_ = self._get_data_from_table_storage(self.file, self.dict_df)

        # Избавимся от столбцов с пустыми наименованиями, поскольку это полностью пустые столбцы
        df = df.loc[:, df.columns.notna()]

        # Приведем наименования столбцов к строковому формату
        df.columns = df.columns.astype(str)

        # Список для номеров столбцов, к которым нужно добавить "до" и "ко"
        indices_to_rename: List[int] = []

        # Пройдемся по всем непустым атрибутам (полям с сальдо/оборотами) регистра
        # и запишем в хранилище номера столбцов с сальдо и оборотами
        # чтобы далее между этими помеченными индексами столбцами корректно добавить "до" "ко"
        # к именам столбцов
        for turnover_type in [fild for fild in register_fields if fild]:
            try:
                if turnover_type in df.columns:
                    index_turnover_type: int = df.columns.get_loc(turnover_type)
                    name_attribute = register.get_inner_attribute_name_by_value(turnover_type)
                    if ('debit' in name_attribute) or ('credit' in name_attribute):
                        self.dict_df[self.file].set_index_column(name_attribute, index_turnover_type)
                        indices_to_rename.append(index_turnover_type)
            except StopIteration:
                continue  # Если ничего не найдено

        # Установим значения по умолчанию для номеров столбцов с сальдо и оборотами
        match register_fields.analytics:
            case 'Субконто':
                end_debit_balance_index: int = 8 if register_fields.type_connection else 5
                end_credit_balance_index: int = 9 if register_fields.type_connection else 6
                credit_turnover_index: int = 7 if register_fields.type_connection else 4
            case 'Счет':
                end_debit_balance_index: int = 6 if register_fields.quantity else 5
                end_credit_balance_index: int = 7 if register_fields.quantity else 6
                credit_turnover_index: int = 5 if register_fields.quantity else 4
            case _:
                raise NoExcelFilesError

        if register_fields.debit_turnover in df.columns:
            # Определяем верхнюю границу для добавления префикса 'до'
            debit_turnover_index: int = getattr(self.dict_df[self.file], 'index_column_debit_turnover', None)
            credit_turnover_index: int = getattr(self.dict_df[self.file], 'index_column_credit_turnover', None)
            end_debit_balance_index: int = getattr(self.dict_df[self.file], 'index_column_end_debit_balance', None)
            end_credit_balance_index: int = getattr(self.dict_df[self.file], 'index_column_end_credit_balance', None)
            upper_bound_index: int = credit_turnover_index or end_debit_balance_index or end_credit_balance_index

            # Создаем новый список названий столбцов с префиксом 'до'
            list_do_columns: List[str] = []
            for idx, col in enumerate(df.columns):
                # Если нашли индекс 'дебетового оборота', добавляем префикс 'до' при выполнении условий
                if debit_turnover_index is not None and idx > debit_turnover_index and (
                        upper_bound_index is None or idx < upper_bound_index):
                    list_do_columns.append(f'{col}_до')
                else:
                    list_do_columns.append(col)
            # Обновляем названия столбцов в DataFrame
            df.columns = list_do_columns

        if register_fields.credit_turnover in df.columns:
            list_ko_columns: List[str] = []

            # Определяем границы для добавления префикса 'ко'
            end_balances_index: int = max(end_debit_balance_index or -1, end_credit_balance_index or -1)  # Определяем конец диапазона

            for idx, col in enumerate(df.columns):
                # Добавляем префикс, если индекс в нужном диапазоне
                if credit_turnover_index is not None and idx > credit_turnover_index and (
                        end_balances_index == -1 or idx < end_balances_index):
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
        self.dict_df[self.file].table = df
        logger.debug(f'10={df.columns}')

