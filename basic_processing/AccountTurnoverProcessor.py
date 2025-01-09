from typing import List, Dict
import pandas as pd
from basic_processing.FileProcessor import IFileProcessor
pd.options.mode.copy_on_write = False
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
from config import new_names, turnover_fields
from additional.ErrorClasses import NoExcelFilesError

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
                    print('turnover_type', turnover_type)
                    name_attribute = turnover_fields.get_inner_attribute_name_by_value(turnover_type)
                    index_turnover_type: int or False = df.columns.get_loc(
                        turnover_type) if turnover_type in df.columns else False
                    self.dict_df[file].set_index_column(name_attribute, index_turnover_type)
                    if ('debit' in name_attribute) or ('credit' in name_attribute) and index_turnover_type:
                        indices_to_rename.append(index_turnover_type)
                except TypeError:
                    continue  # Пропускаем, если turnover_type == None
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
                    if debit_turnover_index is not None and idx > debit_turnover_index and (
                            upper_bound_index is None or idx < upper_bound_index):
                        list_do_columns.append(f'{col}_до')
                    else:
                        list_do_columns.append(col)
                # Обновляем названия столбцов в DataFrame
                df.columns = list_do_columns

            if fields_account_turnover.credit_turnover in df.columns:
                list_ko_columns: List[str] = []

                # Определяем границы для добавления префикса 'ко'
                end_balances_index: int = max(end_debit_balance_index or -1,
                                              end_credit_balance_index or -1)  # Определяем конец диапазона

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
            rename_dict: Dict[str, str] = {current_columns[i]: new_names[j] for j, i in enumerate(indices_to_rename) if
                                           i}

            # Переименовываем столбцы
            df = df.rename(columns=rename_dict)

            # запишем таблицу в словарь
            self.dict_df[file].table = df
