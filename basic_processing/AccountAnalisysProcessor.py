import pandas as pd
from basic_processing.FileProcessor import IFileProcessor
pd.options.mode.copy_on_write = False
import numpy as np
from config import exclude_values, accounts_without_subaccount
from additional.progress_bar import progress_bar

class AccountAnalysisProcessor(IFileProcessor):
    def handle_missing_values(self):
        for x, file in enumerate(self.dict_df):
            progress_bar(x + 1, len(self.dict_df), prefix='Установка специальных заголовков в таблицах')
            df, sign_1c, register, register_fields, *_ = self._get_data_from_table_storage(file, self.dict_df)

            # сохраним столбец "Вид связи КА" в отдельный фрейм
            # чтобы в методе lines_delete проставить пропущенные значения "Вид связи КА"
            if register_fields.type_connection in df.columns:
                df_type_connection = (
                    df
                    .drop_duplicates(subset=[register_fields.analytics, register_fields.type_connection])
                    .dropna(subset=[register_fields.analytics,
                                    register_fields.type_connection])  # Удаляем строки с NaN в указанных столбцах
                    .loc[:, [register_fields.analytics, register_fields.type_connection]]
                )
                self.dict_df[file].table_type_connection = df_type_connection

            # Проверка на пропуски и условия для заполнения
            mask = (
                    df[register_fields.analytics].isna() &
                    ~df[register_fields.corresponding_account].apply(self._is_accounting_code) &
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
                lambda x: f'0{x}' if (len(x) == 1 and self._is_accounting_code(x)) else x)

            # Запишем таблицу в словарь
            self.dict_df[file].table = df

    def corr_account_col(self) -> None:
        for x, file in enumerate(self.dict_df):
            progress_bar(x + 1, len(self.dict_df), prefix='Установка столбца с корреспондирубщим счетом в таблицах')
            df, sign_1c, register, register_fields, *_ = self._get_data_from_table_storage(file, self.dict_df)

            # добавим столбец корр.счет, взяв его из основного столбца, при условии, что значение - бухгалтерских счет (функция is_accounting_code)
            df['Корр_счет'] = df[register_fields.corresponding_account].apply(
                lambda x: str(x) if (self._is_accounting_code(x) or str(x) == '0') else None)

            # добавим нолик, если счет до 10, чтобы было 01 02 04 05 07 08 09
            df['Корр_счет'] = df['Корр_счет'].apply(lambda x: f'0{x}' if len(str(x)) == 1 else x)

            # добавим нолик к счетам и в основном столбце
            df['Корр_счет'] = df['Корр_счет'].apply(lambda x: f'0{x}' if len(str(x)) == 1 else x)

            # Заполнение пропущенных значений в столбце значениями из предыдущей строки
            df['Корр_счет'] = df['Корр_счет'].ffill()

            # Запишем таблицу в словарь
            self.dict_df[file].table = df

    def revolutions_before_processing(self) -> None:
        for x, file in enumerate(self.dict_df):
            progress_bar(x + 1, len(self.dict_df), prefix='Сохраняем данные по оборотам до обработки в таблицах')
            df, sign_1c, register, register_fields, *_ = self._get_data_from_table_storage(file, self.dict_df)
            df_for_check = df[[register_fields.corresponding_account,
                               register_fields.debit_turnover,
                               register_fields.credit_turnover]].copy()
            df_for_check['Кор.счет_ЧЕК'] = df_for_check[register_fields.corresponding_account].apply(
                lambda x: str(x) if self._is_accounting_code(x) else None).copy()
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

            df_for_check = df_for_check.rename(columns={register_fields.debit_turnover: 'С кред. счетов',
                                                        register_fields.credit_turnover: 'В дебет счетов'})

            # запишем таблицу в словарь
            self.dict_df[file].table_for_check = df_for_check

    def lines_delete(self):
        for x, file in enumerate(self.dict_df):
            progress_bar(x + 1, len(self.dict_df), prefix='Удаляем строки с дублирующими оборотами в таблицах')
            df, sign_1c, register, register_fields, *_ = self._get_data_from_table_storage(file, self.dict_df)
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
            acc_with_sub = [i for i in all_acc_dict if self._is_parent(i, list(all_acc_dict.keys()))]
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
                del_acc_with_94 = [i for i in acc_with_94 if i != '94.Н']
            del_acc = list(set(del_acc + del_acc_with_94))


            for i in accounts_without_subaccount:
                unwanted_subaccounts = [n for n in all_acc_dict if i in n]
                del_unwanted_subaccounts = [n for n in unwanted_subaccounts if n != i]
                del_acc = list(set(del_acc + del_unwanted_subaccounts))


            for i in accounts_without_subaccount:
                if i in del_acc:
                    del_acc.remove(i)


            # Оригинальный столбец с корр.счетами может содержать счета без 0, т.е. не 08, а 8
            # добавим в список для удаления счетов счета без 0
            # Создание нового списка
            list_of_accounts_without_zeros_int = []
            list_of_accounts_without_zeros_str = []
            for item in del_acc:
                # Проверка, является ли элемент целым числом с нулями
                if item.isdigit() and item.startswith('0') and len(item) > 1:
                    # Преобразуем строку в целое число и обратно в строку
                    list_of_accounts_without_zeros_str.append(str(int(item)))
                    list_of_accounts_without_zeros_int.append((int(item)))
            del_acc.extend(list_of_accounts_without_zeros_int)
            del_acc.extend(list_of_accounts_without_zeros_str)

            df[register_fields.corresponding_account] = df[register_fields.corresponding_account].apply(
                lambda x: str(x))

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
                merged = df.merge(self.dict_df[file].table_type_connection, on=register_fields.analytics, how='left',
                                  suffixes=('', '_B'))
                df[register_fields.type_connection] = df[register_fields.type_connection].fillna(
                    merged[f'{register_fields.type_connection}_B'])

            df = df[
                ~df[register_fields.corresponding_account].isin(
                    exclude_values) &  # Исключение определенных значений (Сальдо, Оборот и т.д.)
                ~df[register_fields.corresponding_account].isin(del_acc)
                # Исключение счетов, по которым есть расшифровка субконто (60, 60.01 и т.д.)
                ].copy()

            df = df[df['Курсив'] == 0].copy()
            df[register_fields.corresponding_account] = df[register_fields.corresponding_account].astype(str)

            shiftable_level = 'Level_0'
            list_level_col = [i for i in df.columns.to_list() if i.startswith('Level')]
            for i in list_level_col[::-1]:
                if all(df[i].apply(self._is_accounting_code)):
                    shiftable_level = i
                    break

            df['Субсчет'] = df.apply(
                lambda row: row[shiftable_level] if (str(row[shiftable_level]) != '7') else f"0{row[shiftable_level]}",
                axis=1)  # 07 без субсчетов
            df['Субсчет'] = df.apply(
                lambda row: 'Без_субсчетов' if not self._is_accounting_code(row['Субсчет']) else row['Субсчет'], axis=1)

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
                lambda x: 'Не расшифровано' if self._is_accounting_code(x) else x)

            df = df.dropna(subset=['С кред. счетов', 'В дебет счетов'], how='all')
            df = df[(df['С кред. счетов'] != 0) | (df['С кред. счетов'].notna())]
            df = df[(df['В дебет счетов'] != 0) | (df['В дебет счетов'].notna())]

            # Запишем таблицу в словарь
            self.dict_df[file].table = df

    def revolutions_after_processing(self) -> None:
        for x, file in enumerate(self.dict_df):
            progress_bar(x + 1, len(self.dict_df), prefix='Сохраняем данные по оборотам после обработки в таблицах')
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

            turnover_deb = merged_df[
                'С кред. счетов_df_for_check'] if 'С кред. счетов_df_for_check' in merged_df.columns else 0
            turnover_cre = merged_df[
                'В дебет счетов_df_for_check'] if 'В дебет счетов_df_for_check' in merged_df.columns else 0

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

    def reorder_table_columns(self) -> None:
        progress_bar(1, 3, prefix='Сортируем столбцы в нужном порядке:')
        list_lev = [i for i in self.pivot_table.columns.to_list() if 'Level' in i]
        for n in list_lev[::-1]:
            if all(self.pivot_table[n].apply(self._is_accounting_code)):
                self.pivot_table['Субсчет'] = self.pivot_table[n].copy()
                break

        progress_bar(2, 3, prefix='Сортируем столбцы в нужном порядке:')
        for p in list_lev:
            if not all(self.pivot_table[p].apply(self._is_accounting_code)):
                self.pivot_table['Аналитика'] = self.pivot_table['Аналитика'].where(
                    self.pivot_table['Аналитика'] != 'Не_указано', self.pivot_table[p])
                break

        progress_bar(3, 3, prefix='Сортируем столбцы в нужном порядке:')
        self.pivot_table['Субсчет'] = self.pivot_table.apply(
            lambda row: row['Субсчет'] if (str(row['Субсчет']) != '7') else f"0{row['Субсчет']}",
            axis=1)

