"""
В общей ОСВ наименования сальдо/оборотов и дебет/кредит в разных строках,
поэтому добавляем для дебета или кредита 'начало', 'оборот', 'конец'
"""

from basic_processing.FileProcessor import IFileProcessor
from additional.decorators import catch_and_log_exceptions

class OSVGeneralProcessor(IFileProcessor):
    @catch_and_log_exceptions(prefix='Установка специальных заголовков в таблицах')
    def special_table_header(self) -> None:
        # Выгрузим обрабатываемую таблицу из хранилища таблиц
        df, sign_1c, register, register_fields, *_ = self._get_data_from_table_storage(self.file, self.dict_df)

        # счетчик того, сколько столбцов Дебет и Кредит
        counters = {'Дебет': 0, 'Кредит': 0}

        def update_account_list(item):
            if item in counters:
                counters[item] += 1
                return f"{item}_{['начало', 'оборот', 'конец'][counters[item] - 1]}"
            return item

        # берем строку, где есть дебет/кредит (первая, сразу после шапки)
        # и дополняем к этим значениям 'начало', 'оборот', 'конец'
        updated_list = [update_account_list(item) for item in df.iloc[0]]

        name_col = df.columns.to_list()

        replacement_values = [i for i in register_fields.get_attributes_by_suffix('for_rename') if 'Количество_' not in i]

        # обновляем шапку таблицы
        for index, value in enumerate(updated_list):
            if value in replacement_values:
                name_col[index] = value

        df.columns = name_col

        df = df.loc[:, df.columns.notna()]
        df.columns = df.columns.astype(str)
        df = df.iloc[1:]

        #переименуем столбец с названиями счетов для единообразия выгрузок УПП и неУПП
        df = df.rename(columns={register_fields.account_name: register_fields.account_name_for_rename})

        # добавим нолик, если счет до 10, чтобы было 01 02 04 05 07 08 09
        # Предполагаем, что 'additional_columns' содержит наименования, которые вы хотите проверить
        columns_to_check = register_fields.get_attributes_by_suffix('for_rename')

        # Проверка существующих столбцов
        existing_columns = df.columns.intersection(columns_to_check).tolist()

        # Удаление строк с NaN только для существующих столбцов
        df = df.dropna(subset=existing_columns, how='all')


        df[register_fields.analytics] = df[register_fields.analytics].apply(lambda x: f'0{x}' if len(str(x)) == 1 else x)
        df[register_fields.analytics] = df[register_fields.analytics].fillna('').astype('str')

        # запишем таблицу в словарь
        self.dict_df[self.file].table = df


    def handle_missing_values(self) -> None:
        pass


    def horizontal_structure(self) -> None:
        pass


    def revolutions_before_processing(self) -> None:
        pass


    def lines_delete(self) -> None:
        pass


    def revolutions_after_processing(self) -> None:
        pass


    def shiftable_level(self) -> None:
        pass

