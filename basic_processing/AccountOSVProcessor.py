"""
В ОСВ наименования сальдо/оборотов и дебет/кредит в разных строках,
поэтому добавляем для дебета или кредита 'начало', 'оборот', 'конец'
"""

from basic_processing.FileProcessor import IFileProcessor
from additional.progress_bar import progress_bar


class AccountOSVProcessor(IFileProcessor):
    def special_table_header(self) -> None:
        for x, file in enumerate(self.dict_df):
            progress_bar(x + 1, len(self.dict_df), prefix='Установка специальных заголовков в таблицах')
            # Выгрузим обрабатываемую таблицу из хранилища таблиц
            df, sign_1c, register, register_fields, *_ = self._get_data_from_table_storage(file, self.dict_df)

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

            # запишем таблицу в словарь
            self.dict_df[file].table = df