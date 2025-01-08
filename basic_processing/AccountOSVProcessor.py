import pandas as pd
from basic_processing.FileProcessor import IFileProcessor
pd.options.mode.copy_on_write = False
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)


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

            replacement_values = ['Дебет_начало', 'Кредит_начало', 'Дебет_оборот', 'Кредит_оборот', 'Дебет_конец',
                                  'Кредит_конец']

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