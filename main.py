# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 14:42:39 2024

@author: a.karabedyan
"""
from typing import Literal
from basic_processing.FileProcessorFactory import FileProcessorFactory

def main():
    file_type: list[Literal['account_turnover', 'account_analysis', 'account_osv']] = ['account_turnover', 'account_analysis', 'account_osv']
    while True:
        try:
            number_register = int(
                input('Введи номер для обрабатываемого регистра\n0 - Обороты счета\n1 - Анализ счета\n2 - ОСВ:\n '))
            if number_register in [0, 1, 2]:
                break
            else:
                print("Некорректный ввод. Пожалуйста, введите 0, 1 или 2.")
        except ValueError:
            print("Некорректный ввод. Пожалуйста, введите целое число.")
    processor = FileProcessorFactory.create_processor(file_type[number_register])
    processor.conversion()
    processor.preprocessing()
    processor.general_table_header()
    processor.special_table_header()
    processor.handle_missing_values()
    processor.horizontal_structure()
    processor.corr_account_col()
    processor.revolutions_before_processing()
    processor.lines_delete()
    processor.revolutions_after_processing()
    processor.joining_tables()
    processor.shiftable_level()
    processor.reorder_table_columns()
    processor.unloading_pivot_table()
    processor.process_end()

if __name__ == "__main__":
    main()
