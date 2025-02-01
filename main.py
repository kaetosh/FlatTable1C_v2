# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 14:42:39 2024

@author: a.karabedyan
"""

from basic_processing.FileProcessorFactory import FileProcessorFactory

def main():
    file_type = ['account_turnover', 'account_analysis', 'account_osv']
    processor = FileProcessorFactory.create_processor(file_type[0])
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
    #processor.rename_columns()
    processor.unloading_pivot_table()
    processor.process_end()

if __name__ == "__main__":
    main()

    #для теста git
