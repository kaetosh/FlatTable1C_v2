# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 14:42:39 2024

@author: a.karabedyan
"""

from FileProcessorFactory import FileProcessorFactory

def main():
    file_type = ['account_turnover', 'account_analisys', 'account_osv']
    processor = FileProcessorFactory.create_processor(file_type[2])
    processor.process_start()
    processor.general_table_header()
    processor.process_end()

if __name__ == "__main__":
    main()
