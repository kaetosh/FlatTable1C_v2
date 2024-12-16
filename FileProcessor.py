# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 14:32:00 2024

@author: a.karabedyan
"""

import os
from pathlib import Path
from ExcelFileConverter import ExcelFileConverter
from ExcelFilePreprocessor import ExcelFilePreprocessor


class IFileProcessor:
   
    def __init__(self):
        self.path_folder_excel_files = Path(os.getcwd())
        files = list(self.path_folder_excel_files.iterdir())
        print('*******************', files)
        #files = os.listdir(self.path_folder_excel_files)
        self.excel_files = [file for file in files if (str(file).endswith('.xlsx') or str(file).endswith('.xls')) and '_СВОД_' not in str(file)]
        #self.converter = ExcelFileConverter(self.path_folder_excel_files)
        self.converter = ExcelFileConverter(self.excel_files)
        self.preprocessor = ExcelFilePreprocessor(self.path_folder_excel_files)
        #self.preprocessor = ExcelFilePreprocessor(self.excel_files)
        

    def process_start(self):
        print('Начинаем конвертацию файлов...')
        self.converter.save_as_xlsx_no_alert()
        
        print ('Начинаем предобработку файлов...')
        self.preprocessor.preprocessor_openpyxl()
   
    def process(self):
        print('Основная обработка')
    
    def process_end(self):
        print('Закончили обработку')

class AccountTurnoverProcessor(IFileProcessor):
    def process(self):
        print('основная обработка TurnOver account')
        #for file in self.path_folder_excel_files
