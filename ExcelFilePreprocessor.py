# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 16:35:32 2024

@author: a.karabedyan
"""

import openpyxl
from config import analisys_filds

class ExcelFilePreprocessor:
    
    def __init__(self, excel_files, file_type):
        self.excel_files = excel_files
        self.file_type = file_type
    
    def preprocessor_openpyxl(self):

        for oFile in self.excel_files:
            workbook = None
            try:
                workbook = openpyxl.load_workbook(oFile)
            except Exception as e:
                print(f'''{oFile}: Ошибка обработки файла. Возможно открыт обрабатываемый файл. Закройте этот файл и снова запустите скрипт.
                                      Ошибка: {e}''')

            sheet = workbook.active
        
            # Снимаем объединение ячеек
            merged_cells_ranges = list(sheet.merged_cells.ranges)
            for merged_cell_range in merged_cells_ranges:
                sheet.unmerge_cells(str(merged_cell_range))
        
            # Столбец с уровнями группировок
            sheet.insert_cols(idx=1)
            for row_index in range(1, sheet.max_row + 1):
                cell = sheet.cell(row=row_index, column=1)
                cell.value = sheet.row_dimensions[row_index].outline_level
            
            if self.file_type == 'account_analisys':
                sheet.insert_cols(idx=2)
                for row in sheet.iter_rows(values_only=True):

                    found_value = next((value for value in [analisys_filds.upp.corresponding_account,
                                                            analisys_filds.notupp.corresponding_account] if value in row), None)
                    
                    if found_value is not None:
                        kor_schet_col_index = row.index(found_value) + 1  # We add 1 because indexing starts from 0
                        # Мы заполняем новый столбец значениями, основанными на форматировании ячеек курсивом
                        for row_index in range(2, sheet.max_row + 1):  # We start with 2 to skip the title
                            kor_schet_cell = sheet.cell(row=row_index, column=kor_schet_col_index)
                            new_cell = sheet.cell(row=row_index, column=2)
                            new_cell.value = 1 if kor_schet_cell.font and kor_schet_cell.font.italic else 0
                        break

            workbook.save(oFile)
            workbook.close()
