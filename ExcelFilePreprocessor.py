# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 16:35:32 2024

@author: a.karabedyan
"""

import openpyxl

class ExcelFilePreprocessor:
    
    def __init__(self, excel_files):
        self.excel_files = excel_files
    
    def preprocessor_openpyxl(self):

        for oFile in self.excel_files:
            workbook = None
            try:
                workbook = openpyxl.load_workbook(oFile)
            except Exception as e:
                print(f'''{oFile}: Ошибка обработки файла. Возможно открыт обрабатываемый файл. Закройте этот файл и снова запустите скрипт.
                                      Ошибка: {e}''')
            continue  # Пропускаем файл и продолжаем с другими файлами

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
            sheet['A1'] = "Уровень"
            workbook.save(oFile)
            #workbook.close()
