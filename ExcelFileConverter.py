# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 16:18:41 2024

@author: a.karabedyan
"""

import win32com.client


class ExcelFileConverter:

    def __init__(self, excel_files):
        self.excel_files = excel_files
    
    def save_as_xlsx_no_alert(self):
        excel_app = win32com.client.Dispatch('Excel.Application')
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        
        for oFile in self.excel_files:
            self.convert_file(excel_app, oFile)
                
        excel_app.Quit()
        print('Файлы пересохранены')

    def convert_file(self, excel_app, oFile):
        wb = excel_app.Workbooks.Open(str(oFile))
        wb.SaveAs(str(oFile), FileFormat=51)  # Сохраняем под тем же именем, меняя формат
        wb.Close(SaveChanges=False)
