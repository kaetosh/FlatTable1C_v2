# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 16:18:41 2024

@author: kaetosh

Обработка ошибки при использовании openpyxl:
KeyError: "There is no item named 'xl/sharedStrings.xml' in the archive".
Excel-файл по своей сути — это архив данных, который можно открыть любым архиватором.
В его составе есть файл (смотри папку "xl" архива) sharedStrings.xml, который хранит все текстовые поля excel-файла.
Некоторые версии 1С при выгрузке и сохранении своих отчетов (ОСВ, анализ счета и т.д.)
использует старую версию excel, которая создает файл SharedStrings.xml и его название начинается
с верхнего регистра. А современные версии Excel делают это с нижнего.
Пересохранение позволяет пересоздать проблемный файл как xl/sharedStrings.xml,
имя которого начинается с нижнего регистра.
"""
import os
import win32com.client
from typing import List
from pathlib import Path
from additional.progress_bar import progress_bar

class ExcelFileConverter:
    @staticmethod
    def save_as_xlsx_no_alert(excel_files: List[Path]) -> None:
        excel_app = win32com.client.Dispatch('Excel.Application')
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        for i, file in enumerate(excel_files):
            progress_bar(i + 1, len(excel_files), prefix='Пересохранение исходных файлов')
            ExcelFileConverter.convert_file(excel_app, file)
        excel_app.Quit()
    @staticmethod
    def convert_file(excel_app: win32com.client.CDispatch, file: Path) -> None:
        name_xlsx = str(file).replace('.xls', '.xlsx') if str(file).endswith('.xls') else str(file)
        wb = excel_app.Workbooks.Open(str(file))
        wb.SaveAs(name_xlsx, FileFormat=51)  # Сохраняем под тем же именем, меняя формат
        wb.Close(SaveChanges=False)
        if str(file).endswith('.xls'):
            os.remove(str(file))