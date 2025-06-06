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
from tqdm import tqdm
from config import max_desc_length
import tempfile
from zipfile import ZipFile
import shutil

class ExcelFileConverter:
    @staticmethod
    def save_as_xlsx_no_alert(excel_files: List[Path]) -> None:
        # excel_app = win32com.client.Dispatch('Excel.Application')
        # excel_app.Visible = False
        # excel_app.DisplayAlerts = False
        for file in tqdm(excel_files, desc="Пересохранение исходных файлов".ljust(max_desc_length)):
            ExcelFileConverter.fix_excel_filename(file)
            # ExcelFileConverter.convert_file(excel_app, file)
        # excel_app.Quit()
    @staticmethod
    def convert_file(excel_app: win32com.client.CDispatch, file: Path) -> None:
        name_xlsx = str(file).replace('.xls', '.xlsx') if str(file).endswith('.xls') else str(file)
        wb = excel_app.Workbooks.Open(str(file))
        wb.SaveAs(name_xlsx, FileFormat=51)  # Сохраняем под тем же именем, меняя формат
        wb.Close(SaveChanges=False)
        if str(file).endswith('.xls'):
            os.remove(str(file))
    @staticmethod
    def fix_excel_filename(excel_file_path: Path) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_folder = Path(tmp_dir)

            with ZipFile(excel_file_path) as excel_container:
                excel_container.extractall(tmp_folder)

            wrong_file_path = tmp_folder / 'xl' / 'SharedStrings.xml'
            correct_file_path = tmp_folder / 'xl' / 'sharedStrings.xml'

            if wrong_file_path.exists():
                os.rename(wrong_file_path, correct_file_path)

                # Создаем архив с новым именем
                tmp_zip_path = excel_file_path.with_suffix('.zip')
                shutil.make_archive(str(excel_file_path.with_suffix('')), 'zip', tmp_folder)

                # Удаляем исходный файл и переименовываем новый архив
                if excel_file_path.exists():
                    os.remove(excel_file_path)

                os.rename(tmp_zip_path, excel_file_path)
