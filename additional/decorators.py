# -*- coding: utf-8 -*-
"""
Created on Thu Feb 13 22:25:32 2025

@author: a.karabedyan
"""
import functools
from additional.progress_bar import progress_bar
from additional.ErrorClasses import ContinueIteration

import functools

def catch_and_log_exceptions(prefix=''):
    """
    Декоратор для ловли исключений в методе,
    выводит сообщение об ошибке и добавляет проблемный файл в список self.empty_files.
    Очищает словарь self.dict_df для текущего файла в случае ошибки.
    Отображение прогресс-бара.
    """

    def decorator(method):
        @functools.wraps(method)
        def wrapper(self, *args, **kwargs):
            files_to_process = list(self.dict_df.keys()) or self.excel_files
            for x, self.file in enumerate(files_to_process):
                progress_bar(x + 1, len(files_to_process), prefix=f'{prefix} {self.file}')
                while True:
                    try:
                        method(self, *args, **kwargs)  # Вызов метода
                        break  # Если все прошло успешно, выходим из внутреннего цикла
                    except ContinueIteration:
                        # Исключение для продолжения цикла
                        break  # Переход к следующему файлу
                    except Exception as ee:
                        #error_message = f"Ошибка при обработке {self.file}: {str(ee)}"
                        #print('\n', error_message)
                        self.empty_files.append(self.file)
                        if self.file in self.dict_df:
                            del self.dict_df[self.file]
                        break  # Переход к следующему файлу
        return wrapper
    return decorator
