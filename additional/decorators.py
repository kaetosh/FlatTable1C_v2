# -*- coding: utf-8 -*-
"""
Created on Thu Feb 13 22:25:32 2025

@author: a.karabedyan
"""
import functools
import sys
from pathlib import Path
from loguru import logger
import pandas as pd
from tqdm import tqdm
from config import max_desc_length
from additional.ErrorClasses import ContinueIteration

# Настраиваем логгер для вывода только сообщений уровня ERROR
logger.remove()  # Удаляем стандартный обработчик
logger.add(sys.stderr, level="ERROR")  # Добавляем новый обработчик для вывода в консоль

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
            for file in tqdm(files_to_process, desc=prefix.ljust(max_desc_length)):
                # Устанавливаем self.file или self.oFile в зависимости от типа
                if isinstance(file, str):
                    self.file = file  # Если это строка
                elif isinstance(file, Path):
                    self.oFile = file  # Если это объект Path
                else:
                    continue  # Пропускаем, если тип не поддерживается
                #progress_bar(x + 1, len(files_to_process), prefix=prefix)
                while True:
                    try:
                        method(self, *args, **kwargs)  # Вызов метода
                        break  # Если все прошло успешно, выходим из внутреннего цикла
                    except ContinueIteration:
                        # Исключение для продолжения цикла
                        break  # Переход к следующему файлу
                    except pd.errors.EmptyDataError:
                        logger.debug(f"Файл {file} пуст.")
                        self.empty_files.append(file)
                        break
                    except pd.errors.ParserError as e:
                        logger.debug(f"Ошибка парсинг в файле {file}: {e}")
                        self.empty_files.append(file)
                        break
                    except KeyError as e:
                        logger.debug(f"Столбец не найден в файле {file}: {e}")
                        self.empty_files.append(file)
                        break
                    except FileNotFoundError:
                        logger.debug(f"Файл не найден: {file}")
                        self.empty_files.append(file)
                        break
                    except Exception as e:
                        logger.debug(f"Неожиданная ошибка в файле {file}: {str(e)}")
                        self.empty_files.append(file)
                        break  # Переход к следующему файлу
        return wrapper
    return decorator
