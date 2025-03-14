# -*- coding: utf-8 -*-
"""
Created on Sat Dec 28 16:40:37 2024

@author: a.karabedyan
"""
import sys

def progress_bar(iteration, total, prefix='', length=40, fill='█', print_end='\n'):
    """
    Функция для отображения прогресс-бара с фиксированной длиной префикса.

    :param iteration: Текущий итерационный номер.
    :param total: Общее количество итераций.
    :param prefix: Префикс, который будет отображаться слева от прогресс-бара (должен занимать 55 символов).
    :param length: Длина прогресс-бара (в символах).
    :param fill: Символ, заполняющий прогресс-бар.
    :param print_end: Символ в конце строки (по умолчанию — новая строка).
    """
    pass
    # Убедимся, что префикс занимает ровно 55 символов
    # fixed_prefix = (prefix + ' ' * 55)[:55]

    # percent = "{0:.1f}".format(100 * (iteration / float(total)))
    # filled_length = int(length * iteration // total)
    # bar = fill * filled_length + '-' * (length - filled_length)

    # # Используем символ новой строки для полного прогресс-бара
    # if iteration < total:
    #     sys.stdout.write(f'\r{fixed_prefix} |{bar}| {percent}% Complete')
    # else:
    #     sys.stdout.write(f'\r{fixed_prefix} |{bar}| {percent}% Complete\n')

    # sys.stdout.flush()

