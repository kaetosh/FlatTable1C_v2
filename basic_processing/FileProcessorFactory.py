# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 14:40:14 2024

@author: a.karabedyan

Возвращает класс обработчика в зависимости от типа обрабатываемых регистров:
осв, оборотов счета или анализов счета
"""
from typing import Literal

from additional.ErrorClasses import NoExcelFilesError
from basic_processing.FileProcessor import IFileProcessor
from basic_processing.AccountTurnoverProcessor import AccountTurnoverProcessor
from basic_processing.AccountAnalisysProcessor import AccountAnalysisProcessor
from basic_processing.AccountOSVProcessor import AccountOSVProcessor

class FileProcessorFactory:
    @staticmethod
    def create_processor(file_type: Literal['account_turnover',
                                            'account_analysis',
                                            'account_osv']) -> IFileProcessor:
        match file_type:
            case "account_turnover":
                return AccountTurnoverProcessor(file_type)
            case "account_analysis":
                return AccountAnalysisProcessor(file_type)
            case "account_osv":
                return AccountOSVProcessor(file_type)
            case _:
                raise NoExcelFilesError
