# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 14:40:14 2024

@author: a.karabedyan
"""

from FileProcessor import IFileProcessor, AccountTurnoverProcessor, AccountAnalysisProcessor, AccountOSVProcessor

class FileProcessorFactory:
    @staticmethod
    def create_processor(file_type: str) -> IFileProcessor:
        match file_type:
            case "account_turnover":
                return AccountTurnoverProcessor(file_type)
            case "account_analysis":
                return AccountAnalysisProcessor(file_type)
            case "account_osv":
                return AccountOSVProcessor(file_type)
            case _:
                raise ValueError(f"Неизвестный тип файла: {file_type}")
