# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 14:40:14 2024

@author: a.karabedyan
"""

from FileProcessor import IFileProcessor, AccountTurnoverProcessor

class FileProcessorFactory:
    @staticmethod
    def create_processor(file_type: str) -> IFileProcessor:
        if file_type == "account_turnover":
            return AccountTurnoverProcessor()
        else:
            raise ValueError(f"Unknown file type: {file_type}")
