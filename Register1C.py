# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 15:26:33 2024

@author: a.karabedyan
"""

from dataclasses import dataclass
import pandas as pd

# поля регистров 1с
class FieldsRegister:
    def __init__(self,
                 analytics: str or None = None,
                 quantity: str or None = None,
                 type_connection: str or None = None,
                 corresponding_account: str or None = None,
                 start_debit_balance: str or None = None,
                 start_credit_balance: str or None = None, 
                 debit_turnover: str or None = None,
                 credit_turnover: str or None = None, 
                 end_debit_balance: str or None = None,
                 end_credit_balance: str or None = None):
        self.analytics = analytics
        self.quantity = quantity
        self.type_connection = type_connection
        self.corresponding_account = corresponding_account
        self.start_debit_balance = start_debit_balance
        self.start_credit_balance = start_credit_balance
        self.debit_turnover = debit_turnover
        self.credit_turnover = credit_turnover
        self.end_debit_balance = end_debit_balance
        self.end_credit_balance = end_credit_balance
    def __iter__(self):
        return iter((self.analytics,
                     self.quantity,
                     self.type_connection,
                     self.corresponding_account,
                     self.start_debit_balance,
                     self.start_credit_balance,
                     self.debit_turnover,
                     self.credit_turnover,
                     self.end_debit_balance,
                     self.end_credit_balance))

class Register_1C:
    def __init__(self,
                 name_register: str or None,
                 upp: FieldsRegister or None = None,
                 notupp: FieldsRegister or None = None):
        self.name_register = name_register
        self.upp = upp if upp is not None else []
        self.notupp = notupp if notupp is not None else []
    def get_attribute_name_by_value(self, value):
        """
        Метод, который ищет атрибут внешнего класса Register_1C по значению
        атрибута внутреннего класса FieldsRegister
        (признак версии 1С: УПП или нет).
        """
        for attr_name, attr_value in vars(self).items():
            if isinstance(attr_value, FieldsRegister):
                # Проходим по всем атрибутам FieldsRegister
                for inner_attr_name, inner_attr_value in vars(attr_value).items():
                    if inner_attr_value == value:
                        return attr_name  # Возвращаем имя внешнего атрибута
        return None  # Если значение не найдено
    def get_inner_attribute_by_value(self, value):
        """
        Метод, который ищет атрибут внутреннего класса FieldsRegister по значению
        атрибута внутреннего класса FieldsRegister.
        """
        for attr_value in (self.upp, self.notupp):
            if isinstance(attr_value, FieldsRegister):
                for inner_attr_name, inner_attr_value in vars(attr_value).items():
                    if inner_attr_value == value:
                        return inner_attr_name  # Возвращаем имя атрибута внутреннего класса
        return None  # Если значение не найдено
    def __iter__(self):
        yield from self.upp
        yield from self.notupp
    
    def __str__(self):
        return(f'{self.name_register}')

@dataclass
class Table_storage:
    table: pd.DataFrame
    register: Register_1C
    sign_1C: str
    table_type_connection: pd.DataFrame = None

    def set_index_column(self, name_atribure, value):
        # Формируем имя атрибута
        attr_name = f'index_column_{name_atribure}'
        # Устанавливаем значение атрибута
        setattr(self, attr_name, value)
