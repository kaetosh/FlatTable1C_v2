# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 15:26:33 2024

@author: kaetosh

В работе три основных регистра из 1С:
- ОСВ
- Обороты счета
- Анализ счета

Классы помогают идентифицировать эти регистры по их полям, их названия отличаются для версий 1С:Предприятие 8.3:
1. Конфигурация "Бухгалтерия предприятия",
   Конфигурация "1С:ERP Агропромышленный комплекс",
   Конфигурация "1С:ERP Управление предприятием 2" (признак в коде 'notupp')
2. Конфигурация "Управление производственным предприятием" (признак в коде 'upp')

Экземпляры данных классов определены в config.py
"""

from typing import Optional
from dataclasses import dataclass, field
import pandas as pd

# поля регистров 1с
class FieldsRegister:
    def __init__(self,
                 analytics: Optional[str] = None,
                 quantity: Optional[str] = None,
                 type_connection: Optional[str] = None,
                 corresponding_account: Optional[str] = None,
                 start_debit_balance: Optional[str] = None,
                 start_credit_balance: Optional[str] = None,
                 debit_turnover: Optional[str] = None,
                 credit_turnover: Optional[str] = None,
                 end_debit_balance: Optional[str] = None,
                 end_credit_balance: Optional[str] = None):
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

class Register1c:
    def __init__(self,
                 name_register: Optional[str],
                 upp: Optional[FieldsRegister] = None,
                 notupp: Optional[FieldsRegister] = None):
        self.name_register = name_register
        self.upp = upp if upp is not None else []
        self.notupp = notupp if notupp is not None else []
    def get_outer_attribute_name_by_value(self, value):
        """
        Метод, который ищет атрибут внешнего класса Register1c по значению
        атрибута внутреннего класса FieldsRegister
        (признак версии 1С: УПП или нет).
        """
        for attr_name, attr_value in vars(self).items():
            if isinstance(attr_value, FieldsRegister):
                for inner_attr_name, inner_attr_value in vars(attr_value).items():
                    if inner_attr_value == value:
                        return attr_name  # Возвращаем имя внешнего атрибута
        return None  # Если значение не найдено
    def get_inner_attribute_name_by_value(self, value):
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
        return self.name_register

"""
На каждом этапе обработки таблицы записываются в данном классе.
table - обрабатываемая таблица (регистра 1С)
register - поля регистра 1С
sign_1C - Признак 1С - УПП ('upp' в коде) или не УПП ('notupp' в коде)
table_type_connection - вспомогательная таблица Субконто - Вид связи контрагента,
    используется для заполнения пропущенных значений по полю Вид связи контрагента
    в регистрах УПП
"""
@dataclass
class TableStorage:
    table: pd.DataFrame
    register: Register1c
    sign_1C: str
    table_type_connection: pd.DataFrame = field(default=None)
    def set_index_column(self, name_attribute: str, value):
        """
        В регистре Обороты счета наименования столбцов с оборотами с корреспондирующими счетам
        не разделены по дебетовые и кредитовые. Метод записывает их индексы для разделения.
        """
        attr_name = f'index_column_{name_attribute}'
        setattr(self, attr_name, value)
