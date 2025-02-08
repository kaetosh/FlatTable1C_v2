"""
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
import inspect
from enum import Enum
from typing import Optional, Literal, Type, List, cast
from dataclasses import dataclass, field
import pandas as pd
from additional.ErrorClasses import NoExcelFilesError

def _get_class_attributes(cl: Type) -> List[str]:
    """
    Возвращает атрибуты класса (не экземпляра)
    """
    init_signature = inspect.signature(cl.__init__)
    return [i for i in init_signature.parameters if i != 'self']

# поля регистров 1с
class FieldsRegister:
    def __init__(self,
                 analytics: Optional[str] = None, # субконто счета
                 quantity: Optional[str] = None, # количество
                 type_connection: Optional[str] = None, # вид связи контрагента (признак компании группы)
                 corresponding_account: Optional[str] = None, # корреспондирующий счет
                 start_debit_balance: Optional[str] = None, # входящий дебетовый остаток
                 start_credit_balance: Optional[str] = None, # входящий кредитовый остаток
                 debit_turnover: Optional[str] = None, # дебетовый оборот
                 credit_turnover: Optional[str] = None, # кредитовый оборот
                 end_debit_balance: Optional[str] = None, # исходящий дебетовый остаток
                 end_credit_balance: Optional[str] = None, # исходящий кредитовый остаток
                 version_1c_id: Optional[str] = None): # поле для различения версии 1С (для осв и оборотов совпадает с параметром analytics, анализа счета - corresponding_account)
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
        self.version_1c_id = version_1c_id
        self.start_debit_balance_for_rename = 'Дебет_начало' # входящий дебетовый остаток для единообразия имен полей
        self.start_quantity_debit_balance_for_rename = 'Количество_Дебет_начало'
        self.start_credit_balance_for_rename = 'Кредит_начало' # входящий кредитовый остаток для единообразия имен полей
        self.start_quantity_credit_balance_for_rename = 'Количество_Кредит_начало'
        self.debit_turnover_for_rename = 'Дебет_оборот' # дебетовый оборот для единообразия имен полей
        self.debit_quantity_turnover_for_rename = 'Количество_Дебет_оборот'
        self.credit_turnover_for_rename = 'Кредит_оборот' # кредитовый оборот для единообразия имен полей
        self.credit_quantity_turnover_for_rename = 'Количество_Кредит_оборот'
        self.end_debit_balance_for_rename = 'Дебет_конец' # исходящий дебетовый остаток для единообразия имен полей
        self.end_quantity_debit_balance_for_rename = 'Количество_Дебет_конец'
        self.end_credit_balance_for_rename = 'Кредит_конец' # исходящий кредитовый остаток для единообразия имен полей
        self.end_quantity_credit_balance_for_rename = 'Количество_Кредит_конец'
        self.start_balance_before_processing = 'Сальдо_начало_до_обработки'
        self.turnover_before_processing = 'Оборот_до_обработки'
        self.end_balance_before_processing = 'Сальдо_конец_до_обработки'
        self.start_balance_after_processing = 'Сальдо_начало_после_обработки'
        self.turnover_after_processing = 'Оборот_после_обработки'
        self.end_balance_after_processing = 'Сальдо_конец_после_обработки' 
        self.start_balance_deviation = 'Сальдо_начало_разница'
        self.turnover_deviation = 'Оборот_разница'
        self.end_balance_deviation = 'Сальдо_конец_разница'
        self.file_name = 'Исх.файл' 
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
                     self.end_credit_balance,
                     self.version_1c_id))
    def get_attributes_by_suffix(self, suffix: Literal['_before_processing',
                                                       '_for_rename',
                                                       '_after_processing',
                                                       '_deviation']) -> List[str]:
        """
        Получить список имен полей по концовке имени аттрибута.
        """
        attributes = [
            'start_debit_balance_for_rename',
            'start_quantity_debit_balance_for_rename',
            'start_credit_balance_for_rename',
            'start_quantity_credit_balance_for_rename',
            'debit_turnover_for_rename',
            'debit_quantity_turnover_for_rename',
            'credit_turnover_for_rename',
            'credit_quantity_turnover_for_rename',
            'end_debit_balance_for_rename',
            'end_quantity_debit_balance_for_rename',
            'end_credit_balance_for_rename',
            'end_quantity_credit_balance_for_rename',
            'start_balance_before_processing',
            'turnover_before_processing',
            'end_balance_before_processing',
            'start_balance_after_processing',
            'turnover_after_processing',
            'end_balance_after_processing',
            'start_balance_deviation',
            'turnover_deviation',
            'end_balance_deviation'
        ]
        #return [getattr(self, attr) for attr in dir(self) if attr.endswith(suffix) and not attr.startswith('__')]
        return [getattr(self, attr) for attr in attributes if attr.endswith(suffix) and not attr.startswith('__')]


# перечисление аттрибутов FieldsRegister для аннотации метода get_inner_attribute_name_by_value()
list_of_attributes_FieldsRegister = Enum('list_of_attributes_FieldsRegister',
                                         _get_class_attributes(FieldsRegister))

class Register1c:
    def __init__(self,
                 name_register: Optional[str], # осв, обороты счета или анализ счета
                 upp: Optional[FieldsRegister] = None, # поля регистра из 1С УПП
                 notupp: Optional[FieldsRegister] = None): # поля регистра из 1С прочих версий
        self.name_register = name_register
        self.upp = upp if upp is not None else []
        self.notupp = notupp if notupp is not None else []
    def get_outer_attribute_name_by_value(self, value: str) -> Literal["upp", "notupp"]:
        """
        Метод, который ищет атрибут внешнего класса Register1c по значению
        атрибута внутреннего класса FieldsRegister
        (признак версии 1С: УПП или нет).
        """
        for attr_name, attr_value in vars(self).items():
            if isinstance(attr_value, FieldsRegister):
                for inner_attr_name, inner_attr_value in vars(attr_value).items():
                    if inner_attr_value == value and attr_name in ["upp", "notupp"]:
                        return cast(Literal["upp", "notupp"], attr_name)
        raise NoExcelFilesError
    def get_inner_attribute_name_by_value(self, value: str) -> list_of_attributes_FieldsRegister:
        """
        Метод, который ищет атрибут внутреннего класса FieldsRegister по значению
        атрибута внутреннего класса FieldsRegister.
        """
        for attr_value in (self.upp, self.notupp):
            if isinstance(attr_value, FieldsRegister):
                for inner_attr_name, inner_attr_value in vars(attr_value).items():
                    if inner_attr_value == value and inner_attr_name in [i.name for i in list_of_attributes_FieldsRegister]:
                        return cast(list_of_attributes_FieldsRegister, inner_attr_name)
        raise NoExcelFilesError
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
    sign_1C: Literal["upp", "notupp"]
    table_type_connection: pd.DataFrame = field(default=None)
    table_for_check: pd.DataFrame = field(default=None)
    file_name: str = field(default='NoName')
    def set_index_column(self, name_attribute: str, value: int) -> None:
        """
        В регистре Обороты счета наименования столбцов с оборотами с корреспондирующими счетам
        не разделены по дебетовые и кредитовые (один и тот же счет может быть и по Дт, и по Кт).
        Метод записывает их индексы для разделения.
        """
        attr_name = f'index_column_{name_attribute}'
        setattr(self, attr_name, value)
