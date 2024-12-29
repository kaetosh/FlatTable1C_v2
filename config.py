# -*- coding: utf-8 -*-
"""
Created on Tue Dec 17 17:01:45 2024

@author: a.karabedyan
"""

from Register1C import FieldsRegister, Register1c

name_account_balance_movements = {'start_debit_balance':['Начальное сальдо Дт', 'Нач. сальдо деб.'],
                                  'start_credit_balance': ['Начальное сальдо Кт', 'Нач. сальдо кред.'],
                                  'debit_turnover': ['Оборот Дт','Деб. оборот'],
                                  'credit_turnover': ['Оборот Кт', 'Кред. оборот'],
                                  'end_debit_balance':['Конечное сальдо Дт', 'Кон. сальдо деб.'],
                                  'end_credit_balance': ['Конечное сальдо Кт', 'Кон. сальдо кред.']}

# Признак версии 1С (УПП или нет)
sign_1c_upp = 'Субконто'
sign_1c_not_upp = 'Счет'

new_names = ['Дебет_начало',
             'Кредит_начало',
             'Дебет_оборот',
             'Кредит_оборот',
             'Дебет_конец',
             'Кредит_конец']

exclude_values = ['Нач.сальдо',
                  'Оборот',
                  'Итого оборот',
                  'Кон.сальдо',
                  'Начальное сальдо',
                  'Конечное сальдо',
                  'Кор. Субконто1',
                  'Кол-во:']

# счета, обороты по которым в свод попадут без субсчетов и аналитики
accounts_without_subaccount = ['55', '57']

# поля ОСВ
fields_osv_upp = FieldsRegister(analytics='Субконто',
                                type_connection = 'Вид связи КА за период',
                                start_debit_balance = 'Сальдо на начало периода',
                                debit_turnover = 'Оборот за период',
                                end_debit_balance = 'Сальдо на конец периода')

fields_osv_notupp = FieldsRegister(analytics='Счет',
                                   quantity='Показа-\nтели',
                                   start_debit_balance = 'Сальдо на начало периода',
                                   debit_turnover = 'Обороты за период',
                                   end_debit_balance = 'Сальдо на конец периода	')

osv_fields = Register1c('osv', fields_osv_upp, fields_osv_notupp)


# поля Обороты счета
fields_turnover_upp = FieldsRegister(analytics='Субконто',
                                     type_connection = 'Вид связи КА за период',
                                     start_debit_balance = 'Нач. сальдо деб.',
                                     start_credit_balance = 'Нач. сальдо кред.',
                                     debit_turnover = 'Деб. оборот',
                                     credit_turnover = 'Кред. оборот',
                                     end_debit_balance = 'Кон. сальдо деб.',
                                     end_credit_balance = 'Кон. сальдо кред.')

fields_turnover_notupp = FieldsRegister(analytics='Счет',
                                        quantity='Показа-\nтели',
                                        start_debit_balance = 'Начальное сальдо Дт',
                                        start_credit_balance = 'Начальное сальдо Кт',
                                        debit_turnover = 'Оборот Дт',
                                        credit_turnover = 'Оборот Кт',
                                        end_debit_balance = 'Конечное сальдо Дт',
                                        end_credit_balance = 'Конечное сальдо Кт')

turnover_fields = Register1c('turnover', fields_turnover_upp, fields_turnover_notupp)

# поля Анализ счета
fields_analysis_upp = FieldsRegister(analytics='Счет',
                                     type_connection = 'Вид связи КА за период',
                                     corresponding_account = 'Кор.счет',
                                     debit_turnover = 'С кред. счетов',
                                     credit_turnover = 'В дебет счетов')

fields_analysis_notupp = FieldsRegister(analytics='Счет',
                                        corresponding_account = 'Кор. Счет',
                                        quantity='Показа-\nтели',
                                        debit_turnover = 'Дебет',
                                        credit_turnover = 'Кредит')

analysis_fields = Register1c('analysis', fields_analysis_upp, fields_analysis_notupp)
