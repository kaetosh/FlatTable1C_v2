# -*- coding: utf-8 -*-
"""
Created on Tue Dec 17 17:01:45 2024

@author: a.karabedyan
"""

from Register1C import FieldsRegister, Register_1C

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

filds_osv_upp = FieldsRegister(analytics='Субконто',
                               type_connection = 'Вид связи КА за период', 
                               start_debit_balance = 'Нач. сальдо деб.',
                               start_credit_balance = 'Нач. сальдо кред.', 
                               debit_turnover = 'Деб. оборот',
                               credit_turnover = 'Кред. оборот', 
                               end_debit_balance = 'Кон. сальдо деб.',
                               end_credit_balance = 'Кон. сальдо кред.')

osv_upp = Register_1C(FieldsRegister_upp = filds_osv_upp)
