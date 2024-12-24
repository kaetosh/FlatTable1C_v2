# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 15:26:33 2024

@author: a.karabedyan
"""

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
        self.start_debit_balance = start_debit_balance
        self.start_credit_balance = start_credit_balance
        self.debit_turnover = debit_turnover
        self.credit_turnover = credit_turnover
        self.end_debit_balance = end_debit_balance
        self.end_credit_balance = end_credit_balance
    def __iter__(self):
        return iter((self.analytics, self.quantity, self.type_connection,
                     self.start_debit_balance, self.start_credit_balance,
                     self.debit_turnover, self.credit_turnover,
                     self.end_debit_balance, self.end_credit_balance))

class Register_1C:
    def __init__(self,
                 upp: FieldsRegister or None = None,
                 notupp: FieldsRegister or None = None):
        self.upp = upp if upp is not None else []
        self.notupp = notupp if notupp is not None else []
    def __iter__(self):
        yield from self.upp
        yield from self.notupp
    
