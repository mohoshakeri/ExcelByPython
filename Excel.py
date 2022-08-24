from typing import Iterable
from openpyxl import Workbook, load_workbook
from string import ascii_uppercase

from pyrfc3339 import generate

class Excel:
    def __init__(self,file):
        try:
            self.file = file
            self.workbook = load_workbook(file)
            self.new = False
        except FileNotFoundError:
            self.file = file
            self.workbook = Workbook()
            self.workbook.save(file)
            self.new = True
    
    def excel_columns(self):
        break_check = False
        alphabet = list(ascii_uppercase)
        for item in alphabet:
            yield item
        for item_1 in alphabet:
            for item_2 in alphabet:
                yield item_1 + item_2
        for item_1 in alphabet:
            if break_check : break
            for item_2 in alphabet:
                if break_check : break
                for item_3 in alphabet:
                    if (item_1 + item_2 + item_3) == 'XFE':
                        break_check = True
                        break
                    yield item_1 + item_2 + item_3
    
    def write_on_column(self,sheet:str,column:str,values:Iterable):
        sheet = str(sheet)
        if column.upper() not in self.excel_columns():
            raise KeyError("Column letter must be between A,B ... AA,AB and XFD")
        if self.new:
            sheet_name = sheet
            sheet = self.workbook.active
            sheet.title = sheet_name
        else:
            try:
                sheet = self.workbook[sheet]
            except KeyError:
                self.workbook.create_sheet(sheet)
                sheet = self.workbook[sheet]
        for item,index in enumerate(values):
            sheet[f'{column.upper()}{index}'] = item
        self.workbook.save(self.file)
        return True
    
    def write_on_row(self,sheet:str,row:int,values:Iterable):
        sheet = str(sheet)
        if row < 1 or row > 1048576 :
            raise KeyError("Row index must be integer between 1 and 1048576")
        if self.new:
            sheet_name = sheet
            sheet = self.workbook.active
            sheet.title = sheet_name
        else:
            try:
                sheet = self.workbook[sheet]
            except KeyError:
                self.workbook.create_sheet(sheet)
                sheet = self.workbook[sheet]
        for item,column in zip(values,self.excel_columns()):
            sheet[f'{column.upper()}{row}'] = item
        self.workbook.save(self.file)
        return True
    
    def write_on_cell(self,sheet:str,column:str,row:int,value):
        sheet = str(sheet)
        if column.upper() not in self.excel_columns():
            raise KeyError("Column letter must be between A,B ... AA,AB and XFD")
        if row < 1 or row > 1048576 :
            raise KeyError("Row index must be integer between 1 and 1048576")
        if self.new:
            sheet_name = sheet
            sheet = self.workbook.active
            sheet.title = sheet_name
        else:
            try:
                sheet = self.workbook[sheet]
            except KeyError:
                self.workbook.create_sheet(sheet)
                sheet = self.workbook[sheet]
        sheet[f'{column}{row}'] = value
        self.workbook.save(self.file)
        return True

ex = Excel("hi.xlsx")
