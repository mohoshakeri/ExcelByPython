from typing import Iterable
from openpyxl import Workbook, load_workbook
from string import ascii_uppercase

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
    
    def excel_alphabet(self):
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
        if column.upper() not in self.excel_alphabet():
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
