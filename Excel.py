from typing import Iterable
from openpyxl import Workbook, load_workbook
from string import ascii_uppercase

class Excel:
    def __init__(self,file):
        try:
            self.file = file
            self.workbook = load_workbook(file)
        except FileNotFoundError:
            self.file = file
            self.workbook = Workbook()
            self.workbook.save(file)
    
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
    
    def excel_column_index(self,column):
        counter = 0
        for column_ in self.excel_columns():
            counter += 1
            if column_ == column:
                return counter
    
    def sheet_config(self,sheet,read=False):
        try:
            sheet = self.workbook[sheet]
        except KeyError:
            if read : raise KeyError(f"Worksheet {sheet} does not exist")
            sheets = self.workbook.sheetnames
            if 'Sheet' in sheets:
                defult = self.workbook['Sheet']
                self.workbook.remove(defult)
            self.workbook.create_sheet(sheet)
            sheet = self.workbook[sheet]
        return sheet
    
    def column_verify(self,column):
        if column.upper() not in self.excel_columns():
            raise KeyError("Column letter must be between A,B ... AA,AB and XFD")
    
    def row_verify(self,row):
        if row < 1 or row > 1048576 :
            raise KeyError("Row index must be integer between 1 and 1048576")
    
    def write_on_column(self,sheet:str,column:str,values:Iterable):
        sheet = str(sheet)
        self.column_verify(column)
        sheet = self.sheet_config(sheet)
        
        for item,index in enumerate(values):
            sheet[f'{column.upper()}{index}'] = item
        self.workbook.save(self.file)
        return True
    
    def write_on_row(self,sheet:str,row:int,values:Iterable):
        sheet = str(sheet)
        self.row_verify(row)        
        sheet = self.sheet_config(sheet)
        
        for item,column in zip(values,self.excel_columns()):
            sheet[f'{column.upper()}{row}'] = item
        self.workbook.save(self.file)
        return True
    
    def write_on_cell(self,sheet:str,column:str,row:int,value):
        sheet = str(sheet)
        self.column_verify(column)
        self.row_verify(row)        
        sheet = self.sheet_config(sheet)
        
        sheet[f'{column}{row}'] = value
        self.workbook.save(self.file)
        return True
    
    def read_column(self,sheet:str,column:str):
        sheet = str(sheet)
        self.column_verify(column)
        sheet = self.sheet_config(sheet,True)
        
        generator = sheet.values
        for row in generator:
            yield row[self.excel_column_index(column)-1]

    def read_row(self,sheet:str,row:int):
        sheet = str(sheet)
        self.row_verify(row)
        sheet = self.sheet_config(sheet,True)
        
        generator = sheet.values
        counter = 0
        for row_ in generator:
            counter += 1
            if counter == row:
                for cell in row_:
                    yield cell
    
    def read_cell(self,sheet:str,column:str,row:int):
        sheet = str(sheet)
        self.column_verify(column)
        self.row_verify(row)        
        sheet = self.sheet_config(sheet)
        
        generator = sheet.values
        column = self.excel_column_index(column)-1
        counter = 0
        for row_ in generator:
            counter += 1
            if counter == row:
                return row_[column]