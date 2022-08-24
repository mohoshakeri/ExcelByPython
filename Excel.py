from typing import Iterable
from openpyxl import Workbook, load_workbook
from string import ascii_uppercase

class Excel:
    """
    A class that helps you by its functions to work easier with Excel data
    
    params:
        file (str) : file name or file path + 'xlsx'
            Tip : if file not exist its will be create.
    """
    def __init__(self,file):
        try:
            self.file = file
            self.workbook = load_workbook(file)
        except FileNotFoundError:
            self.file = file
            self.workbook = Workbook()
            self.workbook.save(file)
    
    def excel_columns(self):
        """
        Internal generator
        
        Generate excel column : [A,B ... AA,AB ... XFE,XFD]
        """
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
        """
        Internal function
        
        Get index of columns of Excel
        """
        counter = 0
        for column_ in self.excel_columns():
            counter += 1
            if column_ == column:
                return counter
    
    def sheet_config(self,sheet,read=False):
        """
        Internal function
        
        Config sheet of workbook
        """
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
        """
        Internal function
        
        Verify column name
        """
        if column.upper() not in self.excel_columns():
            raise KeyError("Column letter must be between A,B ... AA,AB and XFD")
    
    def row_verify(self,row):
        """
        Internal function
        
        Verify row index
        """
        if row < 1 or row > 1048576 :
            raise KeyError("Row index must be integer between 1 and 1048576")
    
    def write_on_column(self,sheet:str,column:str,values:Iterable):
        """
        Write values of a Iterable on a column
        
        Args:
            sheet (str) : Sheet name. if it not exist will be create\n
            column (str) : A excel column name. Ex: AB **between A and XFD**\n
            values (iterable) : Contains values that will fill cells
        
        Return:
            True
        """
        sheet = str(sheet)
        self.column_verify(column)
        sheet = self.sheet_config(sheet)
        
        for item,index in enumerate(values):
            sheet[f'{column.upper()}{index}'] = item
        self.workbook.save(self.file)
        return True
    
    def write_on_row(self,sheet:str,row:int,values:Iterable):
        """
        Write values of a Iterable on a row
        
        Args:
            sheet (str) : Sheet name. if it not exist will be create\n
            row (int) : A excel row index. Ex: 12 **between 1 and 1048576**\n
            values (iterable) : Contains values that will fill cells
        
        Return:
            True
        """
        sheet = str(sheet)
        self.row_verify(row)        
        sheet = self.sheet_config(sheet)
        
        for item,column in zip(values,self.excel_columns()):
            sheet[f'{column.upper()}{row}'] = item
        self.workbook.save(self.file)
        return True
    
    def write_on_cell(self,sheet:str,column:str,row:int,value):
        """
        Write a value on a cell
        
        Args:
            sheet (str) : Sheet name. if it not exist will be create\n
            column (str) : A excel column name. Ex: AB **between A and XFD**\n
            row (int) : A excel row index. Ex: 12 **between 1 and 1048576**\n
            value (any except Iterables) : value that will fill cells
        
        Return:
            True
        """
        sheet = str(sheet)
        self.column_verify(column)
        self.row_verify(row)        
        sheet = self.sheet_config(sheet)
        
        sheet[f'{column}{row}'] = value
        self.workbook.save(self.file)
        return True
    
    def read_column(self,sheet:str,column:str):
        """
        Read values of a excel column
        
        Args:
            sheet (str) : Sheet name\n
            column (str) : A excel column name. Ex: AB **between A and XFD**\n
        
        Yields:
            Column values
        """
        sheet = str(sheet)
        self.column_verify(column)
        sheet = self.sheet_config(sheet,True)
        
        generator = sheet.values
        for row in generator:
            yield row[self.excel_column_index(column)-1]

    def read_row(self,sheet:str,row:int):
        """
        Read values of a excel row
        
        Args:
            sheet (str) : Sheet name\n
            row (int) : A excel row index. Ex: 12 **between 1 and 1048576**\n
        
        Yields:
            Row values
        """
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
        """
        Read value of a excel cell
        
        Args:
            sheet (str) : Sheet name\n
            column (str) : A excel column name. Ex: AB **between A and XFD**\n
            row (int) : A excel row index. Ex: 12 **between 1 and 1048576**\n
        
        Return:
            The cell value
        """
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