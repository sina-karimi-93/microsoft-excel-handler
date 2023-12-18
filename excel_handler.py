"""
This module is about working with an excel file.
"""
from os import getcwd
from win32com.client import Dispatch
from pprint import pprint
from typing import Generator
from typing import Any

class ExcelHandler:

    def __init__(self,
                 dev: bool = False) -> None:
        self.excel_app = Dispatch("Excel.Application")
        if dev:
            self.excel_app.Visible = 1

    def __enter__(self)-> object:
        return self
    
    def __exit__(self, *args) -> object:
        self.close()
        return self
    
    def open_excel(self, 
                   file_path: str,
                   sheet_name: str = None) -> None:
        """
        Open an excel file if the path provided
        or create new one of file path is empty.
        """
        self.excel_app.Workbooks.open(file_path)
        self.work_book = self.excel_app.WorkBooks(1)
        self.set_sheet(sheet_name)

    def create_new(self)-> None:
        """
        Create new excel file
        """
        self.excel_app.Workbooks.Add()
        self.work_book = self.excel_app.WorkBooks(1)
        self.set_sheet()

    def set_sheet(self, sheet_name: str = None) -> None:
        """
        Set the sheet file in an excel file.
        ------------------------------------
        -> Params
            sheet_name: str
        """
        if not sheet_name:
            self.sheet = self.work_book.Sheets(1)
            return
        for sheet_number in range(1,self.work_book.Sheets.Count+1):
            self.sheet = self.work_book.Sheets(sheet_number)
            if self.sheet.name == sheet_name:
                break
        else:
            raise ValueError("Desired sheet is not in the file.")

    def get_columns_count(self) -> int:
        """
        Loop through the first row and count
        the number of columns until it faces
        empty column and stops. Then return
        the counter as number of columns.
        """
        column_count = 1
        while self.sheet.Cells(1, column_count).value:
            column_count += 1
        return column_count - 1
    
    def get_rows_count(self) -> int:
        """
        Loop through the first column and
        count the number of row until it faces
        empty row and stops. Then return
        the counter as number of row.
        """
        row_count = 1
        while self.sheet.Cells(row_count, 1).value:
            row_count += 1
        return row_count - 1

    def fetch_all(self) -> Generator:
        """
        Fetch all data in the data
        """
        columns_count = self.get_columns_count()
        rows_count = self.get_rows_count()
        for row_number in range(1, rows_count + 1):
            row_object = self.sheet.Range(self.sheet.Cells(row_number, 1),
                                          self.sheet.Cells(row_number, columns_count))
            row = row_object.value[0]
            yield row

    def fetch_range(self, start: tuple, end: tuple) -> Generator:
        """
        Fetch data from sepcific positions range
        in the file.
        ----------------------------------------
        -> Params
            start: tuple → (1, 1)
            end: tuple → (3, 4)
        """
        range_object = self.sheet.Range(self.sheet.Cells(*start),
                                        self.sheet.Cells(*end))
        yield from range_object.value

    def get_as_dict(self, 
                    headers: tuple,
                    data: Generator) -> Generator:
        """
        Returns the data as a dicts
        """
        for row in data:
            yield dict(zip(headers, row))

    def update_cell(self,
                    cell_position: tuple,
                    value: Any) -> None:
        """
        Update a cell in the active sheet
        with new value.
        ------------------------------------------
        -> Params
            cell_positions: tuple → (5, 6)
            value: anything → int, float, datetime, string
        """
        self.sheet.Cells(*cell_position).value = value

    def save(self) -> None:
        """
        Save current open file.
        """
        self.work_book.Save()
    
    def save_as(self, file_path: str) -> None:
        """
        Save current work book to new file.
        """
        self.work_book.SaveAs(file_path)

    def close(self)-> None:
        """
        Close the app
        """
        # self.excel_app.ActiveWorkbook.Close()

if __name__ == "__main__":
    with ExcelHandler(dev=True) as handler:
        handler.open_excel(f"{getcwd()}/data/sample-1.xls")
        data = handler.fetch_all()
        data_as_dict = tuple(handler.get_as_dict(next(data), data))
        # range_data = tuple(handler.fetch_range((1,1), (5,7)))
        # handler.close()