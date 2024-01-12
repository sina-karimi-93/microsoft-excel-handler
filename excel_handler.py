"""
This module is about working with an excel file.
Author: Sina Karimi Aliabad
year:   2023
"""


from os.path import exists
from win32com.client import Dispatch
from typing import Generator
from typing import Any
from errors import NotFoundExcelFileError
from errors import NotFoundSheetError

__version__ = "1.0"
__all__ = ["ExcelHandler"]

class ExcelHandler:
    """
    This class contains all the tools to open
    an excel file, fetch its data and update
    them.

    @methods
        open_excel
        create_new
        set_sheet
        get_columns_count
        get_rows_count
        fetch_all
        fetch_range
        get_as_dict
        update_cell
        save
        save_as
        close
    
    @note
        better to use this class as context manager.
    """
    def __init__(self,
                 dev: bool = False) -> None:
        self.excel_app = Dispatch("Excel.Application")
        if dev:
            self.excel_app.Visible = 1

    def __enter__(self)-> object:
        return self
    
    def __exit__(self, *args) -> object:
        self.close()
    
    def open_excel(self, 
                   file_path: str,
                   sheet_name: str = None) -> None:
        """
        Open an excel file if the path provided
        or create new one of file path is empty.
        """
        if not exists(file_path):
            raise NotFoundExcelFileError("Couldn't find the desired excel file.")
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
        total_sheets_count = self.work_book.Sheets.Count
        for sheet_number in range(1, total_sheets_count + 1):
            self.sheet = self.work_book.Sheets(sheet_number)
            if self.sheet.name == sheet_name:
                break
        else:
            raise NotFoundSheetError("Desired sheet is not in the file.")

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

    def fetch_all(self, 
                  rows_count: int = None,
                  columns_count: int = None) -> Generator:
        """
        Fetch all data in the data.
        ------------------------------------
        -> Params
            rows_count : int
            columns_count: int
        <- Return
            Generator
        @note
            As the end of the sheet is not specified
            and we have to find it by ourselve, user
            can specify the end last row and column
            to ignore the steps finding the last row
            and column.
        """
        if not all((rows_count, columns_count)):
            columns_count = self.get_columns_count()
            rows_count = self.get_rows_count()
        yield from self.fetch_range((1, 1), 
                                    (rows_count, columns_count))

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
        if self.excel_app.ActiveWorkbook:
            self.excel_app.ActiveWorkbook.Close()
        del self.excel_app




if __name__ == "__main__":
    from os import getcwd
    from pprint import pprint
    with ExcelHandler() as handler:
        handler: ExcelHandler
        handler.open_excel(f"{getcwd()}/data/sample-1.xls")
        data = handler.fetch_all()
        data_as_dict = tuple(handler.get_as_dict(next(data), data))
        pprint(data_as_dict)
        # range_data = tuple(handler.fetch_range((1,1), (5,7)))
        # handler.close()