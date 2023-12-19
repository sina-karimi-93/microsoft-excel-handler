"""
Contains custom modules for the excel_handler.
"""

class CustomBaseException(Exception):
    """
    Custom base class for the exceptions
    """


class NotFoundExcelFileError(CustomBaseException):
    """
    Raises when couldn't find the desired file.
    """
    
class NotFoundSheetError(CustomBaseException):
    """
    Raises when couldn't find the desired sheet.
    """