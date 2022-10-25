class Const:
    #:  number of the first line of the source workbook with data
    START_ROW = 30

    #:  number of the last line of the source workbook with data
    FINISH_ROW = 74

    #:  number of the first column of the source workbook with data
    START_COLUMN = 1

    #:  number of the last column of the source workbook with data
    FINISH_COLUMN = 17

    #:  source workbook location
    BOOK_PATH = r"d:\IT\Development\Мои примеры\Excel_Book_analysis" \
                r"\budget_analysis\files for example" \
                r"\18115 (СНЯТИЕ И УСТАНОВКА ИЗОЛЯЦИИ ВЫТЯЖНЫХ ВОЗДУХОВОДОВ ОТ ПЕЧЕЙ ПВК-1000.(УВМ))(1).xlsx"

    #:  the name of the sheet with the original data
    DATA_SHEET = 'Лист1'

    #:  results recording settings
    WRITER_SETTINGS = {
        'row_start': 6,  #: starting row to write
        'column_start': 1,  #: starting column to write
        'new_book_name': 'TEST_2',  #: the name of the upcoming workbook
        'new_sheet_name': 'Sheet'  #: the name of the upcoming sheet
    }
