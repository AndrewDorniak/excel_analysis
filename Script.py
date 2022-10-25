
"""
    This file launches the file analysis program.

    Create sort condition here.

    Application settings in a file settings.py
"""

from analisys import *
from settings import Const


def make_articles():
    """ Use this for single line estimates """
    book = Book.is_correct_path(os.path.abspath(Const.BOOK_PATH))
    book.open_book()
    book.assign_data_sheet(Const.DATA_SHEET)
    articles = []
    for row in range(Const.START_ROW, Const.FINISH_ROW):
        params_list = [book.data_sheet.cell(row=row, column=column).value for column in
                       range(Const.START_COLUMN, Const.FINISH_COLUMN)]
        article = Article(params_list)
        articles.append(article)
    book.save_book()
    book.close()
    return articles


def make_articles_standard_budget():
    """ Use this for standard estimates """
    book = Book.is_correct_path(os.path.abspath(Const.BOOK_PATH))
    book.open_book()
    book.assign_data_sheet(Const.DATA_SHEET)
    articles = []
    for row in range(Const.START_ROW, Const.FINISH_ROW):
        rate = book.data_sheet.cell(row=row, column=3).value
        if rate is not None:
            is_rate = rate[1:].replace('-', '').isdigit()
            if is_rate:
                first_row = [book.data_sheet.cell(row=row, column=column).value for column in
                             range(Const.START_COLUMN, Const.FINISH_COLUMN)]
                second_row = [book.data_sheet.cell(row=row + 1, column=column).value for column in
                              range(Const.START_COLUMN, Const.FINISH_COLUMN)]
                params_list = {'first_row': first_row,
                               'second_row': second_row}
                article = Article(params_list)
                articles.append(article)
    book.save_book()
    book.close()
    return articles


def sorting(articles, specification):
    """ sorting by custom conditions

        :param articles: List of objects of the estimate
        :type articles: list

        :param specification: Custom specification or class AndSpecification/AnySpecification object
        :type specification: Specification class derivative
     """
    wf = WorkingFilter()
    sorted_articles = []
    for art in wf.custom_filter(articles, specification):
        sorted_articles.append(art)
    return sorted_articles


if __name__ == '__main__':

    #: Create list of objects of the estimate
    articles = make_articles_standard_budget()

    #: Example of setting a condition
    class PipesOnlySpecification(Specification):
        def __init__(self, condition):
            self.condition = condition

        def is_satisfied(self, article):
            name = article.parameter_list['first_row'][3].lower()
            if self.condition in name:
                return True

    specification = PipesOnlySpecification('трубопров')
    sorted_articles = sorting(articles, specification)

    #: Writing the result to a new excel workbook
    sw = SimpleWriter(sorted_articles, **Const.WRITER_SETTINGS)
    sw.create_sheet()
    sw.write_to()
    sw.book.save_book()
    sw.book.close()
