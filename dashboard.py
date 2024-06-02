##############################################################################
# USIDOF Project                                                             #
#                                                                            #
# This file contains the main code to excute the script                      #
#                                                                            #
#                                                                            #
# Author: Gary FOY                                                           #
# Date: 20-04-2024                                                           #
# Email: gary.foy.auditeur@lecnam.net                                        #
##############################################################################


import pandas as pd
import xlsxwriter as xls


def create_xls_file(df: pd.DataFrame):
    """_summary_

    Args:
        df (pd.DataFrame): _description_
    """

    with pd.ExcelWriter('gambling.xlsx', engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='data', index=True)
        wb = writer.book
        sheet = wb.add_worksheet('dashboard')

        create_chart(wb, 'line',
                     {'name': '=data!$B$1',
                      'categories': '=data!$A$2:$A$14',
                      'values': '=data!$B$2:$B$14'
                      },
                     sheet, 'B2',
                     title="Evolution du nombre d'opérateurs",
                     xlabel='Année')


def create_chart(book: xls.Workbook,
                 type: str,
                 values: dict,
                 sheet: xls.Workbook.worksheet_class,
                 insert_area: str,
                 title=None,
                 xlabel=None,
                 ylabel=None
                 ):
    """_summary_

    Args:
        book (xls.Workbook): _description_
        type (str): _description_
        values (dict): _description_
        sheet (xls.Workbook.worksheet_class): _description_
        insert_area (str): _description_
        title (_type_, optional): _description_. Defaults to None.
        xlabel (_type_, optional): _description_. Defaults to None.
        ylabel (_type_, optional): _description_. Defaults to None.
    """

    chart = book.add_chart({'type': type})
    chart.add_series(values)
    sheet.insert_chart(insert_area, chart)

    if title:
        chart.set_title({"name": title})

    if xlabel:
        chart.set_x_axis({"name": xlabel})

    if ylabel:
        chart.set_y_axis({"name": ylabel})
