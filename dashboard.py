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
from typing import Union, Tuple


def create_xls_file(df: pd.DataFrame) -> None:
    """_summary_

    Args:
        df (pd.DataFrame): _description_
    """

    with pd.ExcelWriter('gambling.xlsx', engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='data', index=True)
        wb = writer.book
        dashboard = wb.add_worksheet('dashboard')

        data = wb.get_worksheet_by_name('data')

        column_length = [len(str(column)) for column in df.columns]

        for i, width in enumerate(column_length):
            data.set_column(i+1, i+1, width + 5)
        
        

        # Chart 1
        create_chart(wb, 'line',
                     {'name': '=data!$B$1',
                      'categories': '=data!$A$2:$A$14',
                      'values': '=data!$B$2:$B$14'},
                     dashboard, 'B2',
                     title="Évolution du nombre d'opérateurs",
                     xlabel='Année',
                     ylabel="Nombre d'opérateurs",

                     size=(720, 456))
        # Chart 2
        create_chart(wb, 'line',
                     {'name': '=data!$B$1',
                      'categories': '=data!$A$2:$A$14',
                      'values': '=data!$B$2:$B$14'},
                     dashboard, 'B2',
                     title="Évolution du nombre d'opérateurs",
                     xlabel='Année',
                     ylabel="Nombre d'opérateurs",

                     size=(720, 456))

        # Chart 3
        # Chart 4
        # Chart 5


def create_chart(book: xls.Workbook,
                 type: str,
                 values: dict,
                 sheet: xls.Workbook.worksheet_class,
                 insert_area: str,
                 title=Union[None, str],
                 xlabel=Union[None, str],
                 ylabel=Union[None, str],
                 size=Union[None, Tuple[int, int]]
                 ) -> None:
    """_summary_

    Args:
        book (xls.Workbook): _description_
        type (str): _description_
        values (dict): _description_
        sheet (xls.Workbook.worksheet_class): _description_
        insert_area (str): _description_
        title (_type_, optional): _description_. Defaults to Union[None,str].
        xlabel (_type_, optional): _description_. Defaults to Union[None,str].
        ylabel (_type_, optional): _description_. Defaults to Union[None,str].
        size (_type_, optional): _description_. Defaults to Union[None,Tuple[int,int]].
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

    if size:
        chart.set_size({'width': size[0], 'height': size[1]})
