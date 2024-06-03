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

    with pd.ExcelWriter('xlsx/gambling.xlsx', engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='data', index=True)
        wb = writer.book
        dashboard = wb.add_worksheet('dashboard')

        data = wb.get_worksheet_by_name('data')

        column_length = [len(str(column)) for column in df.columns]

        for i, width in enumerate(column_length):
            data.set_column(i+1, i+1, width + 5)

        # Chart 1: Évolution des opérateurs
        chart1_series1 = {'name': '=data!$B$1',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$B$2:$B$14'}

        create_chart(wb, 'line',
                     dashboard, 'B2',
                     chart1_series1,
                     title="Évolution du nombre d'opérateurs",
                     xlabel='Année',
                     ylabel="Nombre d'opérateurs",
                     size=(720, 456))

        # Chart 2 - Répartition des agréments
        chart2_series1 = {'name': 'Paris sportifs',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$D$2:$D$14'}

        chart2_series2 = {'name': 'Paris hippiques',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$E$2:$E$14'}

        chart2_series3 = {'name': 'Paris hippiques',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$F$2:$F$14'}

        create_chart(wb, 'column',
                     dashboard, 'B28',
                     chart2_series1,
                     chart2_series2,
                     chart2_series3,
                     title="Évolution du nombre d'agrément accordé par secteur",
                     xlabel='Année',
                     ylabel="Nombre d'agréments accordés",
                     size=(720, 456),

                     )

        # Chart 3 - Répartion des CJA

        # Chart 4 - Comparaison de l'évolution des mises
        # Chart 5


def create_chart(book: xls.Workbook,
                 type: str,
                 sheet: xls.Workbook.worksheet_class,
                 insert_area: str,
                 *series: dict,
                 title: Union[None, str] = None,
                 xlabel: Union[None, str] = None,
                 ylabel: Union[None, str] = None,
                 size: Union[None, Tuple[int, int]] = None,
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

    for serie in series:
        chart.add_series(serie)

    sheet.insert_chart(insert_area, chart)

    if title:
        chart.set_title({"name": title})

    if xlabel:
        chart.set_x_axis({"name": xlabel})

    if ylabel:
        chart.set_y_axis({"name": ylabel})

    if size:
        chart.set_size({'width': size[0], 'height': size[1]})
