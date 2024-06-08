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
import openpyxl as op
from typing import Union, Tuple

chart1_series1 = {'name': '=data!$B$1',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$B$2:$B$14',
                          'marker': {'type': 'diamond'}}

chart2_series1 = {'name': 'Paris sportifs',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$D$2:$D$14'
                  }

chart2_series2 = {'name': 'Paris hippiques',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$E$2:$E$14'}

chart2_series3 = {'name': 'Poker',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$F$2:$F$14'}

chart3_series1 = {'name': 'Paris sportifs',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$H$2:$H$14'}

chart3_series2 = {'name': 'Paris hippiques',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$I$2:$I$14'}

chart3_series3 = {'name': 'Poker',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$J$2:$J$14'}

chart4_series1 = {'name': 'Paris sportifs',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$L$2:$L$14',
                          'marker': {'type': 'diamond'}}

chart4_series2 = {'name': 'Paris hippiques',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$N$2:$N$14',
                          'marker': {'type': 'diamond'}}

chart4_series3 = {'name': 'Poker',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$P$2:$P$14',
                          'marker': {'type': 'diamond'}}

chart5_series1 = {'name': 'Paris sportifs',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$M$2:$M$14',
                          'marker': {'type': 'diamond'}}

chart5_series2 = {'name': 'Paris hippiques',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!$O$2:$O$14',
                          'marker': {'type': 'diamond'}}

chart5_series3 = {'name': 'Poker',
                          'categories': '=data!$A$2:$A$14',
                          'values': '=data!R$2:$R$14',
                          'marker': {'type': 'diamond'}}

chart6_series1 = {
    'categories': '=dashboard!$A$500:$A$503',
    'values': '=dashboard!B$500:$B$503',
    'data_labels': {'percentage': True,
                    'font': {'size': 14},
                    'position': 'outside_end'}}

chart7_series1 = {'name': 'Hommes',
                  'categories': '=data!$A$2:$A$14',
                  'values': '=data!$AI$2:$AI$14',
                  }
chart7_series2 = {'name': 'Femmes',
                  'categories': '=data!$A$2:$A$14',
                          'values': '=data!$AJ$2:$AJ$14'}

chart8_series1 = {'name': 'Smartphones & Tablettes',
                  'categories': '=data!$A$2:$A$14',
                  'values': '=data!$AK$2:$AK$14',
                  }
chart8_series2 = {'name': 'Ordinateurs',
                  'categories': '=data!$A$2:$A$14',
                          'values': '=data!$AL$2:$AL$14'}


cja_lst = ['Nombre CJA Paris sportifs',
           'Nombre CJA Paris hippiques',
           'Nombre CJA Poker']

mise_lst = ['Mises paris sportifs (en millions euros)',
            'Mises paris hippiques (en millions euros)',
            'Mises poker cash game (en millions euros)',
            "Droits d'entrée en tournois de poker (en millions euros)"]

pbj_lst = ['PBJ paris sportifs (en millions euros)',
           'PBJ paris hippiques (en millions euros)',
           'PBJ poker (en millions euros)']

sports_lst = ['Part mises football',
              'Part mises tennis',
              'Part mises basketball']


#####################################################################################


def create_chart(book: xls.Workbook,
                 type: str,
                 sheet: xls.worksheet.Worksheet,
                 insert_area: str,
                 *series: dict,
                 title: Union[None, str] = None,
                 xlabel: Union[None, str] = None,
                 ylabel: Union[None, str] = None,
                 size: Union[None, Tuple[int, int]] = None,
                 subtype: Union[None, str] = None,
                 ) -> None:
    """_summary_

    Args:
        book (xls.Workbook): _description_
        type (str): _description_
        sheet (xls.Workbook.worksheet_class): _description_
        insert_area (str): _description_
        title (Union[None, str], optional): _description_. Defaults to None.
        xlabel (Union[None, str], optional): _description_. Defaults to None.
        ylabel (Union[None, str], optional): _description_. Defaults to None.
        size (Union[None, Tuple[int, int]], optional): _description_. Defaults to None.
    """

    chart = book.add_chart({'type': type, 'subtype': subtype})

    for serie in series:
        chart.add_series(serie)

    sheet.insert_chart(insert_area, chart)

    if title:
        chart.set_title({"name": title})

    if xlabel:
        chart.set_x_axis({"name": xlabel,
                          'min': 0,
                          'num_font':  {'rotation': 45,
                                        'size': 13},
                          'name_font':  {'size': 14}
                          })

    if ylabel:
        chart.set_y_axis({"name": ylabel,
                          'min': 0,
                          'num_font':  {'size': 13},
                          'name_font':  {'size': 14}
                          })

    if size:
        chart.set_size({'width': size[0], 'height': size[1]})

    chart.set_legend({
        'font': {'size': 12}})


def autosize(df: pd.DataFrame, sheet: xls.worksheet.Worksheet) -> None:

    column_length = [len(str(column)) for column in df.columns]
    for i, width in enumerate(column_length):
        sheet.set_column(i+1, i+1, width + 5)


def compute_data(df: pd.DataFrame, columns: list, compute: str = 'sum') -> pd.Series:
    if compute == 'mean':
        res = df[columns].mean(axis=1).astype(int)
        return res
    elif compute == 'sum':
        res = df[columns].mean(axis=1).astype(int)
        return res
    elif compute == 'mean_col':
        res = df[columns].mean().round(2)
        return res
    else:
        raise ValueError('Please type in a valid value: mean or sum')


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

        # Column must adapt to value length
        autosize(df, data)

        # Dashboard
        # Chart 1
        create_chart(wb, 'line',
                     dashboard, 'B2',
                     chart1_series1,
                     title="Évolution du nombre d'opérateurs",
                     xlabel='Année',
                     ylabel="Nombre d'opérateurs",
                     size=(720, 456))

        # Chart 2
        create_chart(wb, 'column',
                     dashboard, 'F2',
                     chart2_series1,
                     chart2_series2,
                     chart2_series3,
                     title="Évolution du nombre d'agrément accordé par secteur",
                     xlabel='Année',
                     ylabel="Nombre d'agréments accordés",
                     size=(720, 456),

                     )

        # Chart 3
        create_chart(wb, 'column',
                     dashboard, 'J2',
                     chart3_series1,
                     chart3_series2,
                     chart3_series3,
                     title="Évolution du nombre de CJA par secteur",
                     xlabel='Année',
                     ylabel="Nombre de CJA",
                     size=(720, 456))

        # Chart 4
        create_chart(wb, 'line',
                     dashboard, 'B28',
                     chart4_series1,
                     chart4_series2,
                     chart4_series3,
                     title="Évolution des mises par secteur",
                     xlabel='Année',
                     ylabel="Montant en M€",
                     size=(720, 456))

        # Chart 5
        create_chart(wb, 'line',
                     dashboard, 'F28',
                     chart5_series1,
                     chart5_series2,
                     chart5_series3,
                     title="Évolution du prduit brut des jeux (PBJ)",
                     xlabel='Année',
                     ylabel="Montant en M€",
                     size=(720, 456))

        # Computing data we need
        pd_series1 = compute_data(df, cja_lst, 'mean')
        pd_series2 = compute_data(df, mise_lst)
        pd_series3 = compute_data(df, pbj_lst)

        # Getting PBJ/Mise ratio
        pd_series4 = pd_series3 / pd_series2 * 100
        pd_series4 = pd_series4.astype(float).round(1)

        # Creating df from 4 series
        df2 = pd.concat(
            [pd_series1, pd_series2, pd_series3, pd_series4], axis=1)

        # Converting our df to a matrix
        df2_lst = []
        for index, row in df2.iterrows():
            line = [index, row[0], row[1], row[2], row[3]]
            df2_lst.append(line)

        # Creating a table from our list
        dashboard.add_table('B54:E66', {'data': df2_lst,
                                        'columns': [{'header': 'Année'},
                                                    {'header': 'Moyenne CJA'},
                                                    {'header': 'Total des mises (en millions euros)'},
                                                    {'header': 'Total du PBJ (en millions euros)'},
                                                    {'header': 'PBJ / Mise (en %)'}],
                                        'autofilter': False})

        # Getting sport repartition
        pd_series5 = compute_data(df, sports_lst, 'mean_col')
        pd_series5['Autres sports'] = 1 - pd_series5.sum()

        dashboard.write_column("A500", list(pd_series5.index))
        dashboard.write_column("B500", list(pd_series5.values))

        # Chart 6
        create_chart(wb, 'pie',
                     dashboard, 'J28',
                     chart6_series1,
                     title="Répartition des mises en fonction des sports",
                     size=(720, 456))
        # Chart 7
        create_chart(wb, 'bar',
                     dashboard, 'B69',
                     chart7_series1,
                     chart7_series2,
                     title="Répartition mise hommes vs femmes",
                     xlabel='Année',
                     ylabel="Proportion (en %)",
                     size=(720, 456),
                     subtype='percent_stacked')
        # Chart 8
        create_chart(wb, 'bar',
                     dashboard, 'F69',
                     chart8_series1,
                     chart8_series2,
                     title="Répartition mise smartphone/tablette vs ordinateur",
                     xlabel='Année',
                     ylabel="Proportion (en %)",
                     size=(720, 456),
                     subtype='percent_stacked')
       

        # Autosize the dashboard
        autosize(df, dashboard)
