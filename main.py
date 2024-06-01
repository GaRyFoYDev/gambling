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

from preprocessing import clean_gambling_market as df1
import pandas as pd
import xlsxwriter


def create_excel(df: pd.DataFrame):

    with pd.ExcelWriter('gambling.xlsx', engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='data', index=True)

        wb = writer.book
        # workbook.add_worksheet('dashboard_1')
        # workbook.add_worksheet('dashboard_2')


if __name__ == '__main__':
    create_excel(df1)
