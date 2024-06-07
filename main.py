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

import scripts.dashboard as dash
import scripts.preprocessing as prep

if __name__ == '__main__':
    directory_name = 'data'
    data_path = prep.get_path(directory_name)

    gambling_df = prep.create_dataset(data_path[0])
    clean_gambling_df = prep.data_preprocessing(gambling_df)

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

    clean_gambling_df = prep.compute_data(
        clean_gambling_df, cja_lst, 'mean', 'Moyenne CJA')
    clean_gambling_df = prep.compute_data(
        clean_gambling_df, mise_lst, 'sum', 'Total des mises en millions euros')
    clean_gambling_df = prep.compute_data(
        clean_gambling_df, pbj_lst, 'sum', 'Total du PBJ en millions euros')

    dash.create_xls_file(clean_gambling_df)
