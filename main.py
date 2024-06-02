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

import dashboard as dash
import preprocessing as prep

if __name__ == '__main__':

    directory_name = 'data'
    data_path = prep.get_path(directory_name)

    gambling_df = prep.create_dataset(data_path, 0)
    clean_gambling_df = prep.data_preprocessing(gambling_df)

    dash.create_xls_file(clean_gambling_df)
