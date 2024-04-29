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

import preprocessing as p


if __name__ == '__main__':
    directory_name = 'data'
    data_paths = p.getPath(directory_name)
    most_bet_on_matches = p.createDataset(data_paths, 1)

    gambling_market = p.createDataset(data_paths, 0)
    clean_gambling_market = p.preprocess_gambling_market_dataset(
        gambling_market)
    print(clean_gambling_market)
