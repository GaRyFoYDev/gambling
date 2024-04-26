##############################################################################
# USIDOF Project                                                             #
#                                                                            #
# This file contains the code required to preprocess the data                #
#                                                                            #
#                                                                            #
# Author: Gary FOY                                                           #
# Date: 20-04-2024                                                           #
# Email: gary.foy.auditeur@lecnam.net                                        #
##############################################################################

import numpy as np
import pandas as pd
import openpyxl
from pathlib import Path


def getPath(name: str) -> list:
    """The function creates a list of all files in the data directory and returns.
    a list of them.

    Args:
        name (str):It refers to the directory name where the files we need are located. 

    Returns:
        list: The function return a list of the path for all file the data directory.
    """
    directory = Path(name)
    if directory.exists() and directory.is_dir():
        files = directory.rglob('*')
        data_path = [str(file) for file in files if file.is_file()]
        return data_path


def createDataset(paths: list, index: int) -> pd.DataFrame:
    """ This function creates a pandas dataframe with the csv file provided.

    Args:
        paths (list): A list containing all the paths
        index (int): The index path corresponding to the desired file.

    Returns:
        pd.DataFrame: A DataFrame containing the data from the specified CSV file.
    """
    try:
        data = pd.read_csv(paths[index], sep=';')
        return data
    except UnicodeDecodeError:
        data = pd.read_csv(paths[index], sep=';', encoding='latin-1')
        return data


if __name__ == '__main__':
    directory_name = 'data'
    data_paths = getPath(directory_name)
    data1 = createDataset(data_paths, 0)
    data2 = createDataset(data_paths, 1)
    print(data1)
