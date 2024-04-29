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


def preprocess_gambling_market_dataset(df: pd.DataFrame) -> pd.DataFrame:
    
    #Set index
    df = df.set_index('Catégorie/Année')
    
    #Transpose table
    df = df.T
    
    #Slice df to keep good columns
    df = df.iloc[:,:33]
    
    #Get only years from the indexes and replace 
    new_indexes = [date[-4:] for date in df.index ]
    df.index = new_indexes
    
    #Rename columns
    df = df.rename(columns = {
    'Nombre dopérateurs': 'Nombre opérateurs',
    'Nombre dagréments': 'Nombre agréments',
    'Mises paris sportifs (en M)': 'Mises paris sportifs (en millions euros)',
    'PBJ paris sportifs (en M)': 'PBJ paris sportifs (en millions euros)',
    'Mises paris hippiques (en M)' :'Mises paris hippiques (en millions euros)',
    'PBJ paris hippiques (en M)': 'PBJ paris hippiques (en millions euros)' ,
    'Mises poker cash game (en M)': 'Mises poker cash game (en millions euros)',
    "Droits d'entrée en tournois de poker (en M)": "Droits d'entrée en tournois de poker (en millions euros)",
    'PBJ poker (en M)': 'PBJ poker (en millions euros)',
    'Budget marketing médias (en M)': 'Budget marketing médias (en millions euros)'})
    
    #Fill missing values
    df = df.fillna(0)
    
    return df
    
    
