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

import pandas as pd
from pathlib import Path


def get_path(name: str) -> list:
    """Retrieves paths of files within the specified directory.

    Args:
        name (str): Name of the directory.

    Returns:
        list: List of file paths within the directory.
    """

    directory = Path(name)
    if directory.exists() and directory.is_dir():
        files = directory.rglob('*')
        data_path = [str(file) for file in files if file.is_file()]

    return data_path


def create_dataset(path: str) -> pd.DataFrame:
    """Creates a pandas DataFrame from the CSV file provided.

    Args:
        path (str): The path to the CSV file.

    Returns:
        pd.DataFrame: DataFrame containing the data from the specified CSV file.
    """

    try:
        data = pd.read_csv(path, sep=';')
        return data
    except UnicodeDecodeError:
        data = pd.read_csv(path, sep=';', encoding='latin-1')
        return data


def data_preprocessing(df: pd.DataFrame) -> pd.DataFrame:
    """Preprocesses the DataFrame containing gambling data.

    Args:
        df (pd.DataFrame): DataFrame containing the data.

    Returns:
        pd.DataFrame: DataFrame after preprocessing.
    """
    # Set index
    df = df.set_index('Catégorie/Année')

    # Transpose table
    df = df.T

    # Slice df to keep good columns
    df = df.iloc[:, :33]

    # Get only years from the indexes and replace
    new_indexes = pd.Index([date[-4:] for date in df.index])
    df.index = new_indexes

    df.index = df.index.astype(int)

    df = df.rename(columns={
        'Nombre dopérateurs': 'Nombre opérateurs',
        'Nombre dagréments': 'Nombre agréments',
        'Mises paris sportifs (en M)': 'Mises paris sportifs (en millions euros)',
        'PBJ paris sportifs (en M)': 'PBJ paris sportifs (en millions euros)',
        'Mises paris hippiques (en M)': 'Mises paris hippiques (en millions euros)',
        'PBJ paris hippiques (en M)': 'PBJ paris hippiques (en millions euros)',
        'Mises poker cash game (en M)': 'Mises poker cash game (en millions euros)',
        "Droits d'entrée en tournois de poker (en M)": "Droits d'entrée en tournois de poker (en millions euros)",
        'PBJ poker (en M)': 'PBJ poker (en millions euros)',
        'Budget marketing médias (en M)': 'Budget marketing médias (en millions euros)'})

    # For the columns from 0 to 17:
    # Delete the spaces between number
    # Replace NaN values
    # Convert str values to int
    for column in df.iloc[:, :18].columns:
        df[column] = df[column].str.replace(' ', '')
        df[column] = df[column].fillna(0)
        df[column] = df[column].astype(int)

    # For the columns from 18 to the end:
    # Delete the % symbols
    # Replace NaN values
    # Convert str values to float
    # Divide values per 100 to get proportions
    for column in df.iloc[:, 18:].columns:
        df[column] = df[column].str.replace('%', '')
        df[column] = df[column].fillna(0)
        df[column] = df[column].astype(float)
        df[column] = df[column]/100

    # Calcul de la moyenne des différentes colonnes et multiplication par 100
    df['Moyenne répartition mise homme (en %)'] = (df[['Part hommes PS T4',
                                                       'Part hommes PH T4',
                                                       'Part hommes PO T4']].mean(axis=1) * 100).astype(float).round(1)

    df['Moyenne répartition mise femme (en %)'] = (df[['Part femmes PS T4',
                                                       'Part femmes PH T4',
                                                       'Part femmes PO T4']].mean(axis=1) * 100).astype(float).round(1)

    df['Moyenne répartition mise smartphone et tablette (en %)'] = (df[['Part mises smartphones et tablettes PS T4',
                                                                        'Part mises smartphones et tablettes PH T4',
                                                                        'Part mises smartphones et tablettes PO T4']].mean(axis=1) * 100).astype(float).round(1)

    df['Moyenne répartition mise ordinateur (en %)'] = (df[['Part mises ordinateur PS T4',
                                                            'Part mises ordinateur PH T4',
                                                            'Part mises ordinateur PO T4']].mean(axis=1) * 100).astype(float).round(1)

    return df
