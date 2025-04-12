#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: Angelica Goncalves.

Module to clean the format of custom Pandas dataframes.

Functions
---------
replace_empty_space(df, column, missing_value)
    Replace empty-cell with a chosen string for a given column of the df.

handle_missing_values(df, missing_value)
    Returns df after replacing NA in all the columns of the a dataframe."

drop_rows(df, indices)
    Drop the rows of a dataframe with the specified indices.

reset_indices(df)
    Reset the indices of a dataframe.

"""

# Import standard libraries:
import numpy as np
import pandas as pd


def replace_empty_space(df, column, missing_value):
    """
    Replace empty-cell with a chosen string for a given column of the df.

    Arguments
    ---------
    df : Pandas dataframe

    columns : str
        Name of the column.

    missing_value : str
        Name to represent missing values in a dataframe.

    Returns
    -------
    df : Pandas dataframe

    """
    for i in range(len(df)):
        if df.at[i, column] == "":
            df.at[i, column] = missing_value


def handle_missing_values(df, missing_value):
    """
    Return df after replacing NA in all the columns of a dataframe.

    Arguments
    ---------
    df : Pandas dataframe

    missing_value : str
        Name to represent missing values in a dataframe

    Returns
    -------
    df : Pandas dataframe

    """
    df = df.replace(pd.NaT, missing_value)
    df = df.replace(np.nan, missing_value)
    df = df.fillna(missing_value)
    return df


def drop_rows(df, indices):
    """
    Drop the rows of a dataframe with the specified indices.

    Arguments
    ---------
    df : Pandas dataframe

    indices : list
        List of indices to be dropped from the dataframe

    Returns
    -------
    df : Pandas dataframe

    """
    try:
        return df.drop(indices, axis=0)
    except Exception:
        return df


def reset_indices(df):
    """
    Reset the indices of a dataframe.

    Arguments
    ---------
    df : Pandas dataframe

    Returns
    -------
    df : Pandas dataframe

    """
    df.index = np.arange(0, len(df), 1)
    return df
