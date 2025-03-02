#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd


def replace_empty_space(df, column, missing_value):
    """ Replaces empty spaces "" with the string missing_value for
    a given dataframe column
    """
    for i in range(len(df)):
        if df.at[i, column] == "":
            df.at[i, column] = missing_value


def replace_NaN(df, column, missing_value):
    """ Replaces missing values with the string missing_value for
    a given dataframe column
    """
    df[column] = df[column].fillna(missing_value)


def handle_missing_values(df, missing_value):
    """ Replaces missing values in all the columns of the a dataframe"""
    df = df.replace(pd.NaT, missing_value)
    df = df.replace(np.nan, missing_value)
    df = df.fillna(missing_value)
    return df


def drop_rows(df, indices):
    """ Drops the rows of a dataframe with the specified indices """
    try:
        return df.drop(indices, axis=0)
    except:
        return df


def reset_indices(df):
    """ Resets the indices of a dataframe.
    Often necessary after droping rows
    """
    df.index = np.arange(0, len(df), 1)
    return df
