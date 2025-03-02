#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb 23 10:59:17 2025

@author: agds
"""
import pandas as pd
import sys
import os
import argparse
import importlib

parent_dir = os.path.dirname(os.getcwd())
sys.path.append(parent_dir)
sys.path.append(os.path.join(parent_dir, 'src'))
import excel
import clean


def add_channel_info(channel_path, channel_df):
    """ Get name of channel, time interval of the export and relative number
    of reports.
    """
    channel_name = "_".join(channel_path.split("/")[-1].split(".")[0].split('_')[:-3])
    channel_date = " ".join(channel_path.split('/')[-1].split(".")[0].split('_')[-3:])

    df_ = channel_df['projects_parsed'].astype('string')
    reports_in_channel = f"{len(df_[df_ != '0'])}/{len(channel_df)}"

    channel_df['channel'] = [channel_name]*len(channel_df)
    channel_df['export_dates'] = [channel_date]*len(channel_df)
    channel_df['parsed_reports_in_channel'] = [reports_in_channel]*len(channel_df)

    return channel_df


def add_info_of_users_reports(channel_df):
    """ Get latest report date and number of messages in the given channel for
    all the users in the channel.
    """
    users = channel_df['user'].unique()
    for user in users:
        user_df = channel_df[channel_df['user'] == user].sort_values(
            by='msg_date', inplace=False, ignore_index=True
            )
        latest_report_date = user_df['msg_date'].to_list()[-1]
        user_df['latest_report_date'] = [latest_report_date]*len(user_df)
        user_df['number_msgs_in_channel'] = len(user_df)
        if user == users[0]:
            channel_df_ = user_df.copy()
        else:
            channel_df_ = pd.concat(
                [channel_df_, user_df], axis=0, ignore_index=False
                )

    return channel_df_


def apply_excel_adjustments(file_path, settings_mod):
    """ Defines the sequence of changes to be done in the Excel file
    given the user's inputs in the module settings_mod.
    """
    xl = excel.ExcelFormat(file_path)
    xl.set_cell_width(settings_mod.column_widths)
    xl.set_allignment('top')
    xl.format_first_row(
            settings_mod.height_1strow,
            settings_mod.alignment_vert_1strow,
            settings_mod.alignment_horiz_1strow,
            settings_mod.font_size_1strow,
            settings_mod.font_bold_1strow,
            settings_mod.cell_color_1strow
            )
    for cc in settings_mod.font_color_in_column:
        xl.set_font_color_in_column(cc)
    for highlight in settings_mod.highlights:
        xl.format_highlight(highlight)
    for column in settings_mod.text_type_cols:
        xl.format_text_cells(column)
    xl.save_changes()


if __name__ == '__main__':

    # --Define argument parser routine:
    parser = argparse.ArgumentParser(
        description = "Script that compile all the weekly reports from the individual Excel files for each channel."
        )
    parser.add_argument("--settings_file_path", required=True, type=str)
    args = parser.parse_args()
    settings_file_path = args.settings_file_path
    print(f"settings_file_path = {settings_file_path}")

    # --Import settings module:
    parent_path = os.path.dirname(settings_file_path)
    module_name = os.path.basename(settings_file_path).split(".")[0]
    sys.path.append(parent_path)
    settings_module = importlib.import_module(module_name)

    missing_value = settings_module.missing_value
    path_converted = settings_module.excel_channels_path
    compilation_reports_file_name = settings_module.compilation_reports_file_name
    compilation_reports_path = settings_module.compilation_reports_path
    print(f"Module {module_name} imported.")
    
    # --Build dataframe from all the channels:
    for file in os.listdir(path_converted):
        channel_path = f"{path_converted}/{file}"
        if "/.~lock." in channel_path:  # avoid hidden files, if any.
            continue
        else:
            channel_df = pd.read_excel(channel_path, engine='openpyxl')

            channel_df = add_channel_info(channel_path, channel_df)

            channel_df = add_info_of_users_reports(channel_df)

            # --Handle missing values:
            channel_df = clean.handle_missing_values(channel_df, missing_value)

            # --Reorder columns:
            channel_df = channel_df[settings_module.columns_order]

            # --Concatanate channel_df to final dataframe:
            if file == os.listdir(path_converted)[0]:
                df = channel_df.copy()
            else:
                df = pd.concat([df, channel_df], axis=0, ignore_index=False)
    print('Information of all the check-in reports collected.')

    # --Sort columns:
    df.sort_values(
        by=['channel', 'display_name', 'msg_date'],
        inplace=True, ignore_index=True
        )

    # --Set columns types:
    df['projects_parsed'] = df['projects_parsed'].astype('string')
    df['keywords_parsed'] = df['keywords_parsed'].astype('string')

    # --Select rows with parsed projects:
    df = df[df['projects_parsed'] != '0']
    df = df.reset_index().drop(columns=['index'])
    print('Performed minor formatting of columns and rows.')

    # --Save Excel file:
    path = f"{compilation_reports_path}/{compilation_reports_file_name}"
    df.to_excel(path, index=False)
    apply_excel_adjustments(path, settings_module)
    print('Excel file saved.')
