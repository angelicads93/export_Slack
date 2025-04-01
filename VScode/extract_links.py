#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb 23 10:59:17 2025
@author: Angelica Goncalves
"""
import pandas as pd
import sys
import os
import argparse

parent_dir = os.path.dirname(os.getcwd())
sys.path.append(parent_dir)
sys.path.append(os.path.join(parent_dir, 'src'))
import excel
import clean
import parser


def parse_command_input():
    """ Parse the user's input command. """
    parser = argparse.ArgumentParser(
        description="Python script to compile all the weekly reports from "
        + "individual Excel files."
        )
    parser.add_argument("--settings_file_path", required=True, type=str)
    args = parser.parse_args()
    settings_file_path = args.settings_file_path
    print(f"settings_file_path = {settings_file_path}")
    if os.path.exists(settings_file_path) is False:
        print(f"ERROR: Path {settings_file_path} does not exists." + "\n"
              + "       Please review your input for the argument "
              + "--settings_file_path.")
        sys.exit()
    return settings_file_path


def check_input(file_name, compilation_reports_path, excel_channels_path):
    """ Verify validity of the user's inputs in the settings module. """
    if os.path.exists(compilation_reports_path) is False:
        print(f"ERROR: Path {compilation_reports_path} does not exists." + "\n"
              + "       Please review your input for the variable"
              + " 'compilation_reports_path' in {file_name}.")
        sys.exit()
    if os.path.exists((excel_channels_path)) is False:
        print(f"ERROR: Path {excel_channels_path} does not exists." + "\n"
              + "       Please review your input for the variable "
              + "'excel_channels_path' in {file_name}.")
        sys.exit()


def get_list_channels(source_path):
    """ Retrieve the name of all the expected Slack channels from the
    Slack-export source directory.
    """
    channels_path = os.listdir(source_path)
    for i, ch in enumerate(channels_path):
        ch = ch.replace(" ", "").replace("-", "").replace("_", "")
        channels_path[i] = ch
    return channels_path


def check_channel(file_name, list_channels):
    """ Check if file_name is indeed an expected Slack channel. """
    cn = "_".join(file_name.split("_")[:-3])
    cn = cn.replace(" ", "").replace("-", "").replace("_", "")
    if cn in list_channels:
        return True
    else:
        return False


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


def format_msgs_with_urls(df, settings):
    """ Select the messages that contain urls in their text. """
    df_p = df.copy()
    df_p = df_p[df_p['URL(s)'] != settings.get('missing_value')]
    df_p = df_p.reset_index().drop(columns=['index'])
    df_p.sort_values(by=['channel', 'display_name', 'msg_date'],
                     inplace=True, ignore_index=True)
    return df_p


def filter_urls(url_list, settings):
    out = []
    for url in url_list:

        for url_exp in settings.get('urls_to_show'):
            if url_exp in url:
                out.append(url.lstrip(' ').rstrip(' '))

    return out


def select_desired_urls(df, settings):
    indices2drop = []
    for i in range(len(df)):

        urls = df.at[i, "URL(s)"].split("; ")
        filtered_urls = filter_urls(urls, settings)
        if filtered_urls == []:
            indices2drop.append(i)
        else:
            df.at[i, "URL(s)"] = '; '.join(filtered_urls)

    indices2drop = list(set(indices2drop))
    return df.drop(indices2drop, axis=0, inplace=False)


def apply_excel_adjustments(file_path, sheet_name, settings):
    """ Defines the sequence of changes to be done in the Excel file
    given the user's inputs in the module settings.
    """
    xl = excel.ExcelFormat(file_path, settings)
    ws_channel = xl.get_sheet(sheet_name)

    xl.set_cell_width(ws_channel, settings.get('column_widths'))
    xl.set_allignment(ws_channel, 'top')
    xl.format_first_row(ws_channel, settings.get('header_row'))

    xl.set_filters(ws_channel)
    xl.save_changes()


# -----------------------------------------------------------------------------
if __name__ == '__main__':
    # --Parse the settings file:
    print("Checking arguments of Python command")
    settings_file_path = parse_command_input()

    print("Parsing information in settings file")
    file_name = os.path.basename(settings_file_path)
    sett_urls = parser.Parser(os.path.abspath(settings_file_path))

    # --Check the input and retrieve expected name of channels:
    print("Checking validity of input paths")
    check_input(file_name, 
                sett_urls.get('compilation_reports_path'),
                sett_urls.get('excel_channels_path'))
    expected_channels = get_list_channels(sett_urls.get('jsons_source_path'))

    # --Build dataframe from all the channels:
    for file in os.listdir(sett_urls.get('excel_channels_path')):
        file_name = str(file).split(".")[0]
        channel_path = f"{sett_urls.get('excel_channels_path')}/{file}"

        if check_channel(file_name, expected_channels) is True:
            channel_df = pd.read_excel(channel_path, engine='openpyxl',
                                       sheet_name='Relevant messages')

            if len(channel_df) > 0:
                # --Add channels info:
                channel_df = add_channel_info(channel_path, channel_df)

                # --Handle missing values:
                channel_df = clean.handle_missing_values(channel_df,
                                                         sett_urls.get('missing_value')
                                                         )
                # --Reorder columns:
                channel_df = channel_df[sett_urls.get('columns_order')]

                # --Concatanate channel_df to final dataframe:
                if file == os.listdir(sett_urls.get('excel_channels_path'))[0]:
                    df = channel_df.copy()
                else:
                    df = pd.concat([df, channel_df], axis=0, ignore_index=False)
    clean.reset_indices(df)
    print('Information of all the check-in reports collected.')

    # --Select rows with URL(s) in their messages:
    df = format_msgs_with_urls(df, sett_urls)

    # --Filter rows with desire URLs:
    df_selection = select_desired_urls(df, sett_urls)

    # --Save Excel workbook:
    path = f"{sett_urls.get('compilation_reports_path')}/{sett_urls.get('compilation_urls_file_name')}"
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df_selection.to_excel(writer, sheet_name='Selected URLs', index=False)
        df.to_excel(writer, sheet_name='All URLs', index=False)

    # --Apply formatting of Excel worksheets:
    apply_excel_adjustments(path, 'Selected URLs', sett_urls)
    apply_excel_adjustments(path, 'All URLs', sett_urls)
    print('Excel file saved.')

