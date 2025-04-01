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
              + " 'compilation_reports_file_name' in {file_name}.")
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


def add_info_of_users_reports(df):
    """ Get latest report date and number of messages in the given channel for
    all the users in the channel.
    """
    users = df['user'].unique()
    for user in users:
        user_df = df[df['user'] == user].sort_values(
            by='msg_date', inplace=False, ignore_index=True
            )
        latest_report_date = user_df['msg_date'].to_list()[-1]
        user_df['latest_report_date'] = [latest_report_date]*len(user_df)
        user_df['number_msgs_in_channel'] = len(user_df)
        if user == users[0]:
            df_out = user_df.copy()
        else:
            df_out = pd.concat([df_out, user_df], axis=0, ignore_index=False)

    return df_out


def format_parsed_reports(df):
    """ Select the parsed weekly reports and sort the df by channel. """
    df_p = df.copy()
    df_p = df_p[df_p['projects_parsed'] != '0']
    df_p = df_p.reset_index().drop(columns=['index'])
    df_p.sort_values(by=['channel', 'display_name', 'msg_date'],
                     inplace=True, ignore_index=True)
    return df_p


def format_unparsed_reports(df, wr_channel_name):
    """ Select the unparsed weekly reports and sort the df by channel. """
    df_np = df.copy()
    df_np = df_np[df_np['channel'] == wr_channel_name]
    df_np = df_np[df_np['projects_parsed'] == "0"]
    df_np = df_np[df_np['msg_id'] != 'channel_join']
    df_np = df_np[df_np['is_bot'] != True]
    df_np = df_np[df_np['type'] != 'thread']
    df_np = df_np.reset_index().drop(columns=['index'])
    df_np.sort_values(by=['channel', 'display_name', 'msg_date'],
                      inplace=True, ignore_index=True)
    return df_np


def apply_excel_adjustments(file_path, sheet_name, settings):
    """ Defines the sequence of changes to be done in the Excel file
    given the user's inputs in the module "settings".
    """
    xl = excel.ExcelFormat(file_path, settings)
    ws_channel = xl.get_sheet(sheet_name)

    xl.set_cell_width(ws_channel, settings.get('column_widths'))
    xl.draw_vertical_line(ws_channel, settings.get('draw_vert_line'))
    xl.set_allignment(ws_channel, 'top')
    xl.format_first_row(ws_channel, settings.get('header_row'))
    for cc in settings.get('font_color_in_column'):
        xl.set_font_color_in_column(ws_channel, cc)
    for highlight in settings.get('highlights'):
        xl.format_highlight(ws_channel, highlight)
    for column in settings.get('text_type_cols'):
        xl.format_text_cells(ws_channel, column)

    xl.set_filters(ws_channel)
    xl.save_changes()


# -----------------------------------------------------------------------------
if __name__ == '__main__':
    # --Parse the settings file:
    print("Checking arguments of Python command")
    settings_file_path = parse_command_input()

    print("Parsing information in settings file")
    file_name = os.path.basename(settings_file_path)
    sett_stats = parser.Parser(os.path.abspath(settings_file_path))

    # --Check the input and retrieve expected name of channels:
    print("Checking validity of input paths")
    check_input(file_name, 
                sett_stats.get('compilation_reports_path'),
                sett_stats.get('excel_channels_path'))
    expected_channels = get_list_channels(sett_stats.get('jsons_source_path'))

    # --Build dataframe from all the channels:
    for file in os.listdir(sett_stats.get('excel_channels_path')):
        file_name = str(file).split(".")[0]
        channel_path = f"{sett_stats.get('excel_channels_path')}/{file}"

        if check_channel(file_name, expected_channels) is True:
            channel_df = pd.read_excel(channel_path, engine='openpyxl',
                                       sheet_name='Relevant messages')

            if len(channel_df) > 0:
                # --Add channels and reports info:
                channel_df = add_channel_info(channel_path, channel_df)
                channel_df = add_info_of_users_reports(channel_df)

                # --Handle missing values:
                channel_df = clean.handle_missing_values(channel_df,
                                                         sett_stats.get('missing_value'))

                # --Reorder columns:
                channel_df = channel_df[sett_stats.get('columns_order')]

                # --Concatanate channel_df to final dataframe:
                if file == os.listdir(sett_stats.get('excel_channels_path'))[0]:
                    df = channel_df.copy()
                else:
                    df = pd.concat([df, channel_df], axis=0, ignore_index=False)
    clean.reset_indices(df)
    print('Information of all the check-in reports collected.')

    # --Set columns types:
    df['projects_parsed'] = df['projects_parsed'].astype('string')
    df['keywords_parsed'] = df['keywords_parsed'].astype('string')
    print('Set data type of columns.')

    # --Select rows with un-parsed projects in official weekly report channel:
    df_unparsed = format_unparsed_reports(df, sett_stats.get('reports_channel_name'))
    unparsed_ws_name = "Unparsed weekly reports"
    print('Retrieve un-parsed weekly reports.')

    # --Select rows with parsed projects:
    df_parsed = format_parsed_reports(df)
    parsed_ws_name = "Parsed weekly reports"
    print('Retrieve parsed weekly reports from all the channels.')

    # --Save Excel workbook:
    path = f"{sett_stats.get('compilation_reports_path')}/{sett_stats.get('compilation_reports_file_name')}"
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df_parsed.to_excel(writer, sheet_name=parsed_ws_name, index=False)
        df_unparsed.to_excel(writer, sheet_name=unparsed_ws_name, index=False)
        df.to_excel(writer, sheet_name='All messages', index=False)

    # --Apply formatting of Excel worksheets:
    apply_excel_adjustments(path, parsed_ws_name, sett_stats)
    apply_excel_adjustments(path, unparsed_ws_name, sett_stats)
    apply_excel_adjustments(path, 'All messages', sett_stats)
    print('Excel file saved.')

