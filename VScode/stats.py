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


def check_input(compilation_reports_path, excel_channels_path):
    if os.path.exists(excel_channels_path) is False:
        print(f"ERROR: Path {excel_channels_path} does not exists. Please review your input for the variable 'excel_channels_path' in {module_name}.")
        sys.exit()
    if os.path.exists(compilation_reports_path) is False:
        print(f"ERROR: Path {compilation_reports_path} does not exists. Please review your input for the variable '' in {module_name}.")
        sys.exit()


def get_list_channels(source_path):
    channels_path = os.listdir(source_path)
    for i, ch in enumerate(channels_path):
        ch = ch.replace(" ", "").replace("-", "").replace("_", "")
        channels_path[i] = ch
    return channels_path


def check_channel(channel_name, list_channels):
    cn = "_".join(channel_name.split("_")[:-3])
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


def apply_excel_adjustments(file_path, sheet_name, settings_mod):
    """ Defines the sequence of changes to be done in the Excel file
    given the user's inputs in the module settings_mod.
    """
    xl = excel.ExcelFormat(file_path)
    ws_channel = xl.get_sheet(sheet_name)

    xl.set_cell_width(ws_channel, settings_mod.column_widths)
    xl.draw_vertical_line(ws_channel, settings_mod.draw_vert_line)
    xl.set_allignment(ws_channel, 'top')
    xl.format_first_row(ws_channel,
                        height=settings_mod.height_1strow,
                        aling_vert=settings_mod.alignment_vert_1strow,
                        aling_horiz=settings_mod.alignment_horiz_1strow,
                        font_size=settings_mod.font_size_1strow,
                        font_bold=settings_mod.font_bold_1strow,
                        cell_color_1strow=settings_mod.cell_color_1strow
                        )
    for cc in settings_mod.font_color_in_column:
        xl.set_font_color_in_column(ws_channel, cc)
    for highlight in settings_mod.highlights:
        xl.format_highlight(ws_channel, highlight)
    for column in settings_mod.text_type_cols:
        xl.format_text_cells(ws_channel, column)

    xl.set_filters(ws_channel)

    xl.save_changes()


if __name__ == '__main__':

    # --Define argument parser routine:
    parser = argparse.ArgumentParser(
        description="Python script to compile all the weekly reports from individual Excel files."
        )
    parser.add_argument("--settings_file_path", required=True, type=str)
    args = parser.parse_args()
    settings_file_path = args.settings_file_path
    print(f"settings_file_path = {settings_file_path}")
    if os.path.exists(settings_file_path) is False:
        print(f"ERROR: Path {settings_file_path} does not exists. Please review your input for the argument --settings_file_path.")
        sys.exit()

    # --Import settings module:
    parent_path = os.path.dirname(settings_file_path)
    module_name = os.path.basename(settings_file_path).split(".")[0]
    sys.path.append(parent_path)
    settings_module = importlib.import_module(module_name)

    missing_value = settings_module.missing_value
    excel_channels_path = settings_module.excel_channels_path
    compilation_reports_file_name = settings_module.compilation_reports_file_name
    compilation_reports_path = settings_module.compilation_reports_path
    print(f"Module {module_name} imported.")

    # --Check the input and retrieve expected name of channels:
    check_input(compilation_reports_path, excel_channels_path)
    expected_channels = get_list_channels(settings_module.jsons_source_path)

    # --Build dataframe from all the channels:
    for file in os.listdir(excel_channels_path):
        file_name = str(file).split(".")[0]
        channel_path = f"{excel_channels_path}/{file}"

        if check_channel(file_name, expected_channels) is True:
            channel_df = pd.read_excel(channel_path, engine='openpyxl')

            # --Add channels and reports info:
            channel_df = add_channel_info(channel_path, channel_df)
            channel_df = add_info_of_users_reports(channel_df)

            # --Handle missing values:
            channel_df = clean.handle_missing_values(channel_df, missing_value)

            # --Reorder columns:
            channel_df = channel_df[settings_module.columns_order]

            # --Concatanate channel_df to final dataframe:
            if file == os.listdir(excel_channels_path)[0]:
                df = channel_df.copy()
            else:
                df = pd.concat([df, channel_df], axis=0, ignore_index=False)
    print('Information of all the check-in reports collected.')

    # --Set columns types:
    df['projects_parsed'] = df['projects_parsed'].astype('string')
    df['keywords_parsed'] = df['keywords_parsed'].astype('string')

    # --Select rows with parsed projects:
    df = df[df['projects_parsed'] != '0']
    df = df.reset_index().drop(columns=['index'])
    print('Performed minor formatting of columns and rows.')

    # --Sort columns:
    df_by_channels = df.copy()
    df_by_channels.sort_values(
        by=['channel', 'display_name', 'msg_date'],
        inplace=True, ignore_index=True
        )

    # df_by_users = df.copy()
    # df_by_users.sort_values(
    #    by=['display_name', 'channel', 'msg_date'],
    #    inplace=True, ignore_index=True
    #    )

    # --Save Excel file:
    #df.to_excel(path, index=False)
    path = f"{compilation_reports_path}/{compilation_reports_file_name}"
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df_by_channels.to_excel(writer, sheet_name='Weekly reports', index=False)
        #df_by_users.to_excel(writer, sheet_name='Sorted by User', index=False)

    apply_excel_adjustments(path, 'Weekly reports', settings_module)
    #apply_excel_adjustments(path, 'Sorted by User', settings_module)

    print('Excel file saved.')
