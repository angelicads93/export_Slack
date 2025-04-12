#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: Angelica Goncalves.

Module to build an Excel workbook with the compiled weekly reports.

Functions
---------
check_ch(file_name, source_path)
    Check if file_name is indeed an expected Slack channel.

add_channel_info(channel_path, channel_df)
    Get the name of a channel, export time, and relative number of reports.

add_info_of_users_reports(df_all)
    Get the latest report date and number of messages in the given channel.

format_parsed_reports(df_all)
    Select the parsed weekly reports and sort the df by channel.

format_unparsed_reports(df_all, wr_channel_name)
    Select the unparsed weekly reports and sort the df by channel.

apply_excel_adjustments(file_path, sheet_name, settings)
    Format the Excel tables as specified in the settings txt file.
"""

# Import standard libraries:
import sys
import os
import argparse
import pandas as pd


# Import custom modules:
parent_dir = os.path.dirname(os.getcwd())
sys.path.append(parent_dir)
sys.path.append(os.path.join(parent_dir, "src"))
import excel
import clean
import sparser


def check_ch(file_name, source_path):
    """
    Check if file_name is indeed an expected Slack channel.

    Notice that the default name of a channel's directory may contain empty
    spaces (" "), although the name of the corresponding Excel file does not.
    Empty spaces were replaced by underscores in the later.

    Arguments
    ---------
    file_name : str
        Name of the Slack channel.

    source_path : str
        Path to the source directory containing all the Slack JSON files.

    Returns
    -------
    Boolean. Whether file_name is present in source_path.

    """
    # Collect the name of all channel's subdirectories after removing all type
    # of spaces and separators from their names:
    chs_path = sparser.list_dirs_in_path(source_path)
    for i, ch in enumerate(chs_path):
        ch = ch.replace(" ", "").replace("-", "").replace("_", "")
        chs_path[i] = ch

    # Extract channel_name from the full name of the Excel file
    # ("<channel_name>_<start_date>_to_<end_date>.xlsx"):
    cn = "_".join(file_name.split("_")[:-3])

    # Remove all types of spaces and separators to compare the names further:
    cn = cn.replace(" ", "").replace("-", "").replace("_", "")

    # Compare the names:
    out = True
    if cn not in chs_path:
        out = False

    return out


def add_channel_info(channel_path, channel_df):
    """
    Get the name of a channel, export time, and relative number of reports.

    Arguments
    ---------
    channel_path : str
        Absolute path to the channels directory.

    channel_df : Pandas dataframe
        Dataframe with the information of all the channels in the Slack
        workspace.

    Returns
    -------
    Pandas dataframe

    """
    file_name = channel_path.split("/")[-1].split(".")[0]
    channel_name = "_".join(file_name.split("_")[:-3])
    channel_date = "_".join(file_name.split("_")[-3:])

    df_ = channel_df["projects_parsed"].astype("string")
    reports_in_channel = f"{len(df_[df_ != '0'])}/{len(channel_df)}"
    channel_df["channel"] = [channel_name]*len(channel_df)
    channel_df["export_dates"] = [channel_date]*len(channel_df)
    channel_df["parsed_reports_in_channel"] = [reports_in_channel]*len(channel_df)

    return channel_df


def add_info_of_users_reports(df_all):
    """
    Get the latest report date and number of messages in the given channel.

    Arguments
    ---------
    df_all : Pandas dataframe

    Returns
    -------
    Pandas dataframe with extra columns "latest_report_date" and
    "number_msgs_in_channel".

    """
    users = df_all["user"].unique()
    for user in users:
        user_df = df_all[df_all["user"] == user].sort_values(
            by="msg_date", inplace=False, ignore_index=True
            )
        latest_report_date = user_df["msg_date"].to_list()[-1]
        user_df["latest_report_date"] = [latest_report_date]*len(user_df)
        user_df["number_msgs_in_channel"] = len(user_df)
        if user == users[0]:
            df_out = user_df.copy()
        else:
            df_out = pd.concat([df_out, user_df], axis=0, ignore_index=False)

    return df_out


def format_parsed_reports(df_all):
    """
    Select the parsed weekly reports and sort the df by channel.

    Arguments
    ---------
    df_all : Pandas dataframe
        Pandas dataframe with all the messages.

    Returns
    -------
    Pandas dataframe containing messages parsed as "weekly reports".


    """
    df_p = df_all.copy()
    df_p = df_p[df_p["projects_parsed"] != "0"]
    df_p = df_p.reset_index().drop(columns=["index"])
    df_p.sort_values(by=["channel", "display_name", "msg_date"],
                     inplace=True, ignore_index=True)
    return df_p


def format_unparsed_reports(df_all, wr_channel_name):
    """
    Select the unparsed weekly reports and sort the df by channel.

    Arguments
    ---------
    df_all : Pandas dataframe
        Pandas dataframe with all the messages.
    wr_ch_name : str
        Channel name of the given weekly report.

    Returns
    -------
    Pandas dataframe containing messages NOT parsed as "weekly report".

    """
    df_np = df_all.copy()
    df_np = df_np[df_np["channel"] == wr_channel_name]
    df_np = df_np[df_np["projects_parsed"] == "0"]
    df_np = df_np[df_np["msg_id"] != "channel_join"]
    df_np = df_np[df_np["is_bot"] != True]
    df_np = df_np[df_np["type"] != "thread"]
    df_np = df_np.reset_index().drop(columns=["index"])
    df_np.sort_values(by=["channel", "display_name", "msg_date"],
                      inplace=True, ignore_index=True)
    return df_np


def apply_excel_adjustments(file_path, sheet_name, settings_file):
    """
    Format the Excel tables as specified in the settings txt file.

    Arguments
    ---------
    file_path : str
        Absolute path to the Excel file

    sheet_name : str
        Name of the Excel Sheet

    settings_file : parser.Parser(txt_path)
        Parsed user inputs to the settings txt file.

    """
    xl = excel.ExcelFormat(file_path, settings_file)
    ws_channel = xl.get_sheet(sheet_name)

    xl.set_cell_width(ws_channel, settings_file.get("column_widths"))
    xl.draw_vertical_line(ws_channel, settings_file.get("draw_vert_line"))
    xl.set_allignment(ws_channel, "top")
    xl.format_first_row(ws_channel, settings_file.get("header_row"))
    for cc in settings_file.get("font_color_in_column"):
        xl.set_font_color_in_column(ws_channel, cc)
    for highlight in settings_file.get("highlights"):
        xl.format_highlight(ws_channel, highlight)
    for column in settings_file.get("text_type_cols"):
        xl.format_text_cells(ws_channel, column)

    xl.set_filters(ws_channel)
    xl.save_changes()


# -----------------------------------------------------------------------------
if __name__ == "__main__":

    # #########################################################################
    # INSPECT USERS INPUT AND SOURCE DIRECTORY:
    print("------------------------------------------------------------------")
    print("Parsing the user's command...")

    # --Define argument parser routine:
    arg_parser = argparse.ArgumentParser(
        description="Python script to compile all the weekly reports from "
        + "individual Excel files."
        )
    arg_parser.add_argument("--settings_file_path", required=True, type=str)
    args = arg_parser.parse_args()

    # --Verify that settings_file_path exists:
    settings_file_path = os.path.abspath(args.settings_file_path)
    sparser.check_path_in_user_file("your command prompt.",
                                    "--settings_file_path", settings_file_path,
                                    kill=True)

    # Retrieve variables in settings file:
    settings = sparser.Parser(settings_file_path)
    jsons_source_path = settings.get("jsons_source_path")
    excel_channels_path = settings.get("excel_channels_path")
    reports_channel_name = settings.get("reports_channel_name")
    compilation_reports_file_name = settings.get("compilation_reports_file_name")
    compilation_reports_path = settings.get("compilation_reports_path")
    missing_value = settings.get("missing_value")

    # Check that the path with all the converted Excel files exists:
    sparser.check_path_in_user_file(settings.file_name + ".txt",
                                    "compilation_reports_path",
                                    compilation_reports_path,
                                    kill=False)

    # Check that the destination path where files will be saved exists:
    sparser.check_path_in_user_file(settings.file_name + ".txt",
                                    "excel_channels_path",
                                    excel_channels_path,
                                    kill=False)

    # #########################################################################
    # BUILD DATAFRAMES AND WRITE EXCEL FILES:
    print('------------------------------------------------------------------')
    print('Building dataframes and writing Excel files...')

    for file in os.listdir(excel_channels_path):

        # Check that the Excel file corresponds to a Slack channel:
        if check_ch(str(file).split(".")[0], jsons_source_path) is True:

            # If so, load the Excel sheet into a dataframe:
            ch_df = pd.read_excel(f"{excel_channels_path}/{file}",
                                  engine="openpyxl",
                                  sheet_name="Relevant messages")

            # If the dataframe is not empty:
            if len(ch_df) > 0:

                # Add channels and reports info:
                ch_df = add_channel_info(f"{excel_channels_path}/{file}", ch_df)
                ch_df = add_info_of_users_reports(ch_df)

                # Handle missing values:
                ch_df = clean.handle_missing_values(ch_df, missing_value)

                # Reorder columns:
                ch_df = ch_df[settings.get("columns_order")]

                # Concatanate ch_df to final dataframe:
                if file == os.listdir(excel_channels_path)[0]:
                    df = ch_df.copy()
                else:
                    df = pd.concat([df, ch_df], axis=0, ignore_index=False)

    # Reset the indices of the dataframe:
    clean.reset_indices(df)
    print("Information of all the check-in reports collected.")

    # Set column types:
    df["projects_parsed"] = df["projects_parsed"].astype("string")
    df["keywords_parsed"] = df["keywords_parsed"].astype("string")
    print("Set data type of columns.")

    # Select rows with un-parsed projects in the official weekly report channel:
    df_unparsed = format_unparsed_reports(df, reports_channel_name)
    unparsed_ws_name = "Unparsed weekly reports"
    print("Retrieve un-parsed weekly reports.")

    # Select rows with parsed projects:
    df_parsed = format_parsed_reports(df)
    parsed_ws_name = "Parsed weekly reports"
    print("Retrieve parsed weekly reports from all the channels.")

    # Save Excel workbook:
    path = f"{compilation_reports_path}/{compilation_reports_file_name}"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df_parsed.to_excel(writer, sheet_name=parsed_ws_name, index=False)
        df_unparsed.to_excel(writer, sheet_name=unparsed_ws_name, index=False)
        df.to_excel(writer, sheet_name="All messages", index=False)

    # Apply formatting of Excel worksheets:
    apply_excel_adjustments(path, parsed_ws_name, settings)
    apply_excel_adjustments(path, unparsed_ws_name, settings)
    apply_excel_adjustments(path, "All messages", settings)
    print("Excel file saved.")
