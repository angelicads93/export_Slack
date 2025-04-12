#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: Angelica Goncalves.

Module to build an Excel workbook with the URL(s) found in the Slack messages.

Functions
---------
apply_excel_adjustments_urls(file_path, sheet_name, settings)
    Format the Excel tables as specified in the settings txt file.

urls_to_excel(setts)
    Compile message's URLs into Excel workbook.

"""

# Import standard libraries:
import sys
import os
import pandas as pd

# Include the main repo directory (export_Slack/) and the src directory
# (export_Slack/src) to the Python path so the customed modules can be imported
parent_dir = os.path.dirname(os.getcwd())
sys.path.append(parent_dir)
sys.path.append(os.path.join(parent_dir, 'src'))
import excel
import clean
import sparser
import slack


def apply_excel_adjustments_urls(file_path, sheet_name, settings_url):
    """
    Format the Excel tables as specified in the settings txt file.

    Arguments
    ---------
    file_path : str
        Absolute path to the Excel file.
    ws_name : str
        Name of the Excel Sheet.
    settings_url : parser.Parser(txt_path)
        Parsed variables from the settings_url txt file.

    """
    xl = excel.ExcelFormat(file_path, settings_url)
    ws_channel = xl.get_sheet(sheet_name)

    xl.set_cell_width(ws_channel, settings_url.get('column_widths'))
    xl.set_allignment(ws_channel, 'top')
    xl.format_first_row(ws_channel, settings_url.get('header_row'))

    xl.set_filters(ws_channel)
    xl.save_changes()


def urls_to_excel(setts):
    """
    Compile message's URLs into Excel workbook.

    Arguments
    ---------
    setts : parser.Parser(txt_path)
        Parsed variables from the settings_url txt file.

    """
    s = slack.Slack(setts)

    for file in os.listdir(setts.get("excel_channels_path")):

        # Check that the Excel file corresponds to a Slack channel:
        if sparser.check_ch(str(file).split(".")[0], setts.get("jsons_source_path")) is True:

            # If so, load the Excel sheet into a dataframe:
            ch_df = pd.read_excel(f"{setts.get('excel_channels_path')}/{file}",
                                  engine="openpyxl",
                                  sheet_name="Relevant messages")

            # If the dataframe is not empty:
            if len(ch_df) > 0:

                # Add channels info:
                ch_df = s.add_channel_info(f"{setts.get('excel_channels_path')}/{file}", ch_df)

                # Handle missing values:
                ch_df = clean.handle_missing_values(ch_df, setts.get("missing_value"))

                # Reorder columns:
                ch_df = ch_df[setts.get('columns_order')]

                # Concatanate ch_df to final dataframe:
                if file == os.listdir(setts.get("excel_channels_path"))[0]:
                    df = ch_df.copy()
                else:
                    df = pd.concat([df, ch_df], axis=0, ignore_index=False)
    clean.reset_indices(df)
    print('Information of all the check-in reports collected.')

    # Select rows with URL(s) in their messages:
    df = s.format_msgs_with_urls(df, setts)

    # Filter rows with desire URLs:
    df_selection = s.select_desired_urls(df, setts)

    # Save Excel workbook:
    path = f"{setts.get('compilation_urls_path')}/{setts.get('compilation_urls_file_name')}"
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df_selection.to_excel(writer, sheet_name='Selected URLs', index=False)
        df.to_excel(writer, sheet_name='All URLs', index=False)

    # Apply formatting of Excel worksheets:
    apply_excel_adjustments_urls(path, 'Selected URLs', setts)
    apply_excel_adjustments_urls(path, 'All URLs', setts)
    print('Excel file saved.')



# #############################################################################
if __name__ == '__main__':

    # -------------------------------------------------------------------------
    print("\n", "Parsing the user's input...")
    
    # Parse the user's command in the terminal:
    args = sparser.init_command_parser(
        "Python script to extract URLs from Slack messages.",
        ["settings_file_path"])

    # Verify that settings_file_path exists:
    settings_file_path = os.path.abspath(args.settings_file_path)
    sparser.check_path_in_user_file("your command prompt.",
                                    "--settings_file_path", settings_file_path,
                                    kill=True)

    # Retrieve variables in settings file:
    settings = sparser.Parser(settings_file_path)

    # Check that the source path with all Slack JSON files exists:
    sparser.check_path_in_user_file(settings.file_name + ".txt",
                                    "jsons_source_path",
                                    settings.get("jsons_source_path"),
                                    kill=False)

    # Check that the path with all the converted Excel files exists:
    sparser.check_path_in_user_file(settings.file_name + ".txt",
                                    "excel_channels_path",
                                    settings.get("excel_channels_path"),
                                    kill=False)

    # Check that the destination path where files will be saved exists:
    sparser.check_path_in_user_file(settings.file_name + ".txt",
                                    "compilation_urls_path",
                                    settings.get("compilation_urls_path"),
                                    kill=False)

    # -------------------------------------------------------------------------
    print("\n", "Building dataframes and writing Excel files...")
    urls_to_excel(settings)
