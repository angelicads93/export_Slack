#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: Angelica Goncalves.

Module to build an Excel workbook with the compiled weekly reports.

Functions
---------
apply_excel_adjustments_stats(file_path, sheet_name, settings)
    Format the Excel tables as specified in the settings txt file.

checkins_to_excel(setts)
    Compile message's check-in reports into Excel workbook.

"""

# Import standard libraries:
import sys
import os
import pandas as pd

# Include the main repo directory (export_Slack/) and the src directory
# (export_Slack/src) to the Python path so the customed modules can be imported
parent_dir = os.path.dirname(os.getcwd())
sys.path.append(parent_dir)
sys.path.append(os.path.join(parent_dir, "src"))
import excel
import clean
import sparser
import slack


def apply_excel_adjustments_stats(file_path, sheet_name, settings_file):
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


def checkins_to_excel(setts):
    """
    Compile message's check-in reports into Excel workbook.

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

                # Add channels and reports info:
                ch_df = s.add_channel_info(f"{setts.get('excel_channels_path')}/{file}", ch_df)
                ch_df = s.add_info_of_users_reports(ch_df)

                # Handle missing values:
                ch_df = clean.handle_missing_values(ch_df, setts.get("missing_value"))

                # Reorder columns:
                ch_df = ch_df[setts.get("columns_order")]

                # Concatanate ch_df to final dataframe:
                if file == os.listdir(setts.get("excel_channels_path"))[0]:
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

    # Select rows with unparsed projects in the official weekly report channel:
    df_unparsed = s.format_unparsed_reports(df, setts.get("reports_channel_name"))
    unparsed_ws_name = "Unparsed weekly reports"
    print("Retrieve un-parsed weekly reports.")

    # Select rows with parsed projects:
    df_parsed = s.format_parsed_reports(df)
    parsed_ws_name = "Parsed weekly reports"
    print("Retrieve parsed weekly reports from all the channels.")

    # Save Excel workbook:
    path = f"{setts.get('compilation_reports_path')}/{setts.get('compilation_reports_file_name')}"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df_parsed.to_excel(writer, sheet_name=parsed_ws_name, index=False)
        df_unparsed.to_excel(writer, sheet_name=unparsed_ws_name, index=False)
        df.to_excel(writer, sheet_name="All messages", index=False)

    # Apply formatting of Excel worksheets:
    apply_excel_adjustments_stats(path, parsed_ws_name, setts)
    apply_excel_adjustments_stats(path, unparsed_ws_name, setts)
    apply_excel_adjustments_stats(path, "All messages", setts)
    print("Excel file saved.")


# #############################################################################
if __name__ == "__main__":

    # -------------------------------------------------------------------------
    print("\n", "Parsing the user's input...")
    
    # Parse the user's command in the terminal:
    args = sparser.init_command_parser(
        "Python script to compile all the weekly reports from individual Excel files.",
        ["settings_file_path"])

    # --Verify that settings_file_path exists:
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
                                    "compilation_reports_path",
                                    settings.get("compilation_reports_path"),
                                    kill=False)

    # -------------------------------------------------------------------------
    print("\n", "Building dataframes and writing Excel files...")
    checkins_to_excel(settings)
