#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: Angelica Goncalves.

Python code to export Slack messages from JSON files into Excel databases.

"""
import os
import sys
import pandas as pd

# Include the main repo directory (export_Slack/) and the src directory
# (export_Slack/src) to the Python path so the customed modules can be imported
parent_dir = os.path.dirname(os.getcwd())
sys.path.append(parent_dir)
sys.path.append(os.path.join(parent_dir, "src"))
import sparser
import slack
import checkins
import excel


def apply_excel_adjustments_msgs(file_path, ws_name, settings_msgs):
    """
    Format the Excel tables as specified in the settings txt file.

    Arguments
    ---------
    file_path : str
        Absolute path to the Excel file.
    ws_name : str
        Name of the Excel Sheet.
    settings_msgs : parser.Parser(txt_path)
        Parsed variables from the settings txt file.

    """
    xl = excel.ExcelFormat(file_path, settings_msgs)
    ws = xl.get_sheet(ws_name)
    xl.set_cell_width(ws, settings_msgs.get("column_widths"))
    xl.set_allignment(ws, "top")
    xl.format_first_row(ws, settings_msgs.get("header_row"))
    for cc in settings_msgs.get("font_color_in_column"):
        xl.set_font_color_in_column(ws, cc)
    for highlight in settings_msgs.get("highlights"):
        xl.format_highlight(ws, highlight)
    for column in settings_msgs.get("text_type_cols"):
        xl.format_text_cells(ws, column)
    xl.save_changes()


def msgs_to_excel(chs, all_users_df, inps, setts):
    """
    Compile messages in a Slack workspace into Excel workbooks.

    Arguments
    ---------
    chs : list
        List with the names of the Slack channels to analyze.
    all_users_df : Pandas dataframe
        Dataframe with all the Slack user's information.
    inps : parser.Parser(txt_path)
        Parsed variables from the inputs.txt file.
    setts : parser.Parser(txt_path)
        Parsed variables from the settings_messages.txt file.

    """
    s = slack.Slack(setts)

    print("")
    # Iterate over channel's folders:
    for ch in chs:

        # Collect all the current_channel's messages in ch_msgs_df:
        ch_msgs_df = s.get_ch_msgs_df(inps.get('slackexport_folder_path'), ch)
        print(f"{ch} Collected channel msgs from the json files")
        if len(ch_msgs_df) < 1:
            continue

        # Collect all the users in the current channel:
        ch_usrs_df = s.get_ch_usrs_df(ch_msgs_df, all_users_df)
        print(f"{ch} Collected users in current channel")

        # Use ch_usrs_df to fill in the user's information in ch_msgs_df:
        s.add_usrs_info_to_msgs_df(ch_msgs_df, ch_usrs_df)
        print(f"{ch} Included the users info on ch_msgs_df")

        # Replace user and team identifiers with their display_names whenever
        # present in a message:
        s.usr_id_to_name(ch_msgs_df, ch_usrs_df)
        s.ch_id_to_name(ch_msgs_df)
        s.parent_id_to_name(ch_msgs_df, ch_usrs_df)
        print(f"{ch} User's id replaced by their names")

        # Extract hyperlinks from messages, if present (extracted as a list;
        # edit if needed):
        s.extract_urls(ch_msgs_df)
        print(f"{ch} URLs extracted from messages")

        # Change format of the time in seconds to a date in the CST time-zone:
        s.ts_to_tz(ch_msgs_df, "ts", "msg_date")
        s.ts_to_tz(ch_msgs_df, "json_mod_ts", "json_mod_date")
        s.ts_to_tz(ch_msgs_df, "ts_latest_reply", "latest_reply_date")
        s.ts_to_tz(ch_msgs_df, "ts_thread", "thread_date")
        print(f"{ch} Formated the dates and times")

        # Identify if text has emojis:
        ch_msgs_df = s.id_emojis_in_text(ch_msgs_df)
        print(f"{ch} Checked for emojis in messages")

        # Parse for check-in messages:
        ch_msgs_df = checkins.parse_reports(ch_msgs_df, setts)

        ch_msgs_df = s.drop_extra_unparsed_rows(ch_msgs_df)
        print(f"{ch} Parsed check-in messages")

        # Build df with pruned messages:
        sel_msgs_df = s.rm_automatic_msgs(ch_msgs_df)
        sel_msgs_df = s.remove_emojis_in_text(sel_msgs_df)
        sel_msgs_df = s.rm_short_msgs(sel_msgs_df, n_char=15)
        print(f"{ch} Built df with selected rows")

        # Build df with discarded messages (after being pruned):
        auto_msgs = s.get_automatic_msgs(ch_msgs_df)
        short_msgs = s.get_short_msgs(ch_msgs_df, n_char=15)
        dis_msgs = pd.concat([auto_msgs, short_msgs],
                             axis=0, ignore_index=False)
        dis_msgs.sort_values(by="msg_date", inplace=True, ignore_index=True)
        print(f"{ch} Built df with filtered-out messages")

        # Rearrange columns:
        column_names_order = setts.get("columns_order")
        ch_msgs_df = ch_msgs_df[column_names_order]
        sel_msgs_df = sel_msgs_df[column_names_order]
        dis_msgs = dis_msgs[column_names_order]
        print(f"{ch} Rearranged columns")

        # Sort rows by msg_date:
        ch_msgs_df.sort_values(by="msg_date", inplace=True, ignore_index=True)
        sel_msgs_df.sort_values(by="msg_date", inplace=True, ignore_index=True)
        dis_msgs.sort_values(by="msg_date", inplace=True, ignore_index=True)
        print(f"{ch} Sorted rows by msg_date")

        # Write ch_msgs_df to a .xlsx file:
        msgs_mindate = ch_msgs_df["msg_date"].min().split(" ")[0]
        msgs_maxdate = ch_msgs_df["msg_date"].max().split(" ")[0]
        ch = ch.replace(" ", "-")
        ch_msgs_filename = f"{ch}_{msgs_mindate}_to_{msgs_maxdate}"
        path = f"{inps.get('converted_directory')}"
        path += "/" + f"{setts.get('dest_name_ext')}/{ch_msgs_filename}.xlsx"
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            sel_msgs_df.to_excel(writer, index=False,
                                 sheet_name="Relevant messages")
            dis_msgs.to_excel(writer, index=False,
                              sheet_name="Filtered-out messages")
            ch_msgs_df.to_excel(writer, index=False,
                                sheet_name="All messages")

        # Apply formatting of Excel worksheets:
        apply_excel_adjustments_msgs(path, "All messages", setts)
        apply_excel_adjustments_msgs(path, "Relevant messages", setts)
        apply_excel_adjustments_msgs(path, "Filtered-out messages", setts)
        print(f"{ch} Wrote curated messages to Excel \n")

    print("Done")


# #############################################################################
if __name__ == "__main__":

    # -------------------------------------------------------------------------
    print("\n", "Parsing the user's inputs...")

    # Parse the user's command in the terminal:
    args = sparser.init_command_parser(
        "Python script to convert a Slack workspace into Excel databases",
        ["inputs_file_path", "settings_file_path"])

    # Verify that input_file_path exists and parse its content:
    inputs_file_path = os.path.abspath(args.inputs_file_path)
    sparser.check_path_in_user_file("your command prompt.",
                                    "--inputs_file_path", inputs_file_path,
                                    kill=True)
    inputs = sparser.Parser(inputs_file_path)

    # Verify that settings_file_path exists and parse its content:
    settings_file_path = os.path.abspath(args.settings_file_path)
    sparser.check_path_in_user_file("your command prompt.",
                                    "--settings_file_path", settings_file_path,
                                    kill=True)
    settings = sparser.Parser(settings_file_path)

    # Check that the source directory exists:
    sparser.check_path_in_user_file(inputs.file_name + ".txt",
                                    "slackexport_folder_path",
                                    inputs.get('slackexport_folder_path'),
                                    kill=True)

    # List the Slack channels to be analyzed:
    flag_all_chs = sparser.set_flag_analyze_all_chs(inputs.get('chosen_channel_name'))
    if flag_all_chs is True:
        chs2analyze = sparser.list_dirs_in_path(inputs.get('slackexport_folder_path'))
    else:
        ch_path = f"{inputs.get('slackexport_folder_path')}"
        ch_path += "/" + f"{inputs.get('chosen_channel_name')}"
        if sparser.check_file_exists(ch_path, kill=True) is True:
            chs2analyze = [inputs.get('chosen_channel_name')]

    # Check that the "channels" and "users" JSON files exist:
    sparser.check_file_exists(
        f"{inputs.get('slackexport_folder_path')}/{settings.get('channels_json_name')}",
        kill=True)
    sparser.check_file_exists(
        f"{inputs.get('slackexport_folder_path')}/{settings.get('users_json_name')}",
        kill=True)

    # -------------------------------------------------------------------------
    print("\n", "Building dataframes and writing Excel files...")

    # Create the path where the files will be saved:
    sparser.make_dest_path(
        f"{inputs.get('converted_directory')}/{settings.get('dest_name_ext')}")

    # Get dataframes with the channels info:
    all_channels_df = slack.Slack(settings).get_all_channels_info(
        inputs.get('slackexport_folder_path'),
        f"{inputs.get('slackexport_folder_path')}/{settings.get('channels_json_name')}")

    # Get dataframes with the users info:
    all_usrs_df = slack.Slack(settings).get_all_users_info(
        f"{inputs.get('slackexport_folder_path')}/{settings.get('users_json_name')}")

    # Write Excel files from channels.json and users.json if requested by user:
    slack.write_info_to_file(
        inputs.get("write_all_channels_info"),
        all_channels_df,
        settings.get("channels_excel_name").split(".")[0],
        f"{inputs.get('converted_directory')}/{settings.get('dest_name_ext')}")
    slack.write_info_to_file(
        inputs.get("write_all_users_info"),
        all_usrs_df,
        settings.get("users_excel_name").split(".")[0],
        f"{inputs.get('converted_directory')}/{settings.get('dest_name_ext')}")

    # Write the Excel files of the given channel(s):
    msgs_to_excel(chs2analyze, all_usrs_df, inputs, settings)
