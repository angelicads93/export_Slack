#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: Angelica Goncalves.

Python code to export Slack messages from JSON files into Excel databases.

"""
import os
import sys
import argparse

# Include the main repo directory (export_Slack/) and the src directory
# (export_Slack/src) to the Python path so the customed modules can be imported
parent_dir = os.path.dirname(os.getcwd())
sys.path.append(parent_dir)
sys.path.append(os.path.join(parent_dir, "src"))
import sparser
import slack


if __name__ == "__main__":

    # #########################################################################
    # INSPECT USERS INPUT AND SOURCE DIRECTORY:
    print("------------------------------------------------------------------")
    print("Parsing the user's command...")
    # Define argument parser routine:
    arg_parser = argparse.ArgumentParser(
        description="Python script to convert a Slack workspace into "
        + "Excel databases"
        )
    arg_parser.add_argument("--inputs_file_path", required=True, type=str)
    arg_parser.add_argument("--settings_file_path", required=True, type=str)
    args = arg_parser.parse_args()

    print("Parsing the user's inputs in txt files...")
    # Verify that input_file_path exists and parse its content:
    inputs_file_path = os.path.abspath(args.inputs_file_path)
    sparser.check_path_in_user_file("your command prompt.",
                                    "--inputs_file_path", inputs_file_path,
                                    kill=True)
    inputs = sparser.Parser(inputs_file_path)
    source_path = inputs.get("slackexport_folder_path")
    converted_directory = inputs.get("converted_directory")
    ch2analyze = inputs.get("chosen_channel_name")

    # Verify that settings_file_path exists and parse its content:
    settings_file_path = os.path.abspath(args.settings_file_path)
    sparser.check_path_in_user_file("your command prompt.",
                                    "--settings_file_path", settings_file_path,
                                    kill=True)
    settings_messages = sparser.Parser(settings_file_path)
    missing_value = settings_messages.get("missing_value")
    channels_json_name = settings_messages.get("channels_json_name")
    users_json_name = settings_messages.get("users_json_name")
    dest_name_ext = settings_messages.get("dest_name_ext")

    # Check that the source directory exists:
    sparser.check_path_in_user_file(inputs.file_name + ".txt",
                                    "slackexport_folder_path",
                                    source_path,
                                    kill=True)

    # List the Slack channels to be analyzed:
    flag_all_chs = sparser.set_flag_analyze_all_chs(ch2analyze)
    if flag_all_chs is True:
        chs2analyze = sparser.list_dirs_in_path(source_path)
    else:
        channel_path = f"{source_path}/{ch2analyze}"
        if sparser.check_file_exists(channel_path, kill=True) is True:
            chs2analyze = [ch2analyze]

    # Check that the "channels" and "users" JSON files exist:
    sparser.check_file_exists(f"{source_path}/{channels_json_name}",
                              kill=True)
    sparser.check_file_exists(f"{source_path}/{users_json_name}",
                              kill=True)

    # #########################################################################
    # BUILD DATAFRAMES AND WRITE EXCEL FILES:
    print("------------------------------------------------------------------")
    print("Building dataframes and writing Excel files...")

    # Create the path where the files will be saved:
    dest_path = f"{converted_directory}/{dest_name_ext}"
    sparser.make_dest_path(dest_path)

    # Create an instance of the class Slack:
    s = slack.Slack(settings_messages)

    # Get dataframes with the channels info:
    all_channels_df = s.get_all_channels_info(
        source_path, f"{source_path}/{channels_json_name}")

    # Get dataframes with the users info:
    all_users_df = s.get_all_users_info(f"{source_path}/{users_json_name}")

    # Write Excel files if requested by the user:
    slack.write_info_to_file(
        inputs.get("write_all_channels_info"),
        all_channels_df,
        settings_messages.get("channels_excel_name").split(".")[0],
        dest_path)
    slack.write_info_to_file(
        inputs.get("write_all_users_info"),
        all_users_df,
        settings_messages.get("users_excel_name").split(".")[0],
        dest_path)

    # Write the Excel files of the given channel(s):
    channel_messages_df = slack.get_all_messages_df(source_path, dest_path,
                                                    chs2analyze, all_users_df,
                                                    settings_messages)
