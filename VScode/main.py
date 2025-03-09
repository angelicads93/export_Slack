#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import argparse

parent_dir = os.path.dirname(os.getcwd())
sys.path.append(parent_dir)
sys.path.append(os.path.join(parent_dir, 'src'))
import messages

if __name__ == "__main__":

    # --Define argument parser routine:
    parser = argparse.ArgumentParser(
        description="Python script to convert a Slack workspace into Excel databases"
        )
    parser.add_argument("--inputs_file_path", required=True, type=str)
    parser.add_argument("--settings_file_path", required=True, type=str)
    args = parser.parse_args()

    inputs_file_path = args.inputs_file_path
    inputs = os.path.basename(inputs_file_path).split(".")[0]
    print(f"inputs_file_path = {inputs_file_path}")
    if os.path.exists(inputs_file_path) is False:
        print(f"ERROR: Path {inputs_file_path} does not exists." + "/n"
              + "       Please review your input for the argument "
              + "--inputs_file_path.")
        sys.exit()

    settings_file_path = args.settings_file_path
    settings_messages = os.path.basename(settings_file_path).split(".")[0]
    print(f"settings_file_path = {settings_file_path}")
    if os.path.exists(settings_file_path) is False:
        print(f"ERROR: Path {settings_file_path} does not exists." + "/n"
              + "       Please review your input for the argument"
              + "--settings_file_path.")
        sys.exit()

    # --Initialize constructor of the class InspectSource:
    inspect_source = messages.InspectSource(inputs, settings_messages)
    # --Check validity of input paths:
    analyze_all_channels = inspect_source.set_flag_analyze_all_channels()
    inspect_source.check_source_path_exists()
    save_in_path = inspect_source.save_in_path()
    inspect_source.check_save_path_exists(save_in_path)
    # --Check for expected files:
    inspect_source.check_expected_files_exists()
    # --Retrieve variables:
    channels_names = inspect_source.get_channels_names()
    all_channels_jsonFiles_dates = inspect_source.get_all_channels_json_names()

    # --Initialize constructor of the class SlackChannelAndUsers:
    scu = messages.SlackChannelsAndUsers(inputs, settings_messages)
    # --Execute the main functions of the class:
    scu.get_all_channels_info()
    scu.get_all_users_info()
    # --Retrieve variables:
    all_channels_df = scu.all_channels_df
    all_users_df = scu.all_users_df

    # --Initialize constructor of the class SlackMessages:
    sm = messages.SlackMessages(inputs, settings_messages)
    # --Execute the main function of the class:
    channel_messages_df = sm.get_all_messages_df()
