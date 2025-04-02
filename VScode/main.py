#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import argparse

# Include the main repo directory (export_Slack/) and the src directory
# (export_Slack/src) to the Python path so the customed modules can be imported
parent_dir = os.path.dirname(os.getcwd())
sys.path.append(parent_dir)
sys.path.append(os.path.join(parent_dir, 'src'))
import messages
import parser

if __name__ == "__main__":

    # #########################################################################
    # INSPECT SOURCE DIRECTORY:
    print('------------------------------------------------------------------')
    print("Parsing the user's input...")
    # --Define argument parser routine:
    arg_parser = argparse.ArgumentParser(
        description="Python script to convert a Slack workspace into "
        + "Excel databases"
        )
    arg_parser.add_argument("--inputs_file_path", required=True, type=str)
    arg_parser.add_argument("--settings_file_path", required=True, type=str)
    args = arg_parser.parse_args()

    inputs_file_path = os.path.abspath(args.inputs_file_path)
    print(f"inputs_file_path = {inputs_file_path}")
    # --Verify that input_file_path exists:
    if os.path.exists(inputs_file_path) is False:
        print(f"ERROR: Path {inputs_file_path} does not exists." + "/n"
              + "       Please review your input for the argument "
              + "--inputs_file_path.")
        sys.exit()
    # --Parse the information in input_file_path:
    inputs = parser.Parser(inputs_file_path)

    settings_file_path = os.path.abspath(args.settings_file_path)
    print(f"settings_file_path = {settings_file_path}")
    # --Verify that settings_file_path exists:
    if os.path.exists(settings_file_path) is False:
        print(f"ERROR: Path {settings_file_path} does not exists." + "/n"
              + "       Please review your input for the argument"
              + "--settings_file_path.")
        sys.exit()
    # --Parse the information in settings_file_path:
    settings_messages = parser.Parser(settings_file_path)


    # #########################################################################
    # INSPECT SOURCE DIRECTORY:
    print('------------------------------------------------------------------')
    print('Inspecting the source directory...')
    inspect_source = messages.InspectSource(inputs, settings_messages)
    
    # Check that the source directory exists:
    inspect_source.check_src_path_exists()
    
    # Check that the "channels" and "users" JSON files exists:
    inspect_source.check_channels_json_exists()
    inspect_source.check_users_json_exists()
    
    # Create the path where the files will be saved:    
    inspect_source.make_dest_path()
    
    # Get a list with the channel(s) to be converted (and present in the source
    # directory):
    channels_names = inspect_source.present_chs
    
    # If analyzing all the channels, checks for expected channels that are not
    # in the source directory:
    inspect_source.check_missing_chs()
    
    # Get a list with the name of the JSON files with the channel(s) messages:
    all_channels_jsonFiles_dates = inspect_source.chs_jsons

    # Retrieve the variable "continue_analysis":
    continue_analysis = inspect_source.get_cont_analysis()
    print(f"## continue_analysis = {continue_analysis}")


    # #########################################################################
    # GET SLACK CHANNELS AND USERS:
    print('------------------------------------------------------------------')
    print('Retrieving information on the Slack channels and users...')
    # Initialize constructor of the class SlackChannelAndUsers:
    scu = messages.SlackChannelsAndUsers(inputs, settings_messages)
    
    # Get dataframes with the channels and users info:
    all_channels_df = scu.get_all_channels_info()
    all_users_df = scu.get_all_users_info()
    
    # Write Excel files if requested by the user:
    scu.write_info_to_file(
        inputs.get('write_all_channels_info'),
        all_channels_df,
        settings_messages.get('channels_excel_name').split(".")[0],
        inspect_source.get_dest_path())
    scu.write_info_to_file(
        inputs.get('write_all_users_info'),
        all_users_df,
        settings_messages.get('users_excel_name').split(".")[0],
        inspect_source.get_dest_path())

    # #########################################################################
    # GET SLACK MESSAGES:
    print('------------------------------------------------------------------')
    print('Retrieving information on the Slack messages...')
    # Initialize constructor of the class SlackMessages:
    sm = messages.SlackMessages(inputs, settings_messages)
    
    # Execute the main function of the class:
    channel_messages_df = sm.get_all_messages_df()
