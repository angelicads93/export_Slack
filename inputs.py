#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# --If you wish to analyze one Slack channel, enter the channel name.
# --If you wish to analyze all the Slack channels, enter ''.
chosen_channel_name = 'general'

# --Do you wish to generate the file with the information of all the
# --Slack channels?:
write_all_channels_info = False

# --Do you wish to generate the file with the information of all the
# --Slack users?:
write_all_users_info = False

# --Insert absolute path to the LOCAL copy of the GoogleDrive folder.
# Note: To get the absolute path in Windows, 
#  1. Navigate to the desire folder with File Explorer.
#  2. Right-click on the name of the folder.
#  3. Select "Copy as path".
#  4. Paste the path between the quotation marks bellow.
#slackexport_folder_path = r"absolute_path_to_source_directory"
slackexport_folder_path = r"/home/agds/Documents/RET/Source/Oct_3_2024_to_Nov_9_2024"

# --Insert path where the converted files will be saved.
# Note: Enter r"" to save your files in the source directory.
#converted_directory = r"absolute_path_to_destination_directory"
converted_directory = r"/home/agds/Desktop"
