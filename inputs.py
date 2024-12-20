from os.path import exists
from pathlib import Path
import shutil



#AG20241119: extra notation to identify a bit more easily the different comments.
#:
##-- If a comment is part of the description of a given step on the code, it starts with ##--
#    if the comment is for suggesting/implementing changes in the code, it starts with #



### <<<   Global variables and settings section >>>
#

##-- uses for debug with Jupiter's-cells-system
rows_to_show = 1   

##-- Syntax to use for missing values:   
missing_value = 'n/d'

# IP20241125 
##-- set adjust for shift from UTC(Slack export timestamp) to ProjectManager's preferred TimeZone 
timmeshift = 'US/Central'  #IP20241125  chose proper value (! string !)  for TimeZone 


##-- Do you wish to convert only one certain Slack channel? then type it's name: f.e. - 'general'  
#    '' - should be preserved for var  initiation  
chosen_channel_name = 'think-biver-sunday-checkins'  
# AG20241220 lines below were directly defined in the function set_flag_analyze_all_channels of the class InspectSource:
#if len(chosen_channel_name) < 1:
#    analyze_all_channels = True 
#    print('Channel(s) to analyze: All')
#else:
#    analyze_all_channels = False
#    print('Channel(s) to analyze: ', chosen_channel_name)

 
##-- Generate file with the information of all the Slack channels?:
write_all_channels_info = True

##-- Generate file with the information of all the Slack users?:
write_all_users_info = True


##-- Insert path where the LOCAL copy of the GoogleDrive folder is:
slackexport_folder_path = "/home/agds/Documents/RebeccaEverleneTrust/RebeccaEverlene_Slack_export" #AG
#slackexport_folder_path = 'E:\_RET_slack_export\RebeccaEverlene Slack export Apr 30 2021 - Oct 3 2024-2short' #IP - to test locally
#slackexport_folder_path = 'E:\_RET_slack_export\RebeccaEverlene Slack export Oct 3 2024 - Nov 9 2024' #IP - to test locally
#slackexport_folder_path = 'E:\_RET_slack_export\RebeccaEverlene Slack export Nov 1 2024 - Nov 30 2024' #IP - to test locally
#slackexport_folder_path = 'E:\_RET_slack_export\RebeccaEverlene Slack export Nov 1 2024 - Nov 30 2024 -shrt' #IP - to test locally

# AG20241220 Lines below were directly defined in the function check_source_path_exists of the class InspectSource
##-- Check that slackexport_folder_path exists:  #IP20241123
#if exists(slackexport_folder_path)==False:
#    print('Please enter a valid path to the source directory')
#    continue_analysis = False        #  IP20241124  may be add here abort of entire code? like "sys.exit()" ?

##-- Insert path where the converted files will be saved:
converted_directory = "/home/agds/Downloads" #AG
#converted_directory = 'E:\_RET_slack_export\RebeccaEverlene Slack export Apr 30 2021 - Oct 3 2024-2short' #IP - to test locally
#converted_directory = 'E:\_RET_slack_export\RebeccaEverlene Slack export Oct 3 2024 - Nov 9 2024' #IP - to test locally
#converted_directory ='E:\_RET_slack_export\RebeccaEverlene Slack export Nov 1 2024 - Nov 30 2024' #IP - to test locally
#converted_directory ='E:\_RET_slack_export\RebeccaEverlene Slack export Nov 1 2024 - Nov 30 2024 -shrt' #IP - to test locally

# AG20241220 Lines below were directly defined in the function check_converted_path_exists of the class InspectSource
#converted_directory = f"{converted_directory}/_JSONs_converted"
##-- Check that     exprt_folder_path  for resulting Excels (JSONs been converted) exists:  #IP20241118
#if exists(converted_directory)==True:
#    exprt_folder_path = Path(converted_directory)
#    if exprt_folder_path.is_dir():
#        print(f"The folder 'JSONs_converted' already exists in '{converted_directory.split('JSONs')[0][:-1]}' and it will be overwritten.") #AG20241120
#        shutil.rmtree(exprt_folder_path)       
#Path(f"{converted_directory}").mkdir(parents=True, exist_ok=True) #IP20241119


#IP20241205
##-- Do you wish to show keywords in the cells with separated weekly-report's parts?:
key_wrd_text_show = False  # True  # 

#
continue_analysis = True      # IP20241123 moved here,to var's initiating section



"""
## Cleaning the lines above leads to:

##-- uses for debug with Jupiter's-cells-system
rows_to_show = 1   

##-- Syntax to use for missing values:   
missing_value = 'n/d'


##-- set adjust for shift from UTC(Slack export timestamp) to ProjectManager's preferred TimeZone 
timmeshift = 'US/Central'  #IP20241125  chose proper value (! string !)  for TimeZone 


##-- Do you wish to convert only one certain Slack channel? then type it's name: f.e. - 'general'  
#    '' - should be preserved for var  initiation  
chosen_channel_name = 'think-biver-sunday-checkins'  

 
##-- Generate file with the information of all the Slack channels?:
write_all_channels_info = True

##-- Generate file with the information of all the Slack users?:
write_all_users_info = True


##-- Insert path where the LOCAL copy of the GoogleDrive folder is:
slackexport_folder_path = "/home/agds/Documents/RebeccaEverleneTrust/RebeccaEverlene_Slack_export" #AG
#slackexport_folder_path = 'E:\_RET_slack_export\RebeccaEverlene Slack export Apr 30 2021 - Oct 3 2024-2short' #IP - to test locally
#slackexport_folder_path = 'E:\_RET_slack_export\RebeccaEverlene Slack export Oct 3 2024 - Nov 9 2024' #IP - to test locally
#slackexport_folder_path = 'E:\_RET_slack_export\RebeccaEverlene Slack export Nov 1 2024 - Nov 30 2024' #IP - to test locally
#slackexport_folder_path = 'E:\_RET_slack_export\RebeccaEverlene Slack export Nov 1 2024 - Nov 30 2024 -shrt' #IP - to test locally


##-- Insert path where the converted files will be saved:
converted_directory = "/home/agds/Downloads" #AG
#converted_directory = 'E:\_RET_slack_export\RebeccaEverlene Slack export Apr 30 2021 - Oct 3 2024-2short' #IP - to test locally
#converted_directory = 'E:\_RET_slack_export\RebeccaEverlene Slack export Oct 3 2024 - Nov 9 2024' #IP - to test locally
#converted_directory ='E:\_RET_slack_export\RebeccaEverlene Slack export Nov 1 2024 - Nov 30 2024' #IP - to test locally
#converted_directory ='E:\_RET_slack_export\RebeccaEverlene Slack export Nov 1 2024 - Nov 30 2024 -shrt' #IP - to test locally


##-- Do you wish to show keywords in the cells with separated weekly-report's parts?:
key_wrd_text_show = False  # True  # 


continue_analysis = True      # IP20241123 moved here,to var's initiating section
"""


