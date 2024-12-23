from os.path import exists
from pathlib import Path
import shutil


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




