import inputs

# --Change current_working_directory to src/ and access all the modules:
import os, sys
sys.path.append(os.path.join(os.getcwd(), 'src'))
from src import messages

slackexport_folder_path = inputs.slackexport_folder_path

# --Initialize constructor of the class InspectSource:
inspect_source = messages.InspectSource(
    inputs.chosen_channel_name,
    inputs.slackexport_folder_path,
    inputs.converted_directory)
# --Check validity of input paths:
analyze_all_channels = inspect_source.set_flag_analyze_all_channels()
inspect_source.check_source_path_exists()
save_in_path = inspect_source.save_in_path()
inspect_source.check_save_path_exists(save_in_path)
# --Check for expected files:
inspect_source.check_expected_files_exists()
# --Retrieve variables:
channels_names = inspect_source.get_channels_names()
all_channels_jsonFiles_dates = inspect_source.all_channels_jsonFiles_dates


# --Initialize constructor of the class SlackChannelAndUsers:
scu = messages.SlackChannelsAndUsers(
    inputs.chosen_channel_name,
    inputs.write_all_channels_info,
    inputs.write_all_users_info,
    inputs.slackexport_folder_path,
    inputs.converted_directory)
# --Execute the main functions of the class:
scu.get_all_channels_info()
scu.get_all_users_info()
# --Retrieve variables:
all_channels_df = scu.all_channels_df
all_users_df = scu.all_users_df


# --Initialize constructor of the class SlackMessages:
sm = messages.SlackMessages(
    inputs.chosen_channel_name,
    inputs.write_all_channels_info, inputs.write_all_users_info,
    inputs.slackexport_folder_path, inputs.converted_directory)
# --Execute the main function of the class:
channel_messages_df = sm.get_all_messages_df()
