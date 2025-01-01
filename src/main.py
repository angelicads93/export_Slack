import sys
sys.path.append("../")
import messages

# #-- Initialize constructor of the class InspectSource:
inspect_source = messages.InspectSource()
# #-- Check validity of input paths:
analyze_all_channels = inspect_source.set_flag_analyze_all_channels()
inspect_source.check_source_path_exists()
save_in_path = inspect_source.save_in_path()
inspect_source.check_save_path_exists(save_in_path)
# #-- Check for expected files:
inspect_source.check_expected_files_exists()
# #-- Retrieve variables:
channels_names = inspect_source.channels_names
all_channels_jsonFiles_dates = inspect_source.all_channels_jsonFiles_dates


# #-- Initialize constructor of the class SlackChannelAndUsers:
scu = messages.SlackChannelsAndUsers()
# #-- Execute the main functions of the class:
scu.get_all_channels_info()
scu.get_all_users_info()
# #-- Retrieve variables:
all_channels_df = scu.all_channels_df
all_users_df = scu.all_users_df


# #-- Initialize constructor of the class SlackMessages:
sm = messages.SlackMessages()
# #-- Execute the main function of the class:
channel_messages_df = sm.get_all_messages_df()
