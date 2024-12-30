from inputs import chosen_channel_name
import messages
import checkins
import excel

##############################################################################################################################

##-- Initialize constructor of the class InspectSource:
inspect_source = messages.InspectSource()
##-- Check validity of input paths:
analyze_all_channels = inspect_source.set_flag_analyze_all_channels()
inspect_source.check_source_path_exists()
save_in_path = inspect_source.save_in_path()
inspect_source.check_save_path_exists(save_in_path)
##-- Check for expected files:
inspect_source.check_expected_files_exists()
##-- Retrieve variables:
channels_names = inspect_source.channels_names
all_channels_jsonFiles_dates = inspect_source.all_channels_jsonFiles_dates

##-- Initialize constructor of the class SlackChannelAndUsers:
scu = messages.SlackChannelsAndUsers()
##-- Execute the main functions of the class:
scu.get_all_channels_info()
scu.get_all_users_info()
##-- Retrieve variables:
all_channels_df = scu.all_channels_df
all_users_df = scu.all_users_df

##-- Initialize constructor of the class SlackMessages:
sm = messages.SlackMessages()
##-- Execute the main function of the class:
channel_messages_df = sm.get_all_messages_df()


##############################################################################################################################

###--- Applying _TEST Excel format:

##-- Select and reorder the dataframe columns:
columns_order = ['msg_id', 'msg_date', 'user', 'name', 'display_name', 'deactivated', 'is_bot', 'type', 'text', 'reply_count', 'reply_users_count', 'latest_reply_date', 'thread_date', 'parent_user_name', 'URL(s)']
df_TEST = channel_messages_df[columns_order]

##-- Parse Check-in messages:
parseDF = checkins.SlackCheckins()
df_TEST = parseDF.parse_nrows(df_TEST)

##-- Reorder columns:
column_names_1row = ['msg_id', 'msg_date', 'user', 'name', 'display_name', 'deactivated',
       'is_bot', 'type', 'text', 'reply_count', 'reply_users_count',
       'latest_reply_date', 'thread_date', 'parent_user_name', 'URL(s)','projects_parsed',
       'project_name_1', 'working_on_1', 'progress_and_roadblocks_1',
       'progress_1', 'roadblocks_1', 'plans_for_following_week_1',
       'meetings_1', 'project_name_2', 'working_on_2',
       'progress_and_roadblocks_2', 'progress_2', 'roadblocks_2',
       'plans_for_following_week_2', 'meetings_2', 'project_name_3',
       'working_on_3', 'progress_and_roadblocks_3', 'progress_3',
       'roadblocks_3', 'plans_for_following_week_3', 'meetings_3']
column_names_nrows = ['msg_id', 'msg_date', 'user', 'name', 'display_name', 'deactivated',
       'is_bot', 'type', 'text', 'reply_count', 'reply_users_count',
       'latest_reply_date', 'thread_date', 'parent_user_name', 'URL(s)','projects_parsed',
       'project_name', 'working_on', 'progress_and_roadblocks','progress', 'roadblocks', 'plans_for_following_week','meetings']

df_TEST = df_TEST[column_names_nrows]


##-- Save dataframe to Excel file:
save_path = messages.InspectSource().save_in_path()
channel_messages_mindate = df_TEST['msg_date'].min().split(" ")[0]
channel_messages_maxdate = df_TEST['msg_date'].max().split(" ")[0]
channel_messages_filename = f"{chosen_channel_name}_{channel_messages_mindate}_to_{channel_messages_maxdate}_TEST"
channel_messages_folder_path = f"{save_path}/{channel_messages_filename}.xlsx"
df_TEST.to_excel(f"{channel_messages_folder_path}", index=False)

##-- Format Excel file (Parsing of check-in messages in apply_excel_adjustments not done here):
ef = excel.ExcelFormat(f"{channel_messages_folder_path}",chosen_channel_name)
ef.set_cell_width()
ef.set_cell_color()
ef.set_font_color()
ef.set_cell_allignment()
ef.set_format_first_row()
ef.rename_sheet()
ef.save_changes()

