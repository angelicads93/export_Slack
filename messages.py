import pandas as pd
from json import load
import numpy as np
from datetime import datetime
from os import listdir
from os.path import getmtime, exists, isdir, isfile
from pathlib import Path
import shutil
from urlextract import URLExtract
import re

from inputs import missing_value, timmeshift, chosen_channel_name, write_all_channels_info, write_all_users_info, slackexport_folder_path, converted_directory, continue_analysis
import excel


class CleanDF:
    def __init__(self):
        pass
    
    def replace_empty_space(self, df, column):
        """Function to replace empty spaces "" with the string missing_value for a given column"""
        for i in range(len(df)):
            if df.at[i,column] == "":
                df.at[i,column] = missing_value  
                
    def replace_NaN(self, df, column):
        """Function to replace missing values with the string missing_value for a given column """
        df[column] = df[column].fillna(missing_value)    
        
    def handle_missing_values(self, df):
        """Function that replaces missing values in all the columns of the df"""
        df = df.replace(pd.NaT, missing_value)
        df = df.replace(np.nan, missing_value) 
        df = df.fillna(missing_value)
        return df
        
        
        
        
class InspectSource:
    def __init__(self):       
        self.converted_directory = converted_directory
        
    #---------------------------------------------------------
    def set_flag_analyze_all_channels(self):
        if len(chosen_channel_name) < 1:
            analyze_all_channels = True 
            print('Channel(s) to analyze: All')
        else:
            analyze_all_channels = False
            print('Channel(s) to analyze: ', chosen_channel_name)
        return analyze_all_channels
    
    def check_source_path_exists(self):
        if exists(slackexport_folder_path)==False:
            print('Please enter a valid path to the source directory')
            continue_analysis = False        #  IP20241124  may be add here abort of entire codesa? like "sys.exit()" ?
    
    def save_in_path(self):
        return f"{self.converted_directory}/_JSONs_converted"
    
    
    def check_save_path_exists(self, path):
        if exists(path)==True:
            exprt_folder_path = Path(path)
            if exprt_folder_path.is_dir():
                print(f"The path '{path.split('JSONs')[0][:-1]}' already exists and it will be overwritten.") #AG20241120
                shutil.rmtree(exprt_folder_path)
        Path(f"{path}").mkdir(parents=True, exist_ok=True) #IP20241119


    #---------------------------------------------------------
    def check_format_of_json_names(self, list_names):
        """ Iterates over all the json files in a channel's directory, and returns a list with the names of the json files 
        that have the correct format 'yyyy-mm-dd.json' """
        list_names_dates = []
        for i in range(len(list_names)):
            match = re.match(r'(\d{4})(-)(\d{2})(-)(\d{2})(.)(json)',list_names[i])
            if match!=None:
                list_names_dates.append(list_names[i])
        return list_names_dates  
    
    
    def get_channels_names(self):     # AG20241120
        """ Returns a list with the name(s) of the Slack channels to be converted.
        If analysing one channel, check that its directory exists, and default to the 0-th element of channels_names:
        channels_names = [ chosen_channel_name ] for one channel
        channels_names = [channel0, channel1, ...] for all the channels """
        analyze_all_channels = self.set_flag_analyze_all_channels() 
        if analyze_all_channels == False:
            if exists(f"{slackexport_folder_path}/{chosen_channel_name}")==False:
                self.channels_names = []
                print(f"The source directory for the channel '{chosen_channel_name}' was not found in {slackexport_folder_path}")
                continue_analysis = False
            else:
                self.channels_names = [chosen_channel_name]
        else:
            all_in_sourceDir = listdir(slackexport_folder_path)
            self.channels_names = [all_in_sourceDir[i] for i in range(len(all_in_sourceDir)) if isdir(f"{slackexport_folder_path}/{all_in_sourceDir[i]}")==True]
            
        #AG20241120: Pending to check the format of each channel's name. Having empty spaces in the name can cause problems later. 
        return self.channels_names
            
    
    def get_all_channels_json_names(self): 
        """ 
        Check the names of json files in all the channels to be converted and stores them in a list:
        all_channels_jsonFiles_dates = [ [chosen_channel_name_json0, chosen_channel_name_json1, ...] ] for one exportchannel
        all_channels_jsonFiles_dates = [ [channel0_json0, channel0_json1, ...], [channel1_json0, channel1_json1, ...], ... ] for all the channels
        """
        all_channels_jsonFiles_dates = []
        for channel in self.get_channels_names():
            channel_jsonFiles_dates = self.check_format_of_json_names( listdir(f"{slackexport_folder_path}/{channel}") )
            all_channels_jsonFiles_dates.append(channel_jsonFiles_dates)
        return all_channels_jsonFiles_dates
    
    
    def check_missing_channels(self, present_channel_names):   #AG20241127
        ##-- Get name of channels in channels.json:
        expected_channel_names = pd.read_json(f"{slackexport_folder_path}/channels.json")['name'].values
        ##-- Check that all the expected channels are in present channels:
        missing_channels = []
        for channel in expected_channel_names:
            if channel not in present_channel_names:
                missing_channels.append(channel)
        if len(missing_channels) > 0:
            return missing_channels
        else:
            return None

    
    def check_expected_files_exists(self):    
        if exists(slackexport_folder_path)==False:
            print('Please enter a valid path to the source directory')
            continue_analysis = False
        else:
            #  !!! IP2024118  need to check if exist File "channels.json"
            ##-- Check that the channels.json files exists:     # AG20241119:
            if exists(f"{slackexport_folder_path}/channels.json")==False:
                print('File "channels.json" was not found in the source directory')
                continue_analysis = False
             
            ##-- Check that the users.json files exists:
            if exists(f"{slackexport_folder_path}/users.json")==False:
                print('File "users.json" was not found in the source directory')
                continue_analysis = False
    
            ##-- Get a list with the name of the channels to be converted:
            self.channels_names = self.get_channels_names() #AG20241120: Defined routine in function  
    
            #IP20241129 it could be (maybe) useful in further with GUI, but not now)
            ##-- Check for missing channels in the source directory:       #AG20241127
            analyze_all_channels = self.set_flag_analyze_all_channels() 
            if analyze_all_channels == True:
                missing_channels = self.check_missing_channels(self.channels_names)
                if missing_channels != None:
                    print('The following channels are missing in the source directory:', missing_channels)
                    continue_analysis = True
                    #IP20241129  continue_analysis = False    ##AG: pending to prompt the user if continuing with the analysis? (Relevant for the GUI)

            ##-- Get the name of all the json files of the form "yyyy-mm-dd.json" in each channel directory:
            self.all_channels_jsonFiles_dates = self.get_all_channels_json_names() # AG20241120: Defined routine in function
        
        
        
        
        
class SlackChannelsAndUsers:
    def __init__(self):
        self.cleanDF = CleanDF()
        self.inspect_source = InspectSource()
        self.save_path = self.inspect_source.save_in_path()

    def write_info_to_file(self, write_file_flag, filename, df, path):
        if continue_analysis==False:
            print("Please review the input information")
        else:    
            if write_file_flag==True:
                slack_export_user_filename = filename        
                slack_export_user_folder_path_xlsx = f"{path}/{slack_export_user_filename}{'.xlsx'}" #_IP
                df.to_excel(slack_export_user_folder_path_xlsx, index=False) #_IP
                print(datetime.now().time(), f"Wrote file {filename}.xlsx")
                
    def get_all_channels_info(self):
        """
        This function exports the file channels.json into the dataframe all_channels_df and filters/format relevant features.
        The primary features of all_channels_df are: 
            id, name, created, creator, is_archived, is_general, members, pins, topic, purpose.
        The secondary features of 'pins' are:
            id, type, created, user, owner.
            Generally a list of dictionaries.
        The secondary features of 'topic' are:
            value, creator, last_set.
        """
        ##-- Export channels.json to dataframe    
        self.all_channels_df = pd.read_json(f"{slackexport_folder_path}/channels.json")
    
        ##-- Format relevant features on all_channels_df:
        all_json_files = []
        for i in range(len(self.all_channels_df)):
            ##-- Adds df['members']. Writes the list of members into a string separated by commnas:
            tmp_list = self.all_channels_df.at[i, 'members']
            members_str = "".join(f"{tmp_list[j]}, " for j in range(len(tmp_list)))
            self.all_channels_df.at[i,'members'] = members_str[:-2]
            ##-- Adds df['purpose']:
            self.all_channels_df.at[i,'purpose'] = self.all_channels_df.at[i,'purpose']['value']
            ##-- Adds a list with the channel's json_files with the correct format (yyyy-mm-dd.json):
            channel_path = f"{slackexport_folder_path}/{self.all_channels_df.at[i,'name']}"
            
            ##-- Check that the channel_path exists:   #IP20241118
            if exists(channel_path)==True:
                list_names_dates = self.inspect_source.check_format_of_json_names(listdir(channel_path)) #AG20241120: list_names_others not part of the output anymore
                all_json_files.append(list_names_dates)
            else:
                all_json_files.append(missing_value)  
        self.all_channels_df['json_files'] = all_json_files
        
        ##-- Keep the relevant features:
        self.all_channels_df = self.all_channels_df[['id', 'name', 'created', 'creator', 'is_archived', 'is_general', 'members', 'purpose', 'json_files']]
    
        ##-- Handle missing values or empty strings:
        for feature in ['members', 'purpose']:
            self.cleanDF.replace_empty_space(self.all_channels_df, feature)

        ##-- Write all channel's info to .xlsx files, if requested by user:
        self.write_info_to_file(write_all_channels_info, "_all_channels", self.all_channels_df, self.save_path)
    
    
    def get_all_users_info(self):
        """
        This function exports the file users.json into the dataframe all_users_df and filters/format relevant features.
        The primary features of all_users_df are: 
            id, team_id, name, deleted, color, real_name, tz, tz_label, tz_offset, profile, is_admin, is_owner,
            is_primary_owner, is_restricted,is_ultra_restricted, is_bot, is_app_user, updated, is_email_confirmed,
            who_can_share_contact_card, is_invited_user, is_workflow_bot, is_connector_bot.
        Among the secondary features of 'profile', there are:
            title, phone, skype, real_name, real_name_normalized, display_name, display_name_normalized, fields, 
            status_text, status_emoji, status_emoji_display_info, status_expiration, 
            avatar_hash, image_original, is_custom_image, email, huddle_state, huddle_state_expiration_ts, 
            first_name, last_name, image_24, image_32, image_48, image_72, image_192, image_512, image_1024, 
            status_text_canonical, team.
        """
        ##-- Read users.json as a dataframe:
        self.all_users_df = pd.read_json(f"{slackexport_folder_path}/users.json")
        
        ##-- Keep relevant features on all_users_df:
        for i in range(len(self.all_users_df)):
            self.all_users_df.at[i, 'display_name'] = self.all_users_df.at[i, 'profile']['display_name']
            for feature in ['title', 'real_name', 'status_text', 'status_emoji']:
                self.all_users_df.at[i, f"profile_{feature}"] = self.all_users_df.at[i, 'profile'][feature]
        self.all_users_df = self.all_users_df[['id', 'team_id', 'name', 'deleted', 'display_name', 'is_bot', 'profile_title', 'profile_real_name', 
                                     'profile_status_text', 'profile_status_emoji']]
        
        ##-- Handling missing values in all_users_df:
        for feature in ['display_name', 'name', 'team_id', 'id', 'profile_title', 'profile_real_name']:#, 'profile_status_text', 'profile_status_emoji']:
            self.cleanDF.replace_empty_space(self.all_users_df, feature) 
            
        ##-- Write all users's info to .xlsx files, if requested by user:
        self.write_info_to_file(write_all_users_info, "_all_users", self.all_users_df, self.save_path)

        return self.all_users_df




class SlackMessages:
    def __init__(self):
        self.channels_names = InspectSource().get_channels_names()
        self.all_channels_jsonFiles_dates = InspectSource().get_all_channels_json_names()                			
        self.all_users_df = SlackChannelsAndUsers().get_all_users_info()
        self.save_path = InspectSource().save_in_path()   
        
    
    def slack_json_to_dataframe(self, slack_json):
        """ Function to extract channel's messages from a JSON file """
        messages_df = pd.DataFrame(columns=["msg_id", "ts", "user", "type", "text", 
                                            "reply_count", "reply_users_count", 
                                            "ts_latest_reply", "ts_thread", "parent_user_id"])
        for message in range(len(slack_json)):
            #if 'files' in slack_json[message] and slack_json[message]['files']:            #AG:commented out
            #    messages_df.at[message, "msg_id"] = slack_json[message]['files'][0]['id']  #AG:commented out
            if 'client_msg_id' in slack_json[message]:
                messages_df.at[message, "msg_id"] = slack_json[message]['client_msg_id']
            elif 'subtype' in slack_json[message]:                                       #AG:added
                messages_df.at[message, "msg_id"] = slack_json[message]['subtype']       #AG:added
            else:
                messages_df.at[message, "msg_id"] = missing_value #'n/a'
                
            #if 'ts' in slack_json[message]:
            #    messages_df.at[message, "ts"] = slack_json[message]['ts']
            #else:
            #    messages_df.at[message, "ts"] = missing_value  
            
            #messages_df.at[message, "user"] = slack_json[message].get('user', missing_value)  
            
            #if 'text' in slack_json[message]:
            #    messages_df.at[message, "text"] = slack_json[message]['text']
            #else:
            #    messages_df.at[message, "text"] = missing_value  
    
            
            # IP20241124 restored (otherwise missed to store timestamps)
            if 'type' in slack_json[message]:
                messages_df.at[message, "type"] = slack_json[message]['type']
            else:
                messages_df.at[message, "type"] = missing_value  
    
    
            # IP20241124 restored (otherwise missed to store timestamps)
            if 'reply_count' in slack_json[message]:
                #messages_df.at[message, "reply_count"] = slack_json[message]['reply_count']   #AG20241127: line could be deleted if using for loop at the end
                #messages_df.at[message, "reply_users_count"] = slack_json[message]['reply_users_count']  #AG20241127: line could be deleted if using for loop at the end
                messages_df.at[message, "ts_latest_reply"] = slack_json[message]['latest_reply']
            else:
                #messages_df.at[message, "reply_count"] = missing_value   #AG20241127: line could be deleted if using for loop at the end
                #messages_df.at[message, "reply_users_count"] = missing_value  #AG20241127: line could be deleted if using for loop at the end
                messages_df.at[message, "ts_latest_reply"] = missing_value   
            
            # IP20241124 restored (otherwise missed to store timestamps)
            if 'parent_user_id' in slack_json[message]:
                messages_df.at[message, "ts_thread"] = slack_json[message]['thread_ts']
                #messages_df.at[message, "parent_user_id"] = slack_json[message]['parent_user_id']  #AG20241127: line could be deleted if using for loop at the end
                messages_df.at[message, "type"] = "thread"    #IP20241124 to distinguish messages and threads
            else:
                messages_df.at[message, "ts_thread"] = missing_value 
                #messages_df.at[message, "parent_user_id"] = missing_value  #AG20241127: line could be deleted if using for loop at the end
    
            messages_df["text"] = messages_df["text"].astype(str)  #IP20241125  this fixed "FutureWarning: Setting an item of incompatible dtype is deprecated" 
    
            #IP20241125 Replace CR and LF in only the 'text' column  
            #messages_df["text"] = messages_df["text"].apply(lambda x: str(x).replace('\r\n\r\n', '\r\n `rn` ').replace('\r\r', '\r `r` ').replace('\n\n', '\n `n` ') if isinstance(x, str) else x)
            #IP20241125 Replace  CR 
            #messages_df["text"] = messages_df["text"].apply(lambda x: str(x).replace('\n', ' ') if isinstance(x, str) else x)
            #IP20241125 Replace  LF        this chosen as optimal variance
            #messages_df["text"] = messages_df["text"].apply(lambda x: str(x).replace('\r', ' ') if isinstance(x, str) else x)
                
            #AG20241122 simplified commented lines shown above to:
            features = ['ts', 'user',  'text', 'reply_count', 'reply_users_count',  'parent_user_id']  # IP20241124 :: 'type', 'ts_latest_reply', 'ts_thread' - are removed (otherwise missed to store timestamps) 
            for feature in features:
               messages_df.at[message, feature] = slack_json[message].get(feature, missing_value)    
                    
        return messages_df
        
    
    
    
    def get_channel_messages_df(self, export_path, curr_channel_name, json_list):
        """ Extracts all the messages of a given channel from all its JSON files, and stores them on a data frame """
        channel_messages_df = pd.DataFrame(columns=["msg_id", "ts", "user", "type", "text",
                                                    "reply_count", "reply_users_count",
                                                    "ts_latest_reply", "ts_thread", "parent_user_id"])
                                                    # ,"channel_folder", "json_name", "json_mod_date"])          #_IP
        
        ##-- Iterate over JSONs inside the current channel's folder:
        for file_day in range(len(json_list)):
            filejson_path = f"{export_path}/{curr_channel_name}/{json_list[file_day]}" #AG
            
            with open(filejson_path, encoding='utf-8') as f:
                import_file_json = load(f)
            import_file_df = self.slack_json_to_dataframe(import_file_json)
            import_file_df['json_name'] = json_list[file_day]
            import_file_df['json_mod_ts'] = getmtime(filejson_path)  #  un-ZIP of download from Ggl-Drive change ts to the non-sense :: "1980-01-01 00:00:00" 
            
            channel_messages_df = pd.concat([channel_messages_df, import_file_df], axis=0, ignore_index=True) 
        
        channel_messages_df['channel_folder'] = curr_channel_name   #IP
        return channel_messages_df
    
    
    def get_channel_users_df(self, channel_messages_df, users_df ):
        """Returns a data frame with the information of the users in current channel"""
        ##-- Initialize channel_users_df as a copy of users_df:
        channel_users_df = users_df.copy()
        ##-- Find the unique set of users in channel:
        channel_users_list = channel_messages_df['user'].unique()
        ##-- Collect the indices of the users that are NOT in the channel:
        indices_to_drop = [i for i in range(len(users_df)) if users_df.at[i,'id'] not in channel_users_list ]
        ##-- Drop the rows on indices_to_drop:
        channel_users_df.drop(channel_users_df.index[indices_to_drop], inplace=True)
        return channel_users_df

    
    def add_users_info_to_messages(self, df_messages, df_users):
        """Uses the user's id in the format U1234567789 from the df_messages to find the 
        name, display name and if the user is a bot from df_users. 
        The 'name', 'display_name' and 'is_bot' are then added as columns to df_messages"""
        for index in df_messages.index.values:
            i_df = df_users[df_users['id']==df_messages.at[index,'user']]
            ##-- If users in channel_messages_df is not in all_users_df:
            if i_df['display_name'].shape[0]==0 and df_messages.at[index,'user']=='USLACKBOT':        ##AG: 'USLACKBOT' is a special case
                df_messages.at[index, 'name'] =  df_messages.at[index, 'user']
                df_messages.at[index, 'display_name'] =  df_messages.at[index, 'user']
                df_messages.at[index, 'is_bot'] =  True
                df_messages.at[index, 'deactivated'] =  False    #IP20241121
            elif i_df['display_name'].shape[0]==0 and df_messages.at[index,'user']!='USLACKBOT':   
                df_messages.at[index, 'name'] =  "(user not found)"
                df_messages.at[index, 'display_name'] =  "(user not found)"
                df_messages.at[index, 'is_bot'] =  "(user not found)"
                df_messages.at[index, 'deactivated'] =  "(user not found)"
            ##-- If users in channel_messages_df is in all_users_df:
            else:
                df_messages.at[index, 'name'] = i_df['name'].values
                df_messages.at[index, 'display_name'] = i_df['display_name'].values
                df_messages.at[index, 'is_bot'] = i_df['is_bot'].values
                df_messages.at[index, 'deactivated'] =  i_df['deleted'].values  #IP20241121
            del i_df
    
    
    def ts_to_tz(self, df, original_column_name, new_column_name):
        """Transforms timestamps in a dataframe's column to dates on the "US/Central" timezone"""
        df[original_column_name] = pd.to_numeric(df[original_column_name], errors='coerce')   #_IP
        tzs = []
        for i in range(len(df)):
            i_is_null = pd.Series(df.at[i,original_column_name]).isnull().values[0]    #AG20241120
            if i_is_null == True:
                #i_date = '0000-00-00 00:00:00'
                i_date = missing_value
            else:
                # IP20241119    #IP20241125 introduce a var "timmeshift" to adjast timezone from the 1st pfrt of code (easy tocontrol)
                i_date = pd.to_datetime(df.at[i,original_column_name], unit='s').tz_localize('UTC').tz_convert(timmeshift) #('US/Central')
                i_date = datetime.strftime(i_date,"%Y-%m-%d %H:%M:%S")
            tzs.append(i_date)
        df[[original_column_name]].astype('datetime64[s]')
        df[original_column_name] = tzs
        df.rename(columns={original_column_name: new_column_name}, inplace=True)
        
    
    def extract_urls(self, df):
        """Extracts all the url links in df['text'] and stores them as a list in df['URL']"""
        extractor = URLExtract()
        #print('len(df) = ',len(df))  #IP20241125
        for i in range(len(df)):
            urls = []
            urls = extractor.find_urls(df.at[i,'text'])
            #print('i = ', i , 'len(urls)= ', len(urls), 'urls= ', urls)  #IP20241125
            if len(urls)>0:
                urls_string = ' ;  '.join(urls)  #IP20241125  to fix  error_"ValueError: Must have equal len keys and value when setting with an iterable"
                df.at[i,'URL(s)'] = urls_string  #IP20241125 
                #print('i = ', i , 'urls= ', urls)  #IP20241125
            else:
                df.at[i,'URL(s)'] = "" # None   IP2024118
    
    
    #IP20241121 :: AG!  it should be "Add cases where the user_id is not found in users_df." >> like preserve original user_ID and added note "user_not_found"
    #IP20241121 :: AG!  in cases  user's "display_name"=="", then replace "user_ID" with "user_name"
    #IP20241121  ::  AG! :: should Add cases where the user_id is "USLACKBOT" or "SLACKBOT".
    def user_id_to_name(self, df_messages, df_users):
        """Replaces the user_id in the format <@U12345678> to the user's display_name in df_messages['text'], which happens
        when the user is mentioned in an Slack message through the option @user_name. 
         If there is no display_name, then 'user_id' is replaced with 'profile_real_name'.
         All the bots in df_users have an 'id' and 'profile_real_name' (not necessarily 'name' and 'display_id'). Their profile_real_name are:
        Zoom, Google Drive, monday.com, monday.com notifications, GitHub, Google Calendar, Loom, Simple Poll, Figma, 
        OneDrive and SharePoint, Calendly, Outlook Calendar, Rebecca Everlene Trust Company, Slack Team Emoji, New hire onboarding, 
        Welcome, Clockify - Clocking in/out, Zapier, Update Your Slack Team Icon, Jira, Google Sheets, Time Off, Trailhead, 
        Slack Team Emoji Copy, Guru, Guru, Google Calendar, Polly.
         Notice that 'USLACKBOT' and 'B043CSZ0FL7' are the only bot messages if df_messages, but they are not in df_users!
         In the replacements, the "<<>>" are used for clarity on the text, since names can generally have more than one word and many names
        can be referenced one after the other, which can lead to confusion when reading.
        """
        for i in range(len(df_messages)):
            text = df_messages.at[i,'text']
            matches = re.findall(r'<+@[A-Za-z0-9]+>',text)
            if len(matches)>0:
                for match in matches:
                    user = match[2:-1]
                    # AG20241122: begin
                    if user in df_users['id'].values:
                        name = df_users[df_users['id']==user]['display_name'].values[0]
                        is_bot = df_users[df_users['id']==user]['is_bot'].values[0]   
                        if is_bot==True:
                            name = df_users[df_users['id']==user]['profile_real_name'].values[0] + ' (bot)'
                        elif name == missing_value:
                            name = df_users[df_users['id']==user]['profile_real_name'].values[0]
                    else: 
                        name = f"{user} (user not found)"  ## Case for USLACKBOT and B043CSZ0FL7, since they are technically not in df_users!
                    # AG20241122: end
                    text = re.sub(f"<@{user}>", f"@{name}@", text)  #AG20241122: Added "<>" (see function's documentation) 
                    
                    #IP20241121: AG :: should Add cases where the user_id is not found in users_df.
                    #IP20241124:  issue above not solved
    
                    #IP20241121: AG :: should Add cases where the user_id is "USLACKBOT" or "SLACKBOT".
                    #IP20241124:  issue above not solved (or explane - how solved, if solved). Show cell with examples 
                df_messages.at[i,'text'] = text
    
    
    #AG20241122: defined routine that was inside user_id_to_name_test to its own function:
    def parent_user_id_to_name(self, df_messages, df_users):
        # IP20241121   "parent_user_id"  substitution
        '''Replaces the user_id in the format "UA5748HE" to the user's display_name in df_messages['parent_user_id']'''
        for i in range(len(df_messages)):
            #text1 = df_messages.at[i,'parent_user_id']
            #matches = re.findall(r'\bU[A-Za-z0-9]+\b',text1)
            #if len(matches)>0:
            #    for match in matches:
            #        user1 = match   
                    #print("i= ", i, "user1=", user1)   #IP20241121:
            #        if user1 == "SLACKBOT" or user1 == "USLACKBOT":
            #            continue
            #        name1 = df_users[df_users['id']==user1]['display_name'].values[0]
            #        text1 = re.sub(f"{user1}", f"{name1}", text1)
                    #IP20241121: should Add cases where the user_id is not found in users_df.
            #    df_messages.at[i,'parent_user_id'] = text1
    
            #AG20241122: Propose simplifying a bit (since 'matches' will always have the one element in df_messages['parent_user_id'])
            user = df_messages.at[i,'parent_user_id']
            if user!=missing_value:
                name = df_users[df_users['id']==user]['display_name'].values
                if user in df_users['id'].values:
                    is_bot = df_users[df_users['id']==user]['is_bot'].values
                    if is_bot==True:
                        name = df_users[df_users['id']==user]['profile_real_name'].values + ' (bot)'
                    elif name == missing_value:
                        name = df_users[df_users['id']==user]['profile_real_name'].values
                else:
                    name = user+' (user not found)'
                df_messages.at[i,'parent_user_id'] = name
        df_messages.rename(columns={'parent_user_id': 'parent_user_name'}, inplace=True)
            
    
    def channel_id_to_name(self, df_messages, df_users):
        """Replaces <#channel_id|channel_name> to channel_name in df_messages['text'], which happens
        when the channel is mentioned in an Slack message through the option #channel_name"""
        for i in range(len(df_messages)):
            text = df_messages.at[i,'text']
            matches = re.findall(r'#+[A-Za-z0-9]+\|',text)
            if len(matches)>0:
                for match in matches:
                    text = re.sub(match, "", text)
                    text = re.sub(r"<+\|", "<", text)
                df_messages.at[i,'text'] = text
      

        
    def get_all_messages_df(self):
        if continue_analysis==False:
            print("Please review the input information")
        else:    
            ##-- Iterate over channel's folders:
            dfs_list = []
            print(datetime.now().time(), 'Starting loop over channels', '\n')
            for i_channel in range(len(self.channels_names)):
        
                ##-- Define the name of the current channel and the source path containing its json files:
                curr_channel_name = self.channels_names[i_channel] 
                parentfolder_path = f"{slackexport_folder_path}/{curr_channel_name}" 
                print(curr_channel_name, datetime.now().time(), ' Set-up channel name and path to directory')
                
                ##-- Collect all the current_channel's messages in channel_messages_df through the function get_channel_messages_df:
                json_list = self.all_channels_jsonFiles_dates[i_channel]
                channel_messages_df = self.get_channel_messages_df(slackexport_folder_path, curr_channel_name, json_list)  
                print(curr_channel_name, datetime.now().time(), ' Collected channel messages from the json files')
                #
                #IP20241121 move to separate folders-without-messages-JSONs
                if len(channel_messages_df)<1:
                    print("for the folder ",curr_channel_name,"messages_number= ",len(channel_messages_df),"there is no channel's folder", '\n')
                    continue    
        
                ##-- Collect all the users in the current channel through the function get_channel_users_df:
                channel_users_df = self.get_channel_users_df(channel_messages_df, self.all_users_df )
                print(curr_channel_name, datetime.now().time(), ' Collected users in current channel')
                
                ##-- Use channel_users_df to fill-in the user's information in channel_messages_df: 
                self.add_users_info_to_messages(channel_messages_df, channel_users_df)
                print(curr_channel_name, datetime.now().time(), ' Included the users information on channel_messages_df')
                
                ##-- Replace user and team identifiers with their display_names whenever present in a message:
                #user_id_to_name(channel_messages_df, users_df) 
                self.user_id_to_name(channel_messages_df, channel_users_df) 
                self.channel_id_to_name(channel_messages_df, channel_users_df)
                self.parent_user_id_to_name(channel_messages_df, channel_users_df) #AG20241122: routine defined in its own function
                print(curr_channel_name, datetime.now().time(), " User's id replaced by their names in messages")
        
                ##-- Extract hyperlinks from messages, if present (extracted as a list; edit if needed):
                self.extract_urls(channel_messages_df)
                print(curr_channel_name, datetime.now().time(), ' URLs extracted from messages')
        
                ##-- Change format of the time in seconds to a date in the CST time-zone: (Pending 'ts_latest_reply' and 'ts_thread'!)
                #channel_messages_mindate = pd.to_datetime(np.float64(channel_messages_df['ts']), unit='s').min().date()   #AG20241120: Can be deleted
                #channel_messages_maxdate = pd.to_datetime(np.float64(channel_messages_df['ts']), unit='s').max().date()   #AG20241120: Can be deleted
                self.ts_to_tz(channel_messages_df, 'ts', 'msg_date')
                self.ts_to_tz(channel_messages_df, 'json_mod_ts', 'json_mod_date')
                self.ts_to_tz(channel_messages_df, 'ts_latest_reply', 'latest_reply_date')
                self.ts_to_tz(channel_messages_df, 'ts_thread', 'thread_date')
                print('main_analysys ->>',curr_channel_name, "  ", datetime.now().time(), ' Formated the dates and times in the dataframe')
                    
                ##-- Reorder the columns in channel_messages_df:
                #columns_order = ['msg_id', 'msg_date', 'user', 'name', 'display_name', 'deactivated', 'is_bot', 'type', 'text', 'reply_count', 'reply_users_count', 'latest_reply_date', 'thread_date', 'parent_user_name', 'URL(s)']
                #channel_messages_df = channel_messages_df[columns_order]
                
                ##-- Sort the dataframe by msg_date:
                channel_messages_df.sort_values(by='msg_date', inplace=True, ignore_index=True)
                
                ##-- Write channel_messages_df to a .xlsx file:
                channel_messages_mindate = channel_messages_df['msg_date'].min().split(" ")[0]
                channel_messages_maxdate = channel_messages_df['msg_date'].max().split(" ")[0]
                channel_messages_maxdate = channel_messages_df['msg_date'].max().split(" ")[0]
                channel_messages_filename = f"{curr_channel_name}_{channel_messages_mindate}_to_{channel_messages_maxdate}"
                channel_messages_folder_path = f"{self.save_path}/{channel_messages_filename}.xlsx"
                channel_messages_df.to_excel(f"{channel_messages_folder_path}", index=False)
                excel.ExcelFormat(channel_messages_folder_path, curr_channel_name).IP_excel_adjustments() 
                print(curr_channel_name, datetime.now().time(), ' Wrote curated messages to xlsx files', '\n')
        
                dfs_list.append(channel_messages_df)
                
        print(datetime.now().time(), 'Done')
        
        return channel_messages_df
        
        
