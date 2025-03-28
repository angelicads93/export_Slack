#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
from json import load
from datetime import datetime
from os import listdir, getcwd
from os.path import getmtime, exists, isdir, dirname
from pathlib import Path
import sys
import importlib
import shutil
from urlextract import URLExtract
import re

import excel
import clean
import checkins

parent_dir = dirname(getcwd())
if parent_dir not in sys.path:
    sys.path.append(parent_dir)


class InspectSource:
    def __init__(self, inputs, settings_messages):

        self.inputs = importlib.import_module(inputs)
        self.chosen_channel_name = self.inputs.chosen_channel_name
        self.slackexport_folder_path = self.inputs.slackexport_folder_path
        self.converted_directory = self.inputs.converted_directory

        self.settings_messages = importlib.import_module(settings_messages)
        self.dest_name_ext = self.settings_messages.dest_name_ext
        self.channels_json_name = self.settings_messages.channels_json_name
        self.users_json_name = self.settings_messages.users_json_name

    def set_flag_analyze_all_channels(self):
        """ Sets a flag to keep track if one or all the channels are going to
        be analyzed.
        """
        if len(self.chosen_channel_name) < 1:
            analyze_all_channels = True
            print('Channel(s) to analyze: All')
        else:
            analyze_all_channels = False
            print('Channel(s) to analyze: ', self.chosen_channel_name)
        return analyze_all_channels

    def check_source_path_exists(self):
        """ Verifies that the specified path with the source data exists """
        if exists(self.slackexport_folder_path) is False:
            status = 'Please enter a valid path to the source directory'
            print(status)
            self.continue_analysis = False
        else:
            status = ''
        return ''

    def save_in_path(self):
        """ Adds the name extension to the specified path where
        to save the results.
        """
        return f"{self.converted_directory}/{self.dest_name_ext}"

    def check_save_path_exists(self, path):
        """ Checks that the specified path where to store the output
        information exists.
        """
        if exists(path) is True:
            exprt_folder_path = Path(path)
            if exprt_folder_path.is_dir():
                print(
                    f"The path '{path} already exists and it will be replaced."
                    )
                # chmod(exprt_folder_path, stat.S_IRWXU)
                shutil.rmtree(exprt_folder_path)
        Path(f"{path}").mkdir(parents=True, exist_ok=True)

    def check_format_of_json_names(self, list_names):
        """ Iterates over all the json files in a channel's directory, and
        returns a list with the names of the json files that have the correct
        format 'yyyy-mm-dd.json'
        """
        list_names_dates = []
        for i in range(len(list_names)):
            match = re.match(
                r'(\d{4})(-)(\d{2})(-)(\d{2})(.)(json)', list_names[i]
                )
            if match is not None:
                list_names_dates.append(list_names[i])
        return list_names_dates

    def get_channels_names(self):
        """ Returns list with name(s) of Slack channels to be converted.
        If analysing one channel, check that its directory exists, and default
        to the 0-th element of channels_names:
        channels_names = [ chosen_channel_name ] for one channel
        channels_names = [channel0, channel1, ...] for all the channels
        """
        analyze_all_channels = self.set_flag_analyze_all_channels()
        if analyze_all_channels is False:
            if exists(
                    f"{self.slackexport_folder_path}/{self.chosen_channel_name}"
                    ) is False:
                self.channels_names = []
                print("The source directory for the channel "
                      + f"'{self.chosen_channel_name}' was not found in "
                      + f"{self.slackexport_folder_path}")
                self.continue_analysis = False
            else:
                self.channels_names = [self.chosen_channel_name]
        else:
            all_in_sourceDir = listdir(self.slackexport_folder_path)
            self.channels_names = [all_in_sourceDir[i]
                                   for i in range(len(all_in_sourceDir))
                                   if isdir(f"{self.slackexport_folder_path}/{all_in_sourceDir[i]}") is True]

        return self.channels_names

    def get_all_channels_json_names(self):
        """ Check the names of json files in all the channels to be converted
        and stores them in a list:
        all_channels_jsonFiles_dates = [
            [chosen_channel_name_json0, chosen_channel_name_json1, ...]
            ] for one exportchannel
        all_channels_jsonFiles_dates = [
            [channel0_json0, channel0_json1, ...],
            [channel1_json0, channel1_json1, ...], ...
            ] for all the channels
        """
        all_channels_jsonFiles_dates = []
        for channel in self.get_channels_names():
            channel_jsonFiles_dates = self.check_format_of_json_names(
                listdir(f"{self.slackexport_folder_path}/{channel}")
                )
            all_channels_jsonFiles_dates.append(channel_jsonFiles_dates)
        return all_channels_jsonFiles_dates

    def check_missing_channels(self, present_channel_names):
        """ Compares the channels present in the source directory with the
        channels expected from the file channels.json. If they do not agree,
        returns a list with the name of te files that are missing.
        """
        # --Get name of channels in channels.json:
        expected_channel_names = pd.read_json(
            f"{self.slackexport_folder_path}/{self.channels_json_name}"
            )['name'].values
        # --Check that all the expected channels are in present channels:
        missing_channels = []
        for channel in expected_channel_names:
            if channel not in present_channel_names:
                missing_channels.append(channel)
        if len(missing_channels) > 0:
            return missing_channels
        else:
            return None

    def check_expected_files_exists(self):
        """" Makes sure that all the files requires for the analysis exists.
        If files are missing, error messages are printed.
        """
        if exists(self.slackexport_folder_path) is False:
            print('Please enter a valid path to the source directory')
            self.continue_analysis = False
        else:
            # --Check that the channels.json files exists:
            if exists(f"{self.slackexport_folder_path}/{self.channels_json_name}") is False:
                print(f'File {self.channels_json_name} was not found in the '
                      + 'source directory')
                self.continue_analysis = False
            # --Check that the users.json files exists:
            if exists(f"{self.slackexport_folder_path}/{self.users_json_name}") is False:
                print(f'File "{self.users_json_name}" was not found in the '
                      + 'source directory')
                self.continue_analysis = False
            # --Get a list with the name of the channels to be converted:
            self.channels_names = self.get_channels_names()
            # --Check for missing channels in the source directory:
            analyze_all_channels = self.set_flag_analyze_all_channels()
            if analyze_all_channels is True:
                missing_channels = self.check_missing_channels(
                    self.channels_names
                    )
                if missing_channels is not None:
                    print(
                        """The following channels are missing in the source
                        directory:""", missing_channels
                        )
                    self.continue_analysis = True

            # --Get the name of all the json files of the form
            # --"yyyy-mm-dd.json" in each channel directory:
            self.all_channels_jsonFiles_dates = self.get_all_channels_json_names()


class SlackChannelsAndUsers:
    def __init__(self,  inputs, settings_messages):

        self.inputs = importlib.import_module(inputs)
        self.write_all_channels_info = self.inputs.write_all_channels_info
        self.write_all_users_info = self.inputs.write_all_users_info
        self.slackexport_folder_path = self.inputs.slackexport_folder_path
        self.converted_directory = self.inputs.converted_directory

        self.settings_messages = importlib.import_module(settings_messages)
        self.missing_value = self.settings_messages.missing_value
        self.timezone = self.settings_messages.timezone
        self.dest_name_ext = self.settings_messages.dest_name_ext
        self.channels_json_name = self.settings_messages.channels_json_name
        self.users_json_name = self.settings_messages.users_json_name
        self.channels_excel_name = self.settings_messages.channels_excel_name
        self.users_excel_name = self.settings_messages.users_excel_name

        self.inspect_source = InspectSource("inputs", "settings_messages")
        self.save_path = self.inspect_source.save_in_path()
        self.continue_analysis = True

    def write_info_to_file(self, write_file_flag, filename, df, path):
        """ Writes a given dataframe to an Excel file """
        if self.continue_analysis is False:
            print("Please review the input information")
        else:
            if write_file_flag is True:
                slack_export_user_filename = filename
                slack_export_user_folder_path_xlsx = f"{path}/{slack_export_user_filename}{'.xlsx'}"
                df.to_excel(slack_export_user_folder_path_xlsx, index=False)
                print(datetime.now().time(), f"Wrote file {filename}.xlsx")

    def get_all_channels_info(self):
        """ Exports the file channels.json into the dataframe all_channels_df
        and filters/format relevant features.
        The primary features of all_channels_df are:
            id, name, created, creator, is_archived, is_general, members, pins,
            topic, purpose.
        The secondary features of 'pins' are:
            id, type, created, user, owner.
            Generally a list of dictionaries.
        The secondary features of 'topic' are:
            value, creator, last_set.
        """
        # --Export channels.json to dataframe
        self.all_channels_df = pd.read_json(
            f"{self.slackexport_folder_path}/{self.channels_json_name}"
            )

        # --Format relevant features on all_channels_df:
        all_json_files = []
        for i in range(len(self.all_channels_df)):
            # --Adds df['members']. Writes the list of members into a string
            # --separated by commnas:
            tmp_list = self.all_channels_df.at[i, 'members']
            members_str = "".join(
                f"{tmp_list[j]}, " for j in range(len(tmp_list))
                )
            self.all_channels_df.at[i, 'members'] = members_str[:-2]
            # --Adds df['purpose']:
            self.all_channels_df.at[i, 'purpose'] = self.all_channels_df.at[i, 'purpose']['value']
            # --Adds a list with the channel's json_files with the correct
            # --format (yyyy-mm-dd.json):
            channel_path = f"{self.slackexport_folder_path}/{self.all_channels_df.at[i, 'name']}"

            # --Check that the channel_path exists:
            if exists(channel_path) is True:
                list_names_dates = self.inspect_source.check_format_of_json_names(listdir(channel_path))
                all_json_files.append(list_names_dates)
            else:
                all_json_files.append(self.missing_value)
        self.all_channels_df['json_files'] = all_json_files

        # --Keep the relevant features:
        self.all_channels_df = self.all_channels_df[
            ['id', 'name', 'created', 'creator', 'is_archived', 'is_general',
             'members', 'purpose', 'json_files']
            ]

        # --Handle missing values or empty strings:
        for feature in ['members', 'purpose']:
            clean.replace_empty_space(self.all_channels_df, feature, self.missing_value)

        # --Write all channel's info to .xlsx files, if requested by user:
        self.write_info_to_file(
            self.write_all_channels_info, self.channels_excel_name.split(".")[0],
            self.all_channels_df, self.save_path)

    def get_all_users_info(self):
        """ Exports the file users.json into the dataframe all_users_df and
        filters/format relevant features.
        The primary features of all_users_df are:
            id, team_id, name, deleted, color, real_name, tz, tz_label,
            tz_offset, profile, is_admin, is_owner, is_primary_owner,
            is_restricted,is_ultra_restricted, is_bot, is_app_user, updated,
            is_email_confirmed, who_can_share_contact_card, is_invited_user,
            is_workflow_bot, is_connector_bot.
        Among the secondary features of 'profile', there are:
            title, phone, skype, real_name, real_name_normalized, display_name,
            display_name_normalized, fields,status_text, status_emoji,
            status_emoji_display_info, status_expiration, avatar_hash,
            image_original, is_custom_image, email, huddle_state,
            huddle_state_expiration_ts, first_name, last_name, image_24,
            image_32, image_48, image_72, image_192, image_512, image_1024,
            status_text_canonical, team.
        """
        # --Read users.json as a dataframe:
        self.all_users_df = pd.read_json(
            f"{self.slackexport_folder_path}/{self.users_json_name}"
            )

        # --Keep relevant features on all_users_df:
        for i in range(len(self.all_users_df)):
            self.all_users_df.at[i, 'display_name'] = self.all_users_df.at[i, 'profile']['display_name']
            for feature in [
                    'title', 'real_name', 'status_text', 'status_emoji'
                    ]:
                self.all_users_df.at[i, f"profile_{feature}"] = self.all_users_df.at[i, 'profile'][feature]
        self.all_users_df = self.all_users_df[
            ['id', 'team_id', 'name', 'deleted', 'display_name', 'is_bot',
             'profile_title', 'profile_real_name', 'profile_status_text',
             'profile_status_emoji']
            ]

        # --Handling missing values in all_users_df:
        for feature in ['display_name', 'name', 'team_id', 'id',
                        'profile_title', 'profile_real_name']:
            clean.replace_empty_space(self.all_users_df, feature, self.missing_value)

        # --Write all users's info to .xlsx files, if requested by user:
        self.write_info_to_file(
            self.write_all_users_info, self.users_excel_name.split(".")[0],
            self.all_users_df, self.save_path
            )

        return self.all_users_df


class SlackMessages:
    def __init__(self, inputs, settings_messages):

        self.inputs = importlib.import_module(inputs)
        self.slackexport_folder_path = self.inputs.slackexport_folder_path

        self.settings_messages = importlib.import_module(settings_messages)
        self.missing_value = self.settings_messages.missing_value
        self.timezone = self.settings_messages.timezone

        self.inspect_source = InspectSource("inputs", "settings_messages")
        self.channels_names = self.inspect_source.get_channels_names()
        self.all_channels_jsonFiles_dates = self.inspect_source.get_all_channels_json_names()
        self.save_path = self.inspect_source.save_in_path()

        self.slack_channels_users = SlackChannelsAndUsers(
            "inputs", "settings_messages")
        self.all_users_df = self.slack_channels_users.get_all_users_info()
        self.continue_analysis = self.slack_channels_users.continue_analysis

    def slack_json_to_dataframe(self, slack_json):
        """ Extracts channel's messages from a JSON file """
        msgs_df = pd.DataFrame(
            columns=["msg_id", "ts", "user", "type", "text", "reply_count",
                     "reply_users_count", "ts_latest_reply", "ts_thread",
                     "parent_user_id"]
            )
        for msg in range(len(slack_json)):
            if 'client_msg_id' in slack_json[msg]:
                msgs_df.at[msg, "msg_id"] = slack_json[msg]['client_msg_id']
            elif 'subtype' in slack_json[msg]:
                msgs_df.at[msg, "msg_id"] = slack_json[msg]['subtype']
            else:
                msgs_df.at[msg, "msg_id"] = self.missing_value

            if 'type' in slack_json[msg]:
                msgs_df.at[msg, "type"] = slack_json[msg]['type']
            else:
                msgs_df.at[msg, "type"] = self.missing_value

            if 'reply_count' in slack_json[msg]:
                msgs_df.at[msg, "ts_latest_reply"] = slack_json[msg]['latest_reply']
            else:
                msgs_df.at[msg, "ts_latest_reply"] = self.missing_value

            if 'parent_user_id' in slack_json[msg]:
                msgs_df.at[msg, "ts_thread"] = slack_json[msg]['thread_ts']
                msgs_df.at[msg, "type"] = "thread"
            else:
                msgs_df.at[msg, "ts_thread"] = self.missing_value

            msgs_df["text"] = msgs_df["text"].astype(str)

            features = ['ts', 'user',  'text', 'reply_count',
                        'reply_users_count',  'parent_user_id']
            for feature in features:
                msgs_df.at[msg, feature] = slack_json[msg].get(feature, self.missing_value)

        return msgs_df

    def get_ch_msgs_df(self, export_path, curr_ch_name, json_list):
        """ Extracts all the messages of a given channel from all its JSON
        files, and stores them on a data frame
        """
        ch_msgs_df = pd.DataFrame(
            columns=["msg_id", "ts", "user", "type", "text", "reply_count",
                     "reply_users_count", "ts_latest_reply", "ts_thread",
                     "parent_user_id"]
            )

        # --Iterate over JSONs inside the current channel's folder:
        for file_day in range(len(json_list)):
            filejson_path = f"{export_path}/{curr_ch_name}/{json_list[file_day]}"

            with open(filejson_path, encoding='utf-8') as f:
                import_file_json = load(f)
            import_file_df = self.slack_json_to_dataframe(import_file_json)
            import_file_df['json_name'] = json_list[file_day]
            import_file_df['json_mod_ts'] = getmtime(filejson_path)

            ch_msgs_df = pd.concat(
                [ch_msgs_df, import_file_df],
                axis=0, ignore_index=True
                )

        ch_msgs_df['channel_folder'] = curr_ch_name
        return ch_msgs_df

    def get_channel_users_df(self, ch_msgs_df, users_df):
        """ Returns a data frame with the information of the users in current
        channel
        """
        # --Initialize channel_users_df as a copy of users_df:
        channel_users_df = users_df.copy()
        # --Find the unique set of users in channel:
        channel_users_list = ch_msgs_df['user'].unique()
        # --Collect the indices of the users that are NOT in the channel:
        indices_to_drop = [i
                           for i in range(len(users_df))
                           if users_df.at[i, 'id'] not in channel_users_list]
        # --Drop the rows on indices_to_drop:
        channel_users_df.drop(
            channel_users_df.index[indices_to_drop], inplace=True
            )
        return channel_users_df

    def add_users_info_to_messages(self, df_messages, df_users):
        """Uses the user's id in the format U1234567789 from the df_messages to
        find the name, display name and if the user is a bot from df_users.
        The 'name', 'display_name' and 'is_bot' are then added as columns to
        df_messages
        """
        for index in df_messages.index.values:
            i_df = df_users[df_users['id'] == df_messages.at[index, 'user']]
            # --If users in ch_msgs_df is not in all_users_df:
            if i_df['display_name'].shape[0] == 0 \
                    and df_messages.at[index, 'user'] == 'USLACKBOT':
                df_messages.at[index, 'name'] = 'USLACKBOT'
                df_messages.at[index, 'display_name'] = 'USLACKBOT'
                df_messages.at[index, 'is_bot'] = True
                df_messages.at[index, 'deactivated'] = False
            elif i_df['display_name'].shape[0] == 0 \
                    and df_messages.at[index, 'user'] != 'USLACKBOT':
                df_messages.at[index, 'name'] = "(user not found)"
                df_messages.at[index, 'display_name'] = "(user not found)"
                df_messages.at[index, 'is_bot'] = "(user not found)"
                df_messages.at[index, 'deactivated'] = "(user not found)"
            # --If users in ch_msgs_df is in all_users_df:
            else:
                df_messages.at[index, 'name'] = i_df['name'].values[0]
                df_messages.at[index, 'display_name'] = i_df['display_name'].values[0]
                df_messages.at[index, 'is_bot'] = i_df['is_bot'].values[0]
                df_messages.at[index, 'deactivated'] = i_df['deleted'].values[0]
            del i_df

    def ts_to_tz(self, df, original_column_name, new_column_name):
        """Transforms timestamps in a dataframe's column to dates on the
        US/Central timezone
        """
        df[original_column_name] = pd.to_numeric(
            df[original_column_name], errors='coerce'
            )
        tzs = []
        for i in range(len(df)):
            i_is_null = pd.Series(
                df.at[i, original_column_name]
                ).isnull().values[0]
            if i_is_null is True:
                i_date = self.missing_value
            else:
                i_date = pd.to_datetime(
                    df.at[i, original_column_name], unit='s'
                    ).tz_localize('UTC').tz_convert(self.timezone)
                try:
                    i_date = datetime.strftime(i_date, "%Y-%m-%d %H:%M:%S")
                except:
                    i_date = self.missing_value
            tzs.append(i_date)
        df[[original_column_name]].astype('datetime64[s]')
        df[original_column_name] = tzs
        df.rename(
            columns={original_column_name: new_column_name}, inplace=True
            )

    def extract_urls(self, df):
        """Extracts all the url links in df['text'] and stores them as a list
        in df['URL']
        """
        extractor = URLExtract()
        for i in range(len(df)):
            urls = []
            urls = extractor.find_urls(df.at[i, 'text'])
            if len(urls) > 0:
                urls_string = ' ;  '.join(urls)
                df.at[i, 'URL(s)'] = urls_string
            else:
                df.at[i, 'URL(s)'] = self.missing_value

    def user_id_to_name(self, df_messages, df_users):
        """Replaces the user_id in the format <@U12345678> to the user's
        display_name in df_messages['text'], which happens when the user is
        mentioned in an Slack message through the option @user_name.
        If there is no display_name, then 'user_id' is replaced with
        'profile_real_name'.
        All the bots in df_users have an 'id' and 'profile_real_name'
        (not necessarily 'name' and 'display_id'). Their profile_real_name are:
            Zoom, Google Drive, monday.com, monday.com notifications, GitHub,
            Google Calendar, Loom, Simple Poll, Figma, OneDrive and SharePoint,
            Calendly, Outlook Calendar, Rebecca Everlene Trust Company,
            Slack Team Emoji, New hire onboarding, Welcome,
            Clockify - Clocking in/out, Zapier, Update Your Slack Team Icon,
            Jira, Google Sheets, Time Off, Trailhead, Slack Team Emoji Copy,
            Guru, Guru, Google Calendar, Polly.
        'USLACKBOT' and 'B043CSZ0FL7' are the only bot messages if df_messages,
        but they are not in df_users!
        In the replacements, the "<<>>" are used for clarity on the text, since
        names can generally have more than one word and many names can be
        referenced one after the other, which can lead to confusion when
        reading.
        """
        for i in range(len(df_messages)):
            text = df_messages.at[i, 'text']
            matches = re.findall(r'<+@[A-Za-z0-9]+>', text)
            if len(matches) > 0:
                for match in matches:
                    user = match[2:-1]
                    if user in df_users['id'].values:
                        name = df_users[df_users['id'] == user]['display_name'].values[0]
                        is_bot = df_users[df_users['id'] == user]['is_bot'].values[0]
                        if is_bot is True:
                            name = df_users[df_users['id'] == user]['profile_real_name'].values[0] + ' (bot)'
                        elif name == self.missing_value:
                            name = df_users[df_users['id'] == user]['profile_real_name'].values[0]
                    else:
                        name = f"{user} (user not found)"
                    text = re.sub(f"<@{user}>", f"@{name}@", text)
                df_messages.at[i, 'text'] = text

    def parent_user_id_to_name(self, df_messages, df_users):
        # IP20241121   "parent_user_id"  substitution
        """Replaces the user_id in the format "UA5748HE" to the user's
        display_name in df_messages['parent_user_id']
        """
        for i in range(len(df_messages)):
            user = df_messages.at[i, 'parent_user_id']
            if user != self.missing_value:
                name = df_users[df_users['id'] == user]['display_name'].values
                if user in df_users['id'].values:
                    is_bot = df_users[df_users['id'] == user]['is_bot'].values
                    if is_bot is True:
                        name = df_users[df_users['id'] == user]['profile_real_name'].values + ' (bot)'
                    elif name == self.missing_value:
                        name = df_users[df_users['id'] == user]['profile_real_name'].values
                else:
                    name = user+' (user not found)'
                df_messages.at[i, 'parent_user_id'] = name
        df_messages.rename(
            columns={'parent_user_id': 'parent_user_name'}, inplace=True
            )

    def channel_id_to_name(self, df_messages, df_users):
        """Replaces <#channel_id|channel_name> to channel_name in
        df_messages['text'], which happens when the channel is mentioned in an
        Slack message through the option #channel_name
        """
        for i in range(len(df_messages)):
            text = df_messages.at[i, 'text']
            matches = re.findall(r'#+[A-Za-z0-9]+\|', text)
            if len(matches) > 0:
                for match in matches:
                    text = re.sub(match, "", text)
                    text = re.sub(r"<+\|", "<", text)
                df_messages.at[i, 'text'] = text

    def drop_extra_unparsed_rows(self, df):
        """ Drops the rows of the final-version of the df that were created
        from misparsed messages. For example, if the text has two projects, one
        identified as "Name:" and the other as "Project Name:"
        """
        rows_to_drop = []
        for i in range(len(df)):
            if df.at[i, 'projects_parsed'] == self.missing_value:
                rows_to_drop.append(i)
            if df.at[i, 'msg_date'] == self.missing_value \
                    and df.at[i, 'user'] == self.missing_value:
                rows_to_drop.append(i)
        df = df.drop(rows_to_drop)
        df = clean.reset_indices(df)
        return df

    def id_automatic_msgs(self, df):
        """ Get the indices of the messages sent automatically after an action
        has been taken on a Slack channel (rename channel, make the channel
        public and if an user joined or left the channel), and messages sent
        by a bot.
        """
        indices = []
        for i in range(len(df)):
            msg_id = df.at[i, 'msg_id']
            is_bot = df.at[i, 'is_bot']
            if 'channel_join' in msg_id or \
                    'channel_leave' in msg_id or \
                    'channel_name' in msg_id or \
                    'channel_canvas_updated' in msg_id or \
                    'channel_convert_to_public' in msg_id:
                indices.append(i)
            if is_bot is True or is_bot == 'True':  # review. Not being applied
                indices.append(i)
                print('bot message identified')
        return indices

    def get_automatic_msgs(self, df):
        """ Returns a dataframe with the automatic messages identified through
        id_automatic_msgs.
        """
        df_ = df.loc[self.id_automatic_msgs(df)]
        return clean.reset_indices(df_)

    def remove_automatic_msgs(self, df):
        """ Returns a dataframe without the the automatic messages identified
        through id_automatic_msgs.
        """
        indices2drop = self.id_automatic_msgs(df)
        df_ = df.copy()
        df_.drop(indices2drop, inplace=True)
        clean.reset_indices(df_)
        return df_

    def id_emojis_in_text(self, df):
        """ Returns a dataframe with an additional column "contained_emoji",
        indicating if the message had any emoji on the first place.
        No backup of the emojis is kept.
        """
        pattern = r'(:)([a-z0-9\_\-\+]+)(:)'
        for i in range(len(df)):
            text = df.at[i, 'text']
            match = re.search(pattern, text)
            if match is None:
                df.at[i, 'contained_emoji'] = False
            else:
                df.at[i, 'contained_emoji'] = True
                df.at[i, 'text'] = re.sub(pattern, "", text)
        return df

    def remove_emojis_in_text(self, df):
        """ Returns a dataframe where the emojis in "text" has been removed """
        pattern = r'(:)([a-z0-9\_\-\+]+)(:)'
        for i in range(len(df)):
            if df.at[i, 'contained_emoji'] is True:
                df.at[i, 'text'] = re.sub(pattern, "", df.at[i, 'text'])
        return df

    def id_short_msgs(self, df, n_char=15):
        """ Returns a list of indices of messages that contain less characters
        than the given n_char value.
        """
        indices = []
        for i in range(len(df)):
            text = df.at[i, 'text']
            if len(text) <= n_char:
                indices.append(i)
        return indices

    def get_short_msgs(self, df, n_char):
        """ Returns a dataframe including only the short messages identified
        with id_short_messages.
        """
        df_ = df.loc[self.id_short_msgs(df, n_char)]
        return clean.reset_indices(df_)

    def remove_short_msgs(self, df, n_char):
        """ Returns a dataframe without the short messages identifies with
        id_short_messages.
        """
        indices2drop = self.id_short_msgs(df, n_char)
        df_ = df.copy()
        df_.drop(indices2drop, inplace=True)
        return clean.reset_indices(df_)

    def apply_excel_adjustments(self, file_path, ws_name, settings_mod):
        """ Defines the sequence of changes to be done in the Excel file
        given the user's inputs in the module settings_mod.
        """
        xl = excel.ExcelFormat(file_path)
        ws = xl.get_sheet(ws_name)
        xl.set_cell_width(ws, settings_mod.column_widths)
        xl.set_allignment(ws, 'top')
        xl.format_first_row(ws, settings_mod.header_row)
        for cc in settings_mod.font_color_in_column:
            xl.set_font_color_in_column(ws, cc)
        for highlight in settings_mod.highlights:
            xl.format_highlight(ws, highlight)
        for column in settings_mod.text_type_cols:
            xl.format_text_cells(ws, column)
        xl.save_changes()

    def get_all_messages_df(self):
        """ Most generally, it iterates over all the Slack channels and
        extracts the messages of each channel and saves them in a formated
        Excel file. Only one channel is analyzed if the user specifies so in
        the input file.
        """
        if self.continue_analysis is False:
            print("Please review the input information")
        else:
            # --Iterate over channel's folders:
            dfs_list = []
            print(datetime.now().time(), 'Starting loop over channels', '\n')
            for i_channel in range(len(self.channels_names)):

                # --Define the name of the current channel and the source
                # --path containing its json files:
                curr_ch_name = self.channels_names[i_channel]
                print(
                    curr_ch_name, datetime.now().time(),
                    ' Set-up channel name and path to directory'
                    )

                # --Collect all the current_channel's messages in
                # ch_msgs_df::
                json_list = self.all_channels_jsonFiles_dates[i_channel]
                ch_msgs_df = self.get_ch_msgs_df(
                    self.slackexport_folder_path, curr_ch_name, json_list
                    )
                print(f'{curr_ch_name} Collected channel messages from the '
                      + 'json files')
                if len(ch_msgs_df) < 1:
                    print(
                        "for the folder ", curr_ch_name,
                        "messages_number= ", len(ch_msgs_df),
                        "there is no channel's folder", '\n'
                        )
                    continue

                # --Collect all the users in the current channel:
                channel_users_df = self.get_channel_users_df(ch_msgs_df,
                                                             self.all_users_df)
                print(f'{curr_ch_name} Collected users in current channel')

                # --Use channel_users_df to fill-in the user's information in
                # --ch_msgs_df:
                self.add_users_info_to_messages(ch_msgs_df, channel_users_df)
                print(f'{curr_ch_name} Included the users info on ch_msgs_df')

                # --Replace user and team identifiers with their
                # display_names whenever present in a message:
                self.user_id_to_name(ch_msgs_df, channel_users_df)
                self.channel_id_to_name(ch_msgs_df, channel_users_df)
                self.parent_user_id_to_name(ch_msgs_df, channel_users_df)
                print(f"{curr_ch_name} User's id replaced by their names")

                # --Extract hyperlinks from messages, if present
                # (extracted as a list; edit if needed):
                self.extract_urls(ch_msgs_df)
                print(f'{curr_ch_name} URLs extracted from messages')

                # --Change format of the time in seconds to a date in the
                # CST time-zone:
                self.ts_to_tz(ch_msgs_df, 'ts', 'msg_date')
                self.ts_to_tz(ch_msgs_df, 'json_mod_ts', 'json_mod_date')
                self.ts_to_tz(ch_msgs_df, 'ts_latest_reply', 'latest_reply_date')
                self.ts_to_tz(ch_msgs_df, 'ts_thread', 'thread_date')
                print(f'{curr_ch_name} Formated the dates and times')

                # --Identify if text has emojis:
                ch_msgs_df = self.id_emojis_in_text(ch_msgs_df)
                print(f'{curr_ch_name} Checked for emojis in messages')

                # --Parse for check-in messages:
                ch_msgs_df = checkins.parse_nrows(ch_msgs_df,
                                                  self.missing_value)
                ch_msgs_df = self.drop_extra_unparsed_rows(ch_msgs_df)
                print(f'{curr_ch_name} Parsed check-in messages')

                # --Build df with pruned messages:
                sel_msgs_df = self.remove_automatic_msgs(ch_msgs_df)
                sel_msgs_df = self.remove_emojis_in_text(sel_msgs_df)
                sel_msgs_df = self.remove_short_msgs(sel_msgs_df, n_char=15)
                print(f'{curr_ch_name} Built df with selected rows')

                # --Build df with descarded messages (after being pruned):
                auto_msgs = self.get_automatic_msgs(ch_msgs_df)
                short_msgs = self.get_short_msgs(ch_msgs_df, n_char=15)
                dis_msgs = pd.concat([auto_msgs, short_msgs],
                                     axis=0, ignore_index=False)
                dis_msgs.sort_values(by='msg_date',
                                     inplace=True, ignore_index=True)
                print(f'{curr_ch_name} Built df with filtered-out messages')

                # --Rearrange columns:
                column_names_order = self.settings_messages.columns_order
                ch_msgs_df = ch_msgs_df[column_names_order]
                sel_msgs_df = sel_msgs_df[column_names_order]
                dis_msgs = dis_msgs[column_names_order]
                print(f'{curr_ch_name} Rearranged columns')

                # --Sort rows by msg_date:
                ch_msgs_df.sort_values(by='msg_date',
                                       inplace=True, ignore_index=True)
                sel_msgs_df.sort_values(by='msg_date',
                                        inplace=True, ignore_index=True)
                dis_msgs.sort_values(by='msg_date',
                                     inplace=True, ignore_index=True)
                print(f'{curr_ch_name} Sorted rows by msg_date')

                # --Write ch_msgs_df to a .xlsx file:
                msgs_mindate = ch_msgs_df['msg_date'].min().split(" ")[0]
                msgs_maxdate = ch_msgs_df['msg_date'].max().split(" ")[0]
                curr_ch_name = curr_ch_name.replace(' ', '-')
                ch_msgs_filename = f"{curr_ch_name}_{msgs_mindate}_to_{msgs_maxdate}"
                ch_msgs_folder_path = f"{self.save_path}/{ch_msgs_filename}.xlsx"
                path = f"{ch_msgs_folder_path}"
                with pd.ExcelWriter(path, engine='openpyxl') as writer:
                    sel_msgs_df.to_excel(writer, index=False,
                                         sheet_name="Relevant messages")
                    dis_msgs.to_excel(writer, index=False,
                                      sheet_name="Filtered-out messages")
                    ch_msgs_df.to_excel(writer, index=False,
                                        sheet_name="All messages")

                # --Apply formatting of Excel worksheets:
                self.apply_excel_adjustments(path, 'All messages',
                                             self.settings_messages)
                self.apply_excel_adjustments(path, 'Relevant messages',
                                             self.settings_messages)
                self.apply_excel_adjustments(path, 'Filtered-out messages',
                                             self.settings_messages)
                print(f'{curr_ch_name} Wrote curated messages to Excel \n')

                dfs_list.append(ch_msgs_df)

        print(datetime.now().time(), 'Done')

        return ch_msgs_df
