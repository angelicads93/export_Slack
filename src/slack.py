#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: Angelica Goncalves.

Module to convert Slack JSON files into curated dataframes and Excel files.

Classes
-------
SlackChannelsAndUsers

SlackMessages

"""

# Import standard Python libraries:
import pandas as pd
from json import load
from datetime import datetime
from os.path import getmtime, exists
from urlextract import URLExtract
import re

# Import customed Python modules:
import excel
import clean
import checkins
import parser


class SlackChannelsAndUsers:
    """
    Class to handle information on the Slack channels and users.

    Relevant features are added and formatted into "channels" and "users"
    Pandas dataframes, which are later used to complement the information on
    the channel's messages.
    Requires objects of the customed classes "Parser" and  "InspectSource".

    ...

    Attributes
    ----------
    inputs : parser.Parser(txt_path)
        The parsed user's inputs from the file inputs.txt.
        Variables defined in inputs.txt are retrieved as inputs.get(var_name).

    settings : parser.Parser(txt_path)
        The parsed user's inputs from the file settings_messages.txt
        Variables defined in settings_messages.txt are retrieved as
        settings.get(var_name).

    Methods
    -------
    write_info_to_file(flag=None, df=None, filename=None, path=None)
        Write a given Pandas dataframe into an Excel file

    get_all_channels_info()
        Generate a curated Pandas dataframe from the "channels" JSON file.

    get_all_users_info()
        Generate a curated Pandas dataframe from the "users" JSON file.

    """

    def __init__(self, inputs, settings):

        # Retrieve users inputs from inputs.txt:
        self.inputs = inputs
        self.write_all_channels_info = self.inputs.get('write_all_channels_info')
        self.write_all_users_info = self.inputs.get('write_all_users_info')
        self.slackexport_folder_path = self.inputs.get('slackexport_folder_path')
        self.converted_directory = self.inputs.get('converted_directory')

        # Retrieve users inputs from settings.txt:
        self.settings = settings
        self.missing_value = self.settings.get('missing_value')
        self.timezone = self.settings.get('timezone')
        self.dest_name_ext = self.settings.get('dest_name_ext')
        self.channels_json_name = self.settings.get('channels_json_name')
        self.users_json_name = self.settings.get('users_json_name')
        self.channels_excel_name = self.settings.get('channels_excel_name')
        self.users_excel_name = self.settings.get('users_excel_name')

        # Create an instance of the class InspectSource:
        self.inspect_source = parser.InspectSource(self.inputs, self.settings)

    def write_info_to_file(self, flag, df, filename, path):
        """
        Write a given dataframe to an Excel file.

        Arguments
        ---------
        flag : bool
            Boolean specifying if proceeding with writing the Excel file.
        df : pandas.df()
            Pandas dataframe to be written into the Excel file.
        filename : str
            The name of the Excel file to be written.
        path : str
            The absolute path where to store the Excel file.

        """
        if flag is True:
            df.to_excel(f"{path}/{filename}{'.xlsx'}", index=False)
            print(datetime.now().time(), f"Wrote file {filename}.xlsx")

    def get_all_channels_info(self):
        """
        Export the channel's JSON file into a curated Pandas dataframe.

        The primary features of the dataframe are: id, name, created, creator,
            is_archived, is_general, members, pins, topic, purpose.
        The secondary features of 'pins' are: id, type, created, user, owner.
        The secondary features of 'topic' are: value, creator, last_set.

        """
        # Export channels.json to dataframe:
        chs_df = pd.read_json(
            f"{self.slackexport_folder_path}/{self.channels_json_name}")

        # Inspect each row and edit features of chs_df:
        for i in range(len(chs_df)):

            # Transfor "members" from a list to a string separated by ";":
            chs_df.at[i, 'members'] = ", ".join(chs_df.at[i, 'members'])

            # Adds df['purpose']:
            chs_df.at[i, 'purpose'] = chs_df.at[i, 'purpose']['value']

            # Add df['json_files'] with the channel's json_files. Use the
            # function "get_jsons_in_ch_dir" of the module
            # "inspect_source" module to verify the names of the files are in
            # the correct format (yyyy-mm-dd.json):
            ch_path = f"{self.slackexport_folder_path}/{chs_df.at[i, 'name']}"
            if exists(ch_path) is True:
                chs_df.at[i, 'json_files'] = str(self.inspect_source.chs_jsons)
            else:
                chs_df.at[i, 'json_files'] = self.missing_value

        # Keep relevant features:
        chs_df = chs_df[['id', 'name', 'created', 'creator', 'is_archived',
                         'is_general', 'members', 'purpose', 'json_files']]

        # Handle missing values or empty strings:
        for col in ['members', 'purpose']:
            clean.replace_empty_space(chs_df, col, self.missing_value)

        return chs_df

    def get_all_users_info(self):
        """
        Export the user's JSON file into a curated Pandas dataframe.

        The primary features of usrs_df are: id, team_id, name, deleted, color,
            real_name, tz, tz_label, tz_offset, profile, is_admin, is_owner,
            is_primary_owner, is_restricted,is_ultra_restricted, is_bot,
            is_app_user, updated, is_email_confirmed,
            who_can_share_contact_card, is_invited_user, is_workflow_bot,
            is_connector_bot.
        Among the secondary features of 'profile', there are: title, phone,
            skype, real_name, real_name_normalized, display_name,
            display_name_normalized, fields,status_text, status_emoji,
            status_emoji_display_info, status_expiration, avatar_hash,
            image_original, is_custom_image, email, huddle_state,
            huddle_state_expiration_ts, first_name, last_name, image_24,
            image_32, image_48, image_72, image_192, image_512, image_1024,
            status_text_canonical, team.

        """
        # Read users.json as a dataframe:
        usrs_df = pd.read_json(
            f"{self.slackexport_folder_path}/{self.users_json_name}")

        # Inspect each row and edit features of usrs_df:
        for i in range(len(usrs_df)):
            usrs_df.at[i, 'display_name'] = usrs_df.at[i, 'profile']['display_name']
            for col in ['title', 'real_name', 'status_text', 'status_emoji']:
                usrs_df.at[i, f"profile_{col}"] = usrs_df.at[i, 'profile'][col]

        # Keep relevant features:
        usrs_df = usrs_df[['id', 'team_id', 'name', 'deleted', 'display_name',
                           'is_bot', 'profile_title', 'profile_real_name',
                           'profile_status_text', 'profile_status_emoji']]

        # Handle missing values or empty strings:
        for col in ['display_name', 'name', 'team_id', 'id', 'profile_title',
                    'profile_real_name']:
            clean.replace_empty_space(usrs_df, col, self.missing_value)

        return usrs_df


class SlackMessages:
    """
    Class to handle information on the Slack messages.

    ...

    Attributes
    ----------
    inputs : parser.Parser(txt_path)
        The parsed user's inputs from the file inputs.txt.
        Variables defined in inputs.txt are retrieved as inputs.get(var_name).

    settings : parser.Parser(txt_path)
        The parsed user's inputs from the file settings_messages.txt
        Variables defined in settings_messages.txt are retrieved as
        settings.get(var_name).

    Methods
    -------
    msg_json_to_df(slack_json)
        Extract messages from a JSON file to a curated Pandas dataframe.

    get_ch_msgs_df(self, src_path, ch_name, json_list)
        Return a Pandas dataframe with all the messages of a given channel.

    get_ch_usrs_df(df_msgs, df_usrs)
        Return a data frame with the users info in the given Slack channel.

    add_usrs_info_to_msgs_df(df_msgs, df_usrs)
        Add information on the Slack users into the messages dataframe.

    ts_to_tz(df, orig_col_name, new_col_name)
        Transform timestamps to dates in a given column of a Pandas dataframe.

    extract_urls(df)
        Extract all the URLs found in the messages.

    usr_id_to_name(df_msgs, df_usrs)
        Replaces user_id with the user's display_name when mentioned in a msg.

    parent_id_to_name(df_msgs, df_usrs)
        Replace the parent_user_id to the parent_user's display_name.

    ch_id_to_name(df_msgs)
        Replaces the ch_id with the channel's name when mentioned in a msg.

    drop_extra_unparsed_rows(df_msgs)
        Drop extra rows of df_msgs created from misparsed messages.

    id_automatic_msgs(df_msgs)
        Return a list with the indices of the messages sent automatically.

    get_automatic_msgs(df_msgs)
        Return a dataframe with automatic messages.

    rm_automatic_msgs(df_msgs)
        Return a dataframe without automatic messages.

    id_emojis_in_text(df_msgs)
        Return a dataframe indicating if msg message had emoji(s).

    remove_emojis_in_text(df_msgs)
        Return a dataframe where the emojis in "text" have been removed.

    id_short_msgs(df_msgs, n_char)
        Return a list of indices of msgs with fewer characters than specified.

    get_short_msgs(df_msgs, n_char)
        Return a dataframe including only short messages.

    rm_short_msgs(df_msgs, n_char)
        Return a dataframe without short messages.

    apply_excel_adjustments(file_path, ws_name, settings):
        Formats the Excel tables as specified in the settings txt file.

    get_all_messages_df
        Write Excel files from curated dataframes containing all Slack msgs.

    """

    def __init__(self, inputs, settings):

        # Retrieve users inputs from inputs.txt:
        self.inputs = inputs
        self.slackexport_folder_path = self.inputs.get("slackexport_folder_path")

        # Retrieve users inputs from settings.txt:
        self.settings = settings
        self.missing_value = self.settings.get("missing_value")
        self.timezone = self.settings.get("timezone")

        # Create an instance of the class InspectSource:
        self.inspect_source = parser.InspectSource(
            self.inputs, self.settings)
        self.channels_names = self.inspect_source.present_chs
        self.all_channels_jsonFiles_dates = self.inspect_source.chs_jsons
        self.save_path = self.inspect_source.dest_path

        # Create an instance of the class SlackChannelsAndUsers:
        self.slack_channels_users = SlackChannelsAndUsers(
            self.inputs, self.settings)
        self.usrs_df = self.slack_channels_users.get_all_users_info()

    def msg_json_to_df(self, slack_json):
        """
        Extract messages from a JSON file to a curated Pandas dataframe.

        Arguments
        ---------
        slack_json : dict
            Dictionary containing the data from the JSON file.

        """
        # Initialize empty dataframe with given columns:
        msgs_df = pd.DataFrame(
            columns=["msg_id", "ts", "user", "type", "text", "reply_count",
                     "reply_users_count", "ts_latest_reply", "ts_thread",
                     "parent_user_id"])

        # Iterate through each msg and add relevant information to msgs_df:
        for msg in range(len(slack_json)):

            # Add the message id:
            if "client_msg_id" in slack_json[msg]:
                msgs_df.at[msg, "msg_id"] = slack_json[msg]["client_msg_id"]
            elif "subtype" in slack_json[msg]:
                msgs_df.at[msg, "msg_id"] = slack_json[msg]["subtype"]
            else:
                msgs_df.at[msg, "msg_id"] = self.missing_value

            # Add the message type:
            if "type" in slack_json[msg]:
                msgs_df.at[msg, "type"] = slack_json[msg]["type"]
            else:
                msgs_df.at[msg, "type"] = self.missing_value

            # Add the latest reply to the message:
            if "reply_count" in slack_json[msg]:
                msgs_df.at[msg, "ts_latest_reply"] = slack_json[msg]["latest_reply"]
            else:
                msgs_df.at[msg, "ts_latest_reply"] = self.missing_value

            # Add the id of the parent message if message is a reply:
            if "parent_user_id" in slack_json[msg]:
                msgs_df.at[msg, "ts_thread"] = slack_json[msg]["thread_ts"]
                msgs_df.at[msg, "type"] = "thread"
            else:
                msgs_df.at[msg, "ts_thread"] = self.missing_value

            # Add the message itself:
            msgs_df["text"] = msgs_df["text"].astype(str)

            # Add additional columns:
            for col in ["ts", "user",  "text", "reply_count",
                        "reply_users_count",  "parent_user_id"]:
                msgs_df.at[msg, col] = slack_json[msg].get(col,
                                                           self.missing_value)

        return msgs_df

    def get_ch_msgs_df(self, src_path, ch_name, json_list):
        """
        Return a Pandas dataframe with all the messages of a given channel.

        Arguments
        ---------
        src_path : str
            Absolute path of the source directory.

        ch_name : str
            Name of the Slack channel.

        json_list : list
            List with the name of the JSON files with the channel's messages.

        """
        # Initialize empty dataframe with given columns:
        ch_msgs_df = pd.DataFrame(
            columns=["msg_id", "ts", "user", "type", "text", "reply_count",
                     "reply_users_count", "ts_latest_reply", "ts_thread",
                     "parent_user_id"])

        # Iterate over JSONs inside the current channel's folder:
        for file_day in range(len(json_list)):
            filejson_path = f"{src_path}/{ch_name}/{json_list[file_day]}"

            with open(filejson_path, encoding="utf-8") as f:
                import_file_json = load(f)

            # Get the dataframe from the given JSON file:
            import_file_df = self.msg_json_to_df(import_file_json)

            # Add some id_cols:
            import_file_df["json_name"] = json_list[file_day]
            import_file_df["json_mod_ts"] = getmtime(filejson_path)

            # Concatenate the dataframe from the fiven JSON file with the
            # "full" dataframe ch_msgs_df:
            ch_msgs_df = pd.concat([ch_msgs_df, import_file_df],
                                   axis=0, ignore_index=True)

        # Add a column on the dataframe with the name of the channel:
        ch_msgs_df["channel_folder"] = ch_name

        return ch_msgs_df

    def get_ch_usrs_df(self, df_msgs, df_usrs):
        """
        Return a data frame with the users info in the given Slack channel.

        Arguments
        ---------
        df_msgs : Pandas dataframe

        df_usrs : Pandas dataframe

        """
        # Initialize channel_df_usrs as a copy of df_usrs:
        df = df_usrs.copy()
        # Find the unique set of users in channel:
        channel_users_list = df_msgs["user"].unique()
        # Collect the indices of the users that are NOT in the channel:
        indices_to_drop = [i
                           for i in range(len(df_usrs))
                           if df_usrs.at[i, "id"] not in channel_users_list]
        # Drop the rows on indices_to_drop:
        df.drop(df.index[indices_to_drop], inplace=True)
        return df

    def add_usrs_info_to_msgs_df(self, df_msgs, df_usrs):
        """
        Add information of the Slack users into the messages dataframe.

        Uses the user's id in the format U1234567789 from the df_msgs to
        find the name, display name and if the user is a bot from df_usrs.
        The "name", "display_name" and "is_bot" are then added as columns to
        df_msgs

        Arguments
        ---------
        df_msgs : Pandas dataframe

        df_usrs : Pandas dataframe

        """
        for i in df_msgs.index.values:
            i_df = df_usrs[df_usrs["id"] == df_msgs.at[i, "user"]]
            # If users in ch_msgs_df is not in usrs_df:
            if i_df["display_name"].shape[0] == 0 \
                    and df_msgs.at[i, "user"] == "USLACKBOT":
                df_msgs.at[i, "name"] = "USLACKBOT"
                df_msgs.at[i, "display_name"] = "USLACKBOT"
                df_msgs.at[i, "is_bot"] = True
                df_msgs.at[i, "deactivated"] = False
            elif i_df["display_name"].shape[0] == 0 \
                    and df_msgs.at[i, "user"] != "USLACKBOT":
                df_msgs.at[i, "name"] = "(user not found)"
                df_msgs.at[i, "display_name"] = "(user not found)"
                df_msgs.at[i, "is_bot"] = "(user not found)"
                df_msgs.at[i, "deactivated"] = "(user not found)"
            # If users in ch_msgs_df is in usrs_df:
            else:
                df_msgs.at[i, "name"] = i_df["name"].values[0]
                df_msgs.at[i, "display_name"] = i_df["display_name"].values[0]
                df_msgs.at[i, "is_bot"] = i_df["is_bot"].values[0]
                df_msgs.at[i, "deactivated"] = i_df["deleted"].values[0]
            del i_df

    def ts_to_tz(self, df, orig_col_name, new_col_name):
        """
        Transform timestamps to dates in a given column of a Pandas dataframe.

        Changes the name of the column.

        Arguments
        ---------
        df : Pandas dataframe

        orig_col_name : str
            Name of the column with the timestamps

        new_col_name : str
            Name to use when renaming the column.

        """
        # Retrieve the column from the Pandas dataframe:
        df[orig_col_name] = pd.to_numeric(df[orig_col_name], errors="coerce")

        # Store the converted dates into a list:
        tzs = []
        for i in range(len(df)):
            i_is_null = pd.Series(df.at[i, orig_col_name]).isnull().values[0]
            if i_is_null is True:
                i_date = self.missing_value
            else:
                i_date = pd.to_datetime(
                    df.at[i, orig_col_name], unit="s"
                    ).tz_localize("UTC").tz_convert(self.timezone)
                try:
                    i_date = datetime.strftime(i_date, "%Y-%m-%d %H:%M:%S")
                except:
                    i_date = self.missing_value
            tzs.append(i_date)

        # First, change the type of the dataframe column:
        df[[orig_col_name]].astype("datetime64[s]")

        # Then fill the column with the new dates in tzs:
        df[orig_col_name] = tzs

        # Rename the column that it is now a "date":
        df.rename(columns={orig_col_name: new_col_name}, inplace=True)

    def extract_urls(self, df):
        """
        Extract all the URLs found in the messages.

        Arguments
        ---------
        df : Pandas dataframe

        """
        # Initialize constructor of the class URLExtract():
        extractor = URLExtract()

        # Iterate over all the messages in df:
        for i in range(len(df)):
            # Extract any URLs as list
            urls = extractor.find_urls(df.at[i, "text"])
            if len(urls) > 0:
                # Rewrite the list as a string separated by commas:
                urls_string = ";  ".join(urls)
                df.at[i, "URL(s)"] = urls_string
            else:
                df.at[i, "URL(s)"] = self.missing_value

    def usr_id_to_name(self, df_msgs, df_usrs):
        """
        Replace user_id with the user's display_name when mentioned in a msg.

        If there is no display_name, then "user_id" is replaced with
        "profile_real_name".
        All the bots in df_usrs have an "id" and "profile_real_name"
        (not necessarily "name" and "display_id"). Their profile_real_name are:
            Zoom, Google Drive, monday.com, monday.com notifications, GitHub,
            Google Calendar, Loom, Simple Poll, Figma, OneDrive and SharePoint,
            Calendly, Outlook Calendar, Rebecca Everlene Trust Company,
            Slack Team Emoji, New hire onboarding, Welcome,
            Clockify - Clocking in/out, Zapier, Update Your Slack Team Icon,
            Jira, Google Sheets, Time Off, Trailhead, Slack Team Emoji Copy,
            Guru, Guru, Google Calendar, Polly.
        "USLACKBOT" and "B043CSZ0FL7" are the only bot messages if df_msgs,
        but they are not in df_usrs!
        In the replacements, the "@" is used to wrap the display_name for
        clarity on the text, since names can generally have more than one word
        and many names can be referenced one after the other, which can lead
        to confusion when reading.

        Arguments
        ---------
        df_msgs : Pandas dataframe

        df_usrs : Pandas dataframe

        """
        for i in range(len(df_msgs)):
            text = df_msgs.at[i, "text"]
            matches = re.findall(r"<+@[A-Za-z0-9]+>", text)
            if len(matches) > 0:
                for match in matches:
                    user = match[2:-1]
                    if user in df_usrs["id"].values:
                        name = df_usrs[df_usrs["id"] == user]["display_name"].values[0]
                        is_bot = df_usrs[df_usrs["id"] == user]["is_bot"].values[0]
                        if is_bot is True:
                            name = df_usrs[df_usrs["id"] == user]["profile_real_name"].values[0] + " (bot)"
                        elif name == self.missing_value:
                            name = df_usrs[df_usrs["id"] == user]["profile_real_name"].values[0]
                    else:
                        name = f"{user} (user not found)"
                    text = re.sub(f"<@{user}>", f"@{name}@", text)
                df_msgs.at[i, "text"] = text

    def parent_id_to_name(self, df_msgs, df_usrs):
        """
        Replace the parent_user_id to the parent_user's display_name.

        Arguments
        ---------
        df_msgs : Pandas dataframe

        df_usrs : Pandas dataframe

        """
        for i in range(len(df_msgs)):
            user = df_msgs.at[i, "parent_user_id"]
            if user != self.missing_value:
                name = df_usrs[df_usrs["id"] == user]["display_name"].values
                if user in df_usrs["id"].values:
                    is_bot = df_usrs[df_usrs["id"] == user]["is_bot"].values
                    if is_bot is True:
                        name = df_usrs[df_usrs["id"] == user]["profile_real_name"].values + " (bot)"
                    elif name == self.missing_value:
                        name = df_usrs[df_usrs["id"] == user]["profile_real_name"].values
                else:
                    name = user+" (user not found)"
                df_msgs.at[i, "parent_user_id"] = name
        df_msgs.rename(
            columns={"parent_user_id": "parent_user_name"}, inplace=True
            )

    def ch_id_to_name(self, df_msgs):
        """
        Replace the channel_id with the channel's name when mentioned in a msg.

        Arguments
        ---------
        df_msgs : Pandas dataframe

        """
        for i in range(len(df_msgs)):
            text = df_msgs.at[i, "text"]
            matches = re.findall(r"#+[A-Za-z0-9]+\|", text)
            if len(matches) > 0:
                for match in matches:
                    # Replace <#channel_id|channel_name> to channel_name
                    text = re.sub(match, "", text)
                    text = re.sub(r"<+\|", "<", text)
                df_msgs.at[i, "text"] = text

    def drop_extra_unparsed_rows(self, df_msgs):
        """
        Drop extra rows of df_msgs created from misparsed messages.

        This is the case, for example, if the text has two projects, one
        identified as "Name:" and the other as "Project Name:"

        Arguments
        ---------
        df_msgs : Pandas dataframe

        """
        rows_to_drop = []
        for i in range(len(df_msgs)):
            if df_msgs.at[i, "projects_parsed"] == self.missing_value:
                rows_to_drop.append(i)
            if df_msgs.at[i, "msg_date"] == self.missing_value \
                    and df_msgs.at[i, "user"] == self.missing_value:
                rows_to_drop.append(i)
        df_msgs = df_msgs.drop(rows_to_drop)
        df_msgs = clean.reset_indices(df_msgs)
        return df_msgs

    def id_automatic_msgs(self, df_msgs):
        """
        Return a list with the indices of the messages sent automatically.

        Actions include renaming a channel, making the channel public, if
        a user joined or left the channel, and messages sent by a bot.

        Arguments
        ---------
        df_msgs : Pandas dataframe

        """
        indices = []
        for i in range(len(df_msgs)):
            msg_id = df_msgs.at[i, "msg_id"]
            is_bot = df_msgs.at[i, "is_bot"]
            if "channel_join" in msg_id or \
                    "channel_leave" in msg_id or \
                    "channel_name" in msg_id or \
                    "channel_canvas_updated" in msg_id or \
                    "channel_convert_to_public" in msg_id:
                indices.append(i)
            if is_bot is True or is_bot == "True":  # review. Not being applied
                indices.append(i)
                print("bot message identified")
        return indices

    def get_automatic_msgs(self, df_msgs):
        """
        Return a dataframe with automatic messages.

        Arguments
        ---------
        df_msgs : Pandas dataframe

        """
        # Retrieve rows from indices:
        df = df_msgs.loc[self.id_automatic_msgs(df_msgs)]
        # Reset the indices:
        clean.reset_indices(df)
        return df

    def rm_automatic_msgs(self, df_msgs):
        """
        Return a dataframe without automatic messages.

        Arguments
        ---------
        df_msgs : Pandas dataframe

        """
        # Create a copy of the input dataframe and drop the relevant indices:
        df = df_msgs.copy()
        df.drop(self.id_automatic_msgs(df_msgs), inplace=True)
        clean.reset_indices(df)
        return df

    def id_emojis_in_text(self, df_msgs):
        """
        Return a dataframe indicating if msg message had emoji(s).

        No backup of the emojis is kept.

        Arguments
        ---------
        df_msgs : Pandas dataframe

        """
        pattern = r"(:)([a-z0-9\_\-\+]+)(:)"
        for i in range(len(df_msgs)):
            text = df_msgs.at[i, "text"]
            match = re.search(pattern, text)
            if match is None:
                df_msgs.at[i, "contained_emoji"] = False
            else:
                df_msgs.at[i, "contained_emoji"] = True
                df_msgs.at[i, "text"] = re.sub(pattern, "", text)
        return df_msgs

    def remove_emojis_in_text(self, df_msgs):
        """
        Return a dataframe where the emojis in "text" have been removed.

        Arguments
        ---------
        df_msgs : Pandas dataframe

        """
        pattern = r"(:)([a-z0-9\_\-\+]+)(:)"
        for i in range(len(df_msgs)):
            if df_msgs.at[i, "contained_emoji"] is True:
                df_msgs.at[i, "text"] = re.sub(pattern, "",
                                               df_msgs.at[i, "text"])
        return df_msgs

    def id_short_msgs(self, df_msgs, n_char):
        """
        Return a list of indices of msgs with fewer characters than specified.

        Arguments
        ---------
        df_msgs : Pandas dataframe

        n_char : int
            Minimum number of characters that a message should have to not be
            dropped.

        """
        indices = []
        for i in range(len(df_msgs)):
            text = df_msgs.at[i, "text"]
            if len(text) <= n_char:
                indices.append(i)
        return indices

    def get_short_msgs(self, df_msgs, n_char):
        """
        Return a dataframe including only short messages.

        Arguments
        ---------
        df_msgs : Pandas dataframe

        n_char : int
            Minimum number of characters that a message should have to not be
            dropped.

        """
        df = df_msgs.loc[self.id_short_msgs(df_msgs, n_char)]
        clean.reset_indices(df)
        return df

    def rm_short_msgs(self, df_msgs, n_char=15):
        """
        Return a dataframe without short messages.

        Arguments
        ---------
        df_msgs : Pandas dataframe

        n_char : int (optional. Default value is 15)
            Minimum number of characters that a message should have to not be
            dropped.
        """
        indices2drop = self.id_short_msgs(df_msgs, n_char)
        df = df_msgs.copy()
        df.drop(indices2drop, inplace=True)
        clean.reset_indices(df)
        return df

    def apply_excel_adjustments(self, file_path, ws_name, settings):
        """
        Format the Excel tables as specified in the settings txt file.

        Arguments
        ---------
        file_path : str
            Absolute path to the Excel file

        ws_name : str
            Name of the Excel Sheet

        settings : parser.Parser(txt_path)
            Parsed user inputs to the settings txt file.

        """
        xl = excel.ExcelFormat(file_path, settings)
        ws = xl.get_sheet(ws_name)
        xl.set_cell_width(ws, settings.get("column_widths"))
        xl.set_allignment(ws, "top")
        xl.format_first_row(ws, settings.get("header_row"))
        for cc in settings.get("font_color_in_column"):
            xl.set_font_color_in_column(ws, cc)
        for highlight in settings.get("highlights"):
            xl.format_highlight(ws, highlight)
        for column in settings.get("text_type_cols"):
            xl.format_text_cells(ws, column)
        xl.save_changes()

    def get_all_messages_df(self):
        """Write Excel files from curated dataframes containing all Slack msgs."""
        # Iterate over channel's folders:
        dfs_list = []
        print(datetime.now().time(), "Starting loop over channels", "\n")
        for i_channel in range(len(self.channels_names)):

            # Define the name of the current channel and the source path
            # containing its json files:
            curr_ch_name = self.channels_names[i_channel]
            print(curr_ch_name, datetime.now().time())
            print(f"{curr_ch_name} Set-up channel name and path to directory")

            # Collect all the current_channel's messages in ch_msgs_df:
            json_list = self.all_channels_jsonFiles_dates[i_channel]
            ch_msgs_df = self.get_ch_msgs_df(
                self.slackexport_folder_path, curr_ch_name, json_list
                )
            print(f"{curr_ch_name} Collected channel msgs from the json files")
            if len(ch_msgs_df) < 1:
                print(
                    "for the folder ", curr_ch_name,
                    "messages_number= ", len(ch_msgs_df),
                    "there is no channel's folder", "\n"
                    )
                continue

            # Collect all the users in the current channel:
            channel_users_df = self.get_ch_usrs_df(ch_msgs_df, self.usrs_df)
            print(f"{curr_ch_name} Collected users in current channel")

            # Use channel_users_df to fill in the user's information in
            # ch_msgs_df:
            self.add_usrs_info_to_msgs_df(ch_msgs_df, channel_users_df)
            print(f"{curr_ch_name} Included the users info on ch_msgs_df")

            # Replace user and team identifiers with their
            # display_names whenever present in a message:
            self.usr_id_to_name(ch_msgs_df, channel_users_df)
            self.ch_id_to_name(ch_msgs_df)
            self.parent_id_to_name(ch_msgs_df, channel_users_df)
            print(f"{curr_ch_name} User's id replaced by their names")

            # Extract hyperlinks from messages, if present
            # (extracted as a list; edit if needed):
            self.extract_urls(ch_msgs_df)
            print(f"{curr_ch_name} URLs extracted from messages")

            # Change format of the time in seconds to a date in the
            # CST time-zone:
            self.ts_to_tz(ch_msgs_df, "ts", "msg_date")
            self.ts_to_tz(ch_msgs_df, "json_mod_ts", "json_mod_date")
            self.ts_to_tz(ch_msgs_df, "ts_latest_reply", "latest_reply_date")
            self.ts_to_tz(ch_msgs_df, "ts_thread", "thread_date")
            print(f"{curr_ch_name} Formated the dates and times")

            # Identify if text has emojis:
            ch_msgs_df = self.id_emojis_in_text(ch_msgs_df)
            print(f"{curr_ch_name} Checked for emojis in messages")

            # Parse for check-in messages:
            ch_msgs_df = checkins.parse_reports(ch_msgs_df, self.settings)

            ch_msgs_df = self.drop_extra_unparsed_rows(ch_msgs_df)
            print(f"{curr_ch_name} Parsed check-in messages")

            # Build df with pruned messages:
            sel_msgs_df = self.rm_automatic_msgs(ch_msgs_df)
            sel_msgs_df = self.remove_emojis_in_text(sel_msgs_df)
            sel_msgs_df = self.rm_short_msgs(sel_msgs_df, n_char=15)
            print(f"{curr_ch_name} Built df with selected rows")

            # Build df with descarded messages (after being pruned):
            auto_msgs = self.get_automatic_msgs(ch_msgs_df)
            short_msgs = self.get_short_msgs(ch_msgs_df, n_char=15)
            dis_msgs = pd.concat([auto_msgs, short_msgs],
                                 axis=0, ignore_index=False)
            dis_msgs.sort_values(by="msg_date",
                                 inplace=True, ignore_index=True)
            print(f"{curr_ch_name} Built df with filtered-out messages")

            # Rearrange columns:
            column_names_order = self.settings.get("columns_order")
            ch_msgs_df = ch_msgs_df[column_names_order]
            sel_msgs_df = sel_msgs_df[column_names_order]
            dis_msgs = dis_msgs[column_names_order]
            print(f"{curr_ch_name} Rearranged columns")

            # Sort rows by msg_date:
            ch_msgs_df.sort_values(by="msg_date",
                                   inplace=True, ignore_index=True)
            sel_msgs_df.sort_values(by="msg_date",
                                    inplace=True, ignore_index=True)
            dis_msgs.sort_values(by="msg_date",
                                 inplace=True, ignore_index=True)
            print(f"{curr_ch_name} Sorted rows by msg_date")

            # Write ch_msgs_df to a .xlsx file:
            msgs_mindate = ch_msgs_df["msg_date"].min().split(" ")[0]
            msgs_maxdate = ch_msgs_df["msg_date"].max().split(" ")[0]
            curr_ch_name = curr_ch_name.replace(" ", "-")
            ch_msgs_filename = f"{curr_ch_name}_{msgs_mindate}_to_{msgs_maxdate}"
            ch_msgs_folder_path = f"{self.save_path}/{ch_msgs_filename}.xlsx"
            path = f"{ch_msgs_folder_path}"
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                sel_msgs_df.to_excel(writer, index=False,
                                     sheet_name="Relevant messages")
                dis_msgs.to_excel(writer, index=False,
                                  sheet_name="Filtered-out messages")
                ch_msgs_df.to_excel(writer, index=False,
                                    sheet_name="All messages")

            # Apply formatting of Excel worksheets:
            self.apply_excel_adjustments(path, "All messages",
                                         self.settings)
            self.apply_excel_adjustments(path, "Relevant messages",
                                         self.settings)
            self.apply_excel_adjustments(path, "Filtered-out messages",
                                         self.settings)
            print(f"{curr_ch_name} Wrote curated messages to Excel \n")

            dfs_list.append(ch_msgs_df)

        print(datetime.now().time(), "Done")

        return ch_msgs_df
