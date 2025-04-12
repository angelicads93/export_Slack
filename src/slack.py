#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: Angelica Goncalves.

Module to extract and format the exported information of a Slack workspace.

All the custom functions that edits informationfrom the Slack export
are collected here.

Classes
-------
Slack

Functions
---------
write_info_to_file(flag, df, filename, path)

"""

# Import standard Python libraries:
import re
import os
from datetime import datetime
from json import load
import pandas as pd
from urlextract import URLExtract

# Import customed Python modules:
import excel
import clean
import checkins


# NOTE: There may be a better way of doing it, but the class Slack was defined
# to carry on the variables defined by the user in the txt files without
# needing to add them as arguments to the various functions explicitly. They
# are instead treated as attributes of the class.
class Slack:
    """
    Class to handle information on the Slack channels and users.

    Relevant features are added and formatted into "channels" and "users"
    Pandas dataframes, which are later used to complement the information on
    the channel's messages.

    ...

    Attributes
    ----------
    settings : parser.Parser(txt_path)
        The parsed user's inputs from the file settings_messages.txt
        Variables defined in settings_messages.txt are retrieved as
        settings.get(var_name).

    Methods
    -------
    get_jsons_in_ch(ch_path)
        Return a list with the names of the JSON files for a Slack channel.

    get_all_channels_info(source_path, chs_json_path)
        Export the channel's JSON file into a curated Pandas dataframe.

    get_all_users_info(usrs_json_path)
        Export the JSON file with all the Slack users into a Pandas dataframe.

    msg_json_to_df(slack_json)
        Export the messages from one JSON file into a Pandas dataframe.

    get_ch_msgs_df(src_path, ch_name)
        Export all the messages in a Slack channel into a Pandas dataframe.

    get_ch_usrs_df(df_msgs, df_usrs)
        Export the information of the users found in a Slack channel.

    add_usrs_info_to_msgs_df(df_msgs, df_usrs)
        Add information of the Slack users into the messages dataframe.

    ts_to_tz(df, orig_col_name, new_col_name)
        Rewrites timestamps into dates in a given column of a Pandas dataframe.

    extract_urls(df)
        Extract all the URLs found in the column "text" of a Pandas dataframe.

    usr_id_to_name(df_msgs, df_usrs)
        Replace user_id with the user's display_name in df_msgs["text"].

    parent_id_to_name(df_msgs, df_usrs)
        Replace parent_user_id with its display_name in the dataframe.

    ch_id_to_name(df_msgs)
        Replace the channel_id with the channel's name in df_msgs["text"].

    drop_extra_unparsed_rows(df_msgs)
        Drop empty rows in df_msgs created from misparsed messages.

    id_automatic_msgs(df_msgs)
        Return a list with the indices of the messages sent automatically.

    get_automatic_msgs(df_msgs)
        Return a dataframe with "automatic" messages.

    rm_automatic_msgs(df_msgs)
        Return a dataframe without automatic messages.

    id_emojis_in_text(df_msgs)
        Return a dataframe indicating if msg message had emoji(s).

    remove_emojis_in_text(df_msgs)
        Return a dataframe with no emojis in df_msgs["text"].

    id_short_msgs(df_msgs, n_char)
        Identify msgs with fewer characters than n_char.

    get_short_msgs(df_msgs, n_char)
        Return a dataframe including only short messages.

    rm_short_msgs(df_msgs, n_char=15)
        Return a dataframe without short messages.

    add_channel_info(channel_path, channel_df)
        Add "channel", "export_dates" and "parsed_reports_in_channel" to df.

    add_info_of_users_reports(df_all)
        Get the latest report date and number of messages in the given channel.

    format_parsed_reports(df_all)
        Select the parsed weekly reports and sort the df by channel.

    format_unparsed_reports(df_all, wr_channel_name)
        Select the unparsed weekly reports and sort the df by channel.

    format_msgs_with_urls(df, settings)
        Select the messages that contain URL(s) in df["text"].

    filter_urls(url_list, settings)
        Select a message's URLs as specified in the settings txt file.

    select_desired_urls(df, settings)
        Filter-out unwanted URL(s) as specified in the settings txt file.

    """

    def __init__(self, settings):
        # Retrieve the users' variables from the settings txt file:
        self.settings = settings
        self.missing_value = self.settings.get("missing_value")
        self.timezone = self.settings.get("timezone")

    def get_jsons_in_ch(self, ch_path):
        """
        Return a list with the names of the JSON files for a Slack channel.

        The name of the files must be of the form YYYY-MM-DD.json to be valid.

        Arguments
        ---------
        ch_path : str
            Path to the given Slack channel in the source directory.

        Returns
        -------
        list_names_dates : list

        """
        ch_files = os.listdir(ch_path)
        list_names_dates = []
        for i in range(len(ch_files)):
            match = re.match(
                r"(\d{4})(-)(\d{2})(-)(\d{2})(.)(json)", ch_files[i]
                )
            if match is not None:
                list_names_dates.append(ch_files[i])
        return list_names_dates

    def get_all_channels_info(self, source_path, chs_json_path):
        """
        Export the channel's JSON file into a curated Pandas dataframe.

        Arguments
        ---------
        source_path : str
            Path to the source directory.
        chs_json_path : str
            Path to the JSON file "channels.json" in the source directory.

        Returns
        -------
        Pandas dataframe with columns: "id", "name", "created", "creator",
        "is_archived", "is_general", "members", "purpose", "json_files".

        """
        # Export channels.json to dataframe:
        chs_df = pd.read_json(chs_json_path)
        # Note:
        # The primary features of the dataframe are: id, name, created,
        # creator, is_archived, is_general, members, pins, topic, purpose.
        # The secondary features of "pins" are: id, type, created, user, owner.
        # The secondary features of "topic" are: value, creator, last_set.

        # Inspect each row and edit features of chs_df:
        for i in range(len(chs_df)):

            # Transfor "members" from a list to a string separated by ";":
            chs_df.at[i, "members"] = ", ".join(chs_df.at[i, "members"])

            # Adds df["purpose"]:
            chs_df.at[i, "purpose"] = chs_df.at[i, "purpose"]["value"]

            # Add df["json_files"] with the channel's json_files. Use the
            # function "get_jsons_in_ch_dir" to verify the names of the files
            # are in the correct format (yyyy-mm-dd.json):
            ch_path = f"{source_path}/{chs_df.at[i, 'name']}"
            if os.path.exists(ch_path) is True:
                chs_df.at[i, "json_files"] = str(self.get_jsons_in_ch(ch_path))
            else:
                chs_df.at[i, "json_files"] = self.missing_value

        # Keep relevant features:
        chs_df = chs_df[["id", "name", "created", "creator", "is_archived",
                         "is_general", "members", "purpose", "json_files"]]

        # Handle missing values or empty strings:
        for col in ["members", "purpose"]:
            clean.replace_empty_space(chs_df, col, self.missing_value)

        return chs_df

    def get_all_users_info(self, usrs_json_path):
        """
        Export the JSON file with all the Slack users into a Pandas dataframe.

        Arguments
        ---------
        usrs_json_path : str
            Path to the JSON file "users.json" in the source directory.

        Returns
        -------
        Pandas dataframe with columns: "id", "team_id", "name", "deleted",
        "display_name", "is_bot", "profile_title", "profile_real_name",
        "profile_status_text", "profile_status_emoji".

        """
        # Read users.json as a dataframe:
        usrs_df = pd.read_json(usrs_json_path)
        # Note:
        # The primary features of usrs_df are: id, team_id, name, deleted,
        # color, real_name, tz, tz_label, tz_offset, profile, is_admin,
        # is_owner, is_primary_owner, is_restricted,is_ultra_restricted,
        # is_bot, is_app_user, updated, is_email_confirmed,
        # who_can_share_contact_card, is_invited_user, is_workflow_bot,
        # is_connector_bot.
        # The secondary features of "profile" contain: title, phone, skype,
        # real_name, real_name_normalized, display_name,
        # display_name_normalized, fields,status_text, status_emoji,
        # status_emoji_display_info, status_expiration, avatar_hash,
        # image_original, is_custom_image, email, huddle_state,
        # huddle_state_expiration_ts, first_name, last_name, image_24,
        # image_32, image_48, image_72, image_192, image_512, image_1024,
        # status_text_canonical, team.

        # Inspect each row and edit features of usrs_df:
        for i in range(len(usrs_df)):
            usrs_df.at[i, "display_name"] = usrs_df.at[i, "profile"]["display_name"]
            for col in ["title", "real_name", "status_text", "status_emoji"]:
                usrs_df.at[i, f"profile_{col}"] = usrs_df.at[i, "profile"][col]

        # Keep relevant features:
        usrs_df = usrs_df[["id", "team_id", "name", "deleted", "display_name",
                           "is_bot", "profile_title", "profile_real_name",
                           "profile_status_text", "profile_status_emoji"]]

        # Handle missing values or empty strings:
        for col in ["display_name", "name", "team_id", "id", "profile_title",
                    "profile_real_name"]:
            clean.replace_empty_space(usrs_df, col, self.missing_value)

        return usrs_df

    def msg_json_to_df(self, slack_json):
        """
        Export the messages from one JSON file into a Pandas dataframe.

        Arguments
        ---------
        slack_json : dict
            Dictionary containing the data from the JSON file (Product of
            using Python's load() function).

        Returns
        -------
        Pandas dataframe with columns: "msg_id", "ts", "user", "type", "text",
        "reply_count", "reply_users_count", "ts_latest_reply", "ts_thread",
        "parent_user_id".

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

    def get_ch_msgs_df(self, src_path, ch_name):
        """
        Export all the messages in a Slack channel into a Pandas dataframe.

        Arguments
        ---------
        src_path : str
            Absolute path to the source directory.
        ch_name : str
            Name of the Slack channel.

        Returns
        -------
        Pandas dataframe with columns: "msg_id", "ts", "user",
        "type", "text", "reply_count", "reply_users_count", "ts_latest_reply",
        "ts_thread", "parent_user_id".

        """
        # Initialize empty dataframe with given columns:
        ch_msgs_df = pd.DataFrame(
            columns=["msg_id", "ts", "user", "type", "text", "reply_count",
                     "reply_users_count", "ts_latest_reply", "ts_thread",
                     "parent_user_id"])

        json_list = self.get_jsons_in_ch(f"{src_path}/{ch_name}")

        # Iterate over JSONs inside the current channel's folder:
        for file_day in range(len(json_list)):
            filejson_path = f"{src_path}/{ch_name}/{json_list[file_day]}"

            with open(filejson_path, encoding="utf-8") as f:
                import_file_json = load(f)

            # Get the dataframe from the given JSON file:
            import_file_df = self.msg_json_to_df(import_file_json)

            # Add some id_cols:
            import_file_df["json_name"] = json_list[file_day]
            import_file_df["json_mod_ts"] = os.path.getmtime(filejson_path)

            # Concatenate the dataframe from the fiven JSON file with the
            # "full" dataframe ch_msgs_df:
            ch_msgs_df = pd.concat([ch_msgs_df, import_file_df],
                                   axis=0, ignore_index=True)

        # Add a column on the dataframe with the name of the channel:
        ch_msgs_df["channel_folder"] = ch_name

        return ch_msgs_df

    def get_ch_usrs_df(self, df_msgs, df_usrs):
        """
        Export the information of the users found in a Slack channel.

        Arguments
        ---------
        df_msgs : Pandas dataframe
            Dataframe containing the messages of a Slack channel.
        df_usrs : Pandas dataframe
            Dataframe built from the file users.json, containing the information
            from all the users in the Slack workspace.

        Returns
        -------
        Pandas dataframe with columns: "id", "team_id", "name", "deleted",
        "display_name", "is_bot", "profile_title", "profile_real_name",
        "profile_status_text", "profile_status_emoji"

        """
        # Initialize channel_df_usrs as a copy of df_usrs:
        df = df_usrs.copy()
        # Find the unique set of users in the channel:
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

        1. Take the user's id (in the format U1234567789) from the df_msgs.
        2. Find the user's name, display name, and bot status from df_usrs.
        3. Add (in-place) "name", "display_name" and "is_bot" to df_msgs.

        Arguments
        ---------
        df_msgs : Pandas dataframe
            Dataframe containing the messages of a Slack channel.
        df_usrs : Pandas dataframe
            Dataframe built from the file users.json, containing the information
            from all the users in the Slack workspace.

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
        Rewrites timestamps into dates in a given column of a Pandas dataframe.

        Arguments
        ---------
        df : Pandas dataframe
            Dataframe containing the messages of a Slack channel.
        orig_col_name : str
            Name of the column with the timestamps.
        new_col_name : str
            Name to use when renaming the column with the dates.

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

        # Then, fill the column with the new dates in tzs:
        df[orig_col_name] = tzs

        # Rename the column that it is now a "date":
        df.rename(columns={orig_col_name: new_col_name}, inplace=True)

    def extract_urls(self, df):
        """
        Extract all the URLs found in the column "text" of a Pandas dataframe.

        The URL(s) are stored in a new column called "URL(s)".

        Arguments
        ---------
        df : Pandas dataframe
            Dataframe containing the messages of a Slack channel.

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
        Replace user_id with the user's display_name in df_msgs["text"].

        If there is no display_name, then "user_id" is replaced with
        "profile_real_name". The replacements are done in-place.

        In the replacements, the "@" is used to wrap the display_name for
        clarity on the text, since names can generally have more than one word
        and many names can be referenced one after the other, which can lead
        to confusion when reading.

        Arguments
        ---------
        df_msgs : Pandas dataframe
            Dataframe containing the messages of a Slack channel.
        df_usrs : Pandas dataframe
            Dataframe containing the user's information.

        """
        # Note:
        # All the bots in df_usrs have an "id" and "profile_real_name"
        # (not necessarily "name" and "display_id"). Their profile_real_name
        # are: Zoom, Google Drive, monday.com, monday.com notifications,
        # GitHub, Google Calendar, Loom, Simple Poll, Figma,
        # OneDrive and SharePoint, Calendly, Outlook Calendar,
        # Rebecca Everlene Trust Company, Slack Team Emoji,
        # New hire onboarding, Welcome, Clockify - Clocking in/out, Zapier,
        # Update Your Slack Team Icon, Jira, Google Sheets, Time Off,
        # Trailhead, Slack Team Emoji Copy, Guru, Guru, Google Calendar, Polly.
        # "USLACKBOT" and "B043CSZ0FL7" are the only bot messages if df_msgs,
        # but they are not in df_usrs!
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
        Replace parent_user_id with its display_name in the dataframe.

        If there is no display_name, then "parent_user_id" is replaced with
        "profile_real_name". The replacements are done in-place.

        Arguments
        ---------
        df_msgs : Pandas dataframe
            Dataframe containing the messages of a Slack channel.

        df_usrs : Pandas dataframe
            Dataframe contains the user's information.

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
        Replace the channel_id with the channel's name in df_msgs["text"].

        The replacements are done in-place.

        Arguments
        ---------
        df_msgs : Pandas dataframe
            Dataframe containing the messages of a Slack channel.

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
        Drop empty rows in df_msgs created from misparsed messages.

        This is the case, for example, if the text has two projects, one
        identified as "Name:" and the other as "Project Name:"

        Arguments
        ---------
        df_msgs : Pandas dataframe
            Dataframe containing the messages of a Slack channel.

        Returns
        -------
        Pandas dataframe

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
            Dataframe containing the messages of a Slack channel.

        Returns
        -------
        List with the indices of the automatic messages in the dataframe.

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
        Return a dataframe with "automatic" messages.

        Arguments
        ---------
        df_msgs : Pandas dataframe
            Dataframe containing the messages of a Slack channel.

        Returns
        -------
        Pandas dataframe

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
            Dataframe containing the messages of a Slack channel.

        Returns
        -------
        Pandas dataframe

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
            Dataframe containing the messages of a Slack channel.

        Returns
        -------
        Pandas dataframe

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
        Return a dataframe with no emojis in df_msgs["text"].

        Arguments
        ---------
        df_msgs : Pandas dataframe
            Dataframe containing the messages of a Slack channel.

        Returns
        -------
        Pandas dataframe

        """
        pattern = r"(:)([a-z0-9\_\-\+]+)(:)"
        for i in range(len(df_msgs)):
            if df_msgs.at[i, "contained_emoji"] is True:
                df_msgs.at[i, "text"] = re.sub(pattern, "", df_msgs.at[i, "text"])
        return df_msgs

    def id_short_msgs(self, df_msgs, n_char):
        """
        Identify msgs with fewer characters than n_char.

        Arguments
        ---------
        df_msgs : Pandas dataframe
            Dataframe containing the messages of a Slack channel.
        n_char : int
            Minimum number of characters needed for a message not to be
            dropped.

        Returns
        -------
        List with the indices of the short messages in the dataframe.

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
            Dataframe containing the messages of a Slack channel.
        n_char : int
            Minimum number of characters needed for a message not to be
            dropped.

        Returns
        -------
        Pandas dataframe

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
            Dataframe containing the messages of a Slack channel.
        n_char : int (optional. Default value is 15)
            Minimum number of characters needed for a message not to be
            dropped.

        Returns
        -------
        Pandas dataframe

        """
        indices2drop = self.id_short_msgs(df_msgs, n_char)
        df = df_msgs.copy()
        df.drop(indices2drop, inplace=True)
        clean.reset_indices(df)
        return df

    def add_channel_info(self, channel_path, channel_df):
        """
        Add "channel", "export_dates" and "parsed_reports_in_channel" to df.

        Requires the dataframe to have the column "projects_parsed".

        Arguments
        ---------
        channel_path : str
            Absolute path to the channels directory.

        channel_df : Pandas dataframe

        Returns
        -------
        Pandas dataframe

        """
        file_name = channel_path.split("/")[-1].split(".")[0]
        channel_name = "_".join(file_name.split("_")[:-3])
        channel_date = "_".join(file_name.split("_")[-3:])

        df_ = channel_df["projects_parsed"].astype("string")
        reports_in_channel = f"{len(df_[df_ != '0'])}/{len(channel_df)}"
        channel_df["channel"] = [channel_name]*len(channel_df)
        channel_df["export_dates"] = [channel_date]*len(channel_df)
        channel_df["parsed_reports_in_channel"] = [reports_in_channel]*len(channel_df)

        return channel_df

    def add_info_of_users_reports(self, df_all):
        """
        Get the latest report date and number of messages in the given channel.

        Arguments
        ---------
        df_all : Pandas dataframe

        Returns
        -------
        Pandas dataframe with extra columns "latest_report_date" and
        "number_msgs_in_channel".

        """
        users = df_all["user"].unique()
        for user in users:
            user_df = df_all[df_all["user"] == user].sort_values(
                by="msg_date", inplace=False, ignore_index=True
                )
            latest_report_date = user_df["msg_date"].to_list()[-1]
            user_df["latest_report_date"] = [latest_report_date]*len(user_df)
            user_df["number_msgs_in_channel"] = len(user_df)
            if user == users[0]:
                df_out = user_df.copy()
            else:
                df_out = pd.concat([df_out, user_df],
                                   axis=0, ignore_index=False)

        return df_out

    def format_parsed_reports(self, df_all):
        """
        Select the parsed weekly reports and sort the df by channel.

        Arguments
        ---------
        df_all : Pandas dataframe
            Pandas dataframe with all the messages.

        Returns
        -------
        Pandas dataframe containing messages parsed as "weekly reports".


        """
        df_p = df_all.copy()
        df_p = df_p[df_p["projects_parsed"] != "0"]
        df_p = df_p.reset_index().drop(columns=["index"])
        df_p.sort_values(by=["channel", "display_name", "msg_date"],
                         inplace=True, ignore_index=True)
        return df_p

    def format_unparsed_reports(self, df_all, wr_channel_name):
        """
        Select the unparsed weekly reports and sort the df by channel.

        Arguments
        ---------
        df_all : Pandas dataframe
            Pandas dataframe with all the messages.
        wr_ch_name : str
            Channel name of the given weekly report.

        Returns
        -------
        Pandas dataframe containing messages NOT parsed as "weekly report".

        """
        df_np = df_all.copy()
        df_np = df_np[df_np["channel"] == wr_channel_name]
        df_np = df_np[df_np["projects_parsed"] == "0"]
        df_np = df_np[df_np["msg_id"] != "channel_join"]
        df_np = df_np[df_np["is_bot"] != True]
        df_np = df_np[df_np["type"] != "thread"]
        df_np = df_np.reset_index().drop(columns=["index"])
        df_np.sort_values(by=["channel", "display_name", "msg_date"],
                          inplace=True, ignore_index=True)
        return df_np

    def format_msgs_with_urls(self, df, settings):
        """
        Select the messages that contain URL(s) in df["text"].

        Arguments
        ---------
        df : Pandas dataframe

        settings : parser.Parser(txt_path)
            Parsed variables from the settings txt file.

        Returns
        -------
        Pandas dataframe

        """
        df_p = df.copy()
        df_p = df_p[df_p['URL(s)'] != settings.get('missing_value')]
        df_p = df_p.reset_index().drop(columns=['index'])
        df_p.sort_values(by=['channel', 'display_name', 'msg_date'],
                         inplace=True, ignore_index=True)
        return df_p

    def filter_urls(self, url_list, settings):
        """
        Select a message's URLs as specified in the settings txt file.

        Arguments
        ---------
        url_list : list

        settings : parser.Parser(txt_path)
            Parsed variables from the settings txt file.

        Returns
        -------
        List URL(s) from a Slack message as specified in the settings txt file.

        """
        out = []
        for url in url_list:

            for url_exp in settings.get('urls_to_show'):
                if url_exp in url:
                    out.append(url.lstrip(' ').rstrip(' '))

        return out

    def select_desired_urls(self, df, settings):
        """
        Filter-out unwanted URL(s) as specified in the settings txt file.

        Arguments
        ---------
        df : Pandas dataframe

        settings : parser.Parser(txt_path)
            Parsed variables from the settings txt file.

        Returns
        -------
        Pandas dataframe

        """
        indices2drop = []
        for i in range(len(df)):

            urls = df.at[i, "URL(s)"].split("; ")
            filtered_urls = self.filter_urls(urls, settings)
            if filtered_urls == []:
                indices2drop.append(i)
            else:
                df.at[i, "URL(s)"] = '; '.join(filtered_urls)

        indices2drop = list(set(indices2drop))
        return df.drop(indices2drop, axis=0, inplace=False)


# #############################################################################
# FUNCTIONS
# #########

def write_info_to_file(flag, df, filename, path):
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
