#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: Angelica Goncalves.

Module to parse the weekly reports.

Classes
-------
CheckIns

Functions
---------
parse_reports(df, settings)
    Return a dataframe with the parsed text checkin messages.

"""

# Import standard libraries:
import pandas as pd
import numpy as np
import clean


class CheckIns:
    """
    Class to handle information on the Slack channels and users.

    Relevant features are added and formatted into "channels" and "users"
    Pandas dataframes, which are later used to complement the information on
    the channel's messages.
    Requires objects of the customed classes "Parser" and  "InspectSource".

    ...

    Attributes
    ----------
    settings : parser.Parser(txt_path)
        The parsed user's inputs from the file settings_messages.txt
        Variables defined in settings_messages.txt are retrieved as
        settings.get(var_name).

    Methods
    -------
    match_to_category(line, category_name)
        Check if a text's line contains the definition of a checkin category.

    get_idx_of_lines_with_cat(text)
        Identify checkin categories in the text and their line location.

    group_lines(text, indices_start_of_category)
        Group lines on checkins message by categories.

    count_idx_kw(category_names, keyword)
        Return an integer with the number of identified projects in the text.

    id_label_in_df(df, label)
        Identify the indices of the df whose text contains the given label.

    extract_answers(blocks_list)
        Format the answer to each identified category in the checkin text.

    create_empty_df_with_cats(n_rows, cols)
        Return an empty dataframe with n_rows and given columns.

    projs_parsed_to_fraction(projs_parsed, idx, sample_msg_idx)
        Count projects_parsed as fractions of total projects.

    kws_parsed_to_fraction(projs_parsed, category_names)
        Count keywords_parsed as fractions of total keywords.

    checkin_categories_to_df(df, category_names, answers)
        Fill in the dataframe with the parsed checkin messages.

    id_sample_msg(df)
        Add "sample" in the column "projects_parsed" if msg is a sample report.

    """

    def __init__(self, settings):
        self.settings = settings
        self.missing_value = self.settings.get("missing_value")
        self.all_keywords = self.settings.get("all_keywords")
        self.keywords_dictionary = self.settings.get("keywords_dictionary")
        self.index_keyword = self.settings.get("index_keyword")
        self.sample_text_list = self.settings.get("sample_text_list")

    def match_to_category(self, line, category_name):
        """
        Check if a text's line contains the definition of a checkin category.

        Returns True if the category_name matches the text before ":" in
        the given line. Using "==" instead of "in" prevents "roadblocks" to
        trigger both the categories roadblockc and progress_and_roadblocks
        simultaneously.

        Arguments
        ---------
        line : string
            A line in the weekly report text.

        category_name : str
            Name of the checkin category.

        Returns
        -------
        out : bool

        """
        # Prepare line
        # Change all cases to lowercases and remove "*" (bold text):
        line = line.lower().replace("*", "")
        # Remove unwanted symbols and numbers (often used in bullet list):
        line = line.lstrip("*-_•.◦ 1234567890").rstrip("*-_•.◦ 1234567890")
        line = line.replace("&gt;", "")
        # Remove white spaces:
        line = line.replace(" ", "").replace("_", "")
        # Initialize boolean output as False:
        out = False
        # Check if line contains ":" needed to define a keyword:
        if ":" in line:
            for keyword in self.keywords_dictionary[category_name]:
                keyword_ = keyword.lower().replace(" ", "")
                line_ = (line.split(":")[0]).lower().replace(" ", "")
                # Check if line contains a keyword in the checkin category:
                if keyword_ == line_:
                    out = True
                    break
        return out

    def get_idx_of_lines_with_cat(self, text):
        """
        Identify checkin categories in the text and their line location.

        Arguments
        ---------
        text : str
           Text from the checkin message.

        Returns
        -------
        indices_start_of_category : list
            List with the line number of the lines where a keyword was defined.

        category_names : list
            List with the name of the categories identified in
            indices_start_of_category (keeping the order of the list).

        """
        # Initialize output lists:
        indices_start_of_category = []
        category_names = []
        # Inspect the text, line by line:
        if text != "":
            text_to_lines = text.splitlines()
            for i, line in enumerate(text_to_lines):
                for category_name in self.all_keywords:
                    if self.match_to_category(line, category_name) is True:
                        indices_start_of_category.append(i)
                        category_names.append(category_name)
                        break
        return indices_start_of_category, category_names

    def group_lines(self, text, indices_start_of_category):
        """
        Group lines on checkins message by categories.

        Arguments
        ---------
        text : str
            Text from the checkin message.

        indices_start_of_category : list
            List with the line number of lines containing a checkin keyword.
            Output of the method get_idx_of_lines_with_cat(text).

        Returns
        -------
        blocks : list
            List of list, where each elements collects the content (1 or more
            lines) for each category in the text.

        """
        # Initialize output list:
        blocks = []
        # Split the checkin text into lines:
        text_to_lines = text.splitlines()
        # Inspect all but the last identified category:
        for i in range(len(indices_start_of_category)-1):
            # Identify the first and last line defining the current category:
            begin = indices_start_of_category[i]
            end = indices_start_of_category[i+1]
            # Add the lines into new entry in the output list:
            blocks.append(text_to_lines[begin:end])
        # Add the last category to the output list:
        blocks.append(text_to_lines[end:])

        return blocks

    def count_idx_kw(self, category_names, keyword):
        """
        Return an integer with the number of identified projects in the text.

        A project is usually identified throught the label "Project name:",
        independently of lowercase or uppercase letters. It most contain ":".

        Arguments
        ---------
        category_name : str
            Name of the category.

        keyword : str
            Keyword identifying the checkin category.

        Returns
        -------
        counter : int
            Number of times the given keyword was identified in the checkin
            text.

        """
        # Initialize the output counter:
        counter = 0
        # Inspect the categories already identified in the text:
        for name in category_names:
            # Check if it is the expected keyword:
            if name == keyword:
                counter += 1
        return counter

    def id_label_in_df(self, df, label):
        """
        Identify the indices of the df whose text contains the given label.

        Arguments
        ---------
        df : Pandas dataframe

        label : str
            Name of the label to be counted.

        Returns
        -------
        indices : list
            List with the dataframe's indices where the given label was
            identified.

        """
        # Initialize output list:
        indices = []
        # Inspect each row of the dataframe:
        for i in range(len(df)):
            text = df.at[i, "text"]
            for line in text.splitlines():
                line = line.lower().lstrip("*-•. ").rstrip("*-•. ")
                line = line.replace("*", "")
                # Check if the given label is in the current checkin text:
                if label in line:
                    indices.append(i)
        return indices

    def extract_answers(self, blocks_list):
        """
        Format the answer to each identified category in the checkin text.

        Arguments
        ---------
        blocks_list : list
            List with the lines of the checkin text grouped by category.

        Returns
        -------
        answers : list
            List of strings, where each string corresponds to the
            "answer" of a given category. It removes the category_name label,
            and combined multiple lines if necessary.

        """
        # Initialize the output list:
        answers = []
        # Inspect each identified category:
        for block in blocks_list:
            ans = ""
            for line in block:
                # Initialize the line_matches boolean to False:
                line_matches = False
                # If line contains a keyworkd, add to "ans" the text after
                # ":" and remove unwanted symbols:
                for category in self.all_keywords:
                    if self.match_to_category(line, category) is True:
                        line_matches = True
                        ans += line.split(":")[1]
                        ans = ans.lstrip("*-_•.◦ 1234567890")
                        ans = ans.rstrip("*-_•.◦ 1234567890")
                        ans = ans.replace("*", "").replace("&gt;", "")
                        break
                # If line doesn't contain a keyword, add the line as it is:
                if line_matches is False:
                    ans += line
            # Check that answer is not being parsed as missing_value:
            if ans.lower() == "none":
                ans = "none "
            elif ans.lower() == "n/a":
                ans = "n/a "

            # Add the edited answer to the output list:
            answers.append(ans)

        return answers

    def create_empty_df_with_cats(self, n_rows, cols):
        """
        Return an empty dataframe with n_rows and given columns.

        The column index_ is for internal development of the code. It can be
        removed at the end.

        Arguments
        ---------
        n_rows : int
            Number of empty rows to include when defining the Pandas dataframe.

        cols : list
            List with the name of the dataframe columns.

        Returns
        -------
        df : Pandas dataframe

        """
        # Create empty dataframe with n_rows:
        df = pd.DataFrame([[np.nan]*n_rows]*len(cols)).T
        # Assign the columns of the dataframe:
        df.columns = cols
        # Assign the type of the dataframe:
        df = df.astype("object")
        return df

    def id_sample_msg(self, df):
        """
        Add "sample" in the column "projects_parsed" if msg is a sample report.

        A "sample" messages is a message that contain the instructions of how
        to properly write a check-in report.

        Arguments
        ---------
        df : Pandas dataframe

        Returns
        -------
        indices : list
            List with the index of "sample" messages.

        """
        indices = []
        for i in list(df.index):
            text_i = df.at[i, "text"]
            for sample_text in self.sample_text_list:
                if sample_text in text_i:
                    df.loc[i, "projects_parsed"] = "sample"
                    indices.append(i)
                else:
                    df.loc[i, "projects_parsed"] = ""
        return indices

    def projs_parsed_to_fraction(self, projs_parsed, idx, sample_msg_idx):
        """
        Count projects_parsed as fractions of total projects.

        Arguments
        ---------
        projs_parsed : int
            Number of projects successfully parsed in the checkin text.

        idx : int
            Dataframe index of the current checkin message.

        sample_msg_idx : list
            Indices of messages that are "sample" messages.

        Returns
        -------
        projects_parsed_str : str
            String with either "sample" or a list with the number of projects
            parsed as a fraction of the total projects in the checkin message.

        """
        projects_parsed_str = ""
        # Check in the checkin is a "sample" message:
        if idx in sample_msg_idx:
            projects_parsed_str = "sample"
        # Edit string as a fraction of the total projects in message:
        if projs_parsed == 0:
            projects_parsed_str = str(0)
        elif projs_parsed == 1:
            projects_parsed_str = "1/1"
        elif projs_parsed == 2:
            projects_parsed_str = ["1/2", "2/2"]
        elif projs_parsed == 3:
            projects_parsed_str = ["1/3", "2/3", "3/3"]

        return projects_parsed_str

    def kws_parsed_to_fraction(self, projs_parsed, category_names):
        """
        Count keywords_parsed as fractions of total keywords.

        Arguments
        ---------
        projs_parsed : int
            Number of projects successfully parsed in the checkin text.

        category_names : str
            Dataframe index of the current checkin message.

        Returns
        -------
        kws : str
            String with the numbers of keywords parsed as a fraction of the
            total keywords expected in the checkin message.

        """
        # Initialize output list:
        kws = []
        # When 0 or 1 projects are parsed:
        if projs_parsed == 0:
            kws.append("0")
        elif projs_parsed == 1:
            kws.append(f"{len(category_names)}/{len(self.all_keywords)}")
        # If more than 1 project was parsed, count parsed_keywords per project:
        elif projs_parsed > 1:
            indices = []
            for index, key in enumerate(category_names):
                if key == self.index_keyword:
                    indices.append(index)
            indices.append(index+1)
            for i in range(len(indices)-1):
                kws.append(f"{len(category_names[indices[i]:indices[i+1]])}" +
                           f"/{len(self.all_keywords)}")
        return kws

    def checkin_categories_to_df(self, df, category_names, answers):
        """
        Fill in the dataframe with the parsed checkin messages.

        Take the empty dataframe created with the function
        "create_empty_df_with_cats(n_rows)" and fills the cells with the
        "answers" to the categories that were correctly identified in the text.
        Multiple projects are stored in multiple rows, while preserving the
        identification information of the message.

        Arguments
        ---------
        df : Pandas dataframe

        category_name : list
            Name of the checkin category.

        answers : list
            List of strings, where each string corresponds to the
            "answer" of a given category.

        Returns
        -------
        df : Pandas dataframe

        """
        project_counter = -1
        for i, cat in enumerate(category_names):
            if cat == self.index_keyword:
                project_counter += 1
            df.at[project_counter, cat] = answers[i]
        return df


def parse_reports(df, settings):
    """
    Return a dataframe with the parsed text checkin messages.

    Messages with weekly reports of more than one project are splitted in
    as many rows as projects in the report.

    Arguments
    ---------
    df : Pandas dataframe

    settings : parser.Parser(txt_path)
        The parsed user's inputs from the file settings_messages.txt
        Variables defined in settings_messages.txt are retrieved as
        settings.get(var_name).

    Returns
    -------
    parsed_df : Pandas dataframe

    """
    # Create instance of class CheckIns:
    ci = CheckIns(settings)

    # Identify rows with sample text:
    sample_msg_idx = ci.id_sample_msg(df.copy())

    for i in range(len(df)):

        # Build the dataframe with the columns pertaining to a check-in
        # message.
        cat_idx, cat_names = ci.get_idx_of_lines_with_cat(df.at[i, "text"])
        projs_parsed = ci.count_idx_kw(cat_names, ci.index_keyword)

        # Create empty df with checkin categories:
        cols = list(ci.all_keywords) + ["projects_parsed",
                                          "keywords_parsed", "index_"]
        if projs_parsed == 0 or i in sample_msg_idx:
            df_i_report = ci.create_empty_df_with_cats(1, cols)
        else:
            df_i_report = ci.create_empty_df_with_cats(projs_parsed, cols)

        # Fill the empty df with the parsed categories:
        if len(cat_idx) > 0 and projs_parsed > 0:
            blocks_list = ci.group_lines(df.at[i, "text"], cat_idx)
            answers = ci.extract_answers(blocks_list)
            df_i_report = ci.checkin_categories_to_df(
                df_i_report, cat_names, answers)

        # Count the number of projects and keywords parsed:
        df_i_report["projects_parsed"] = ci.projs_parsed_to_fraction(
            projs_parsed, i, sample_msg_idx)
        df_i_report["keywords_parsed"] = ci.kws_parsed_to_fraction(
            projs_parsed, cat_names)

        df_i_report["index_"] = i
        df_i_report = clean.handle_missing_values(df_i_report, ci.missing_value)

        # Build the dataframe with the original information and full text.
        # Rows are dublicated as many times as projects in the checkin msg:
        df_i_msg = pd.DataFrame([list(df.loc[i].values)]*len(df_i_report))
        df_i_msg.columns = df.columns
        df_i_msg["index"] = i
        df_i_msg = clean.handle_missing_values(df_i_msg, ci.missing_value)

        # Horizontally concatenate df_i_msg and df_i_report for i-th
        # message:
        df_i_all = pd.concat(
            [df_i_msg, df_i_report], axis=1, ignore_index=True)
        df_i_all.columns = list(df_i_msg.columns) + list(df_i_report.columns)
        df_i_all = clean.handle_missing_values(df_i_all, ci.missing_value)

        # Initialized the parsed_df dataframe if parsing the first row:
        if i == 0:
            parsed_df = df_i_all.copy()
        # Vertically concatenate parsed_df and df_i_all:
        else:
            parsed_df = pd.concat(
                [parsed_df, df_i_all], axis=0, ignore_index=True)

    return parsed_df
