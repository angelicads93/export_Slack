#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
import clean


class CheckIns:
    def __init__(self, settings):
        self.settings = settings
        self.missing_value = self.settings.get('missing_value')
        self.all_keywords = self.settings.get('all_keywords')
        self.keywords_dictionary = self.settings.get('keywords_dictionary')
        self.index_keyword = self.settings.get('index_keyword')
        self.sample_text_list = self.settings.get('sample_text_list')
        
        
    def match_to_category(self, line, category_name):
        """ Returns True if the category_name matches the text before ':' in
        the given line. Using "==" instead of "in" prevents 'roadblocks' to trigger
        both the categories roadblockc and progress_and_roadblocks simultaneously.
        """
        line = line.lower().lstrip('*-_•.◦ 1234567890').rstrip('*-_•.◦ 1234567890')
        line = line.replace('*', '').replace(' ', '').replace('_', '')
        line = line.replace("&gt;", "")
        out = False
        if ":" in line:
            for keyword in self.keywords_dictionary[category_name]:
                keyword_ = keyword.lower().replace(' ', '')
                line_ = (line.split(':')[0]).lower().replace(' ', '')
                if keyword_ == line_:
                    out = True
                    break
        return out
    
    
    def get_indices_of_lines_with_category_name(self, text):
        """ Returns two lists. One with the number of the line in the text where a
        keyword was identified and the other list with the corresponding
        category_names.
        """
        indices_start_of_category = []
        category_names = []
        if text != '':
            text_to_lines = text.splitlines()
            for i_line in range(len(text_to_lines)):
                line = text_to_lines[i_line]
                for category_name in self.all_keywords:
                    if self.match_to_category(line, category_name) is True:
                        indices_start_of_category.append(i_line)
                        category_names.append(category_name)
                        break
        return indices_start_of_category, category_names
    
    
    def group_lines(self, text, indices_start_of_category):
        """ Returns a list of lists, where each elements collects the content
        (1 or more lines) for each category in the text.
        """
        blocks = []
        begin = indices_start_of_category[0]
        end = indices_start_of_category[-1]
        text_to_lines = text.splitlines()
        for i in range(len(indices_start_of_category)-1):
            begin = indices_start_of_category[i]
            end = indices_start_of_category[i+1]
            blocks.append(text_to_lines[begin:end])
        blocks.append(text_to_lines[end:])
        return blocks
    
    
    def count_index_keyword(self, category_names, keyword):
        """ Returns an integer with the number of identified projects in the text.
        A project is identified throught the label "Project name:", independently
        of lowercase or uppercase letters.
        """
        counter = 0
        for name in category_names:
            if name == keyword:
                counter += 1
        return counter
    
    
    def count_label_in_df(self, df, label):
        """ Collects the indices of a given dataframe, if the label
        'Weekly report:' was found in the corresponding text. Returns a list.
        """
        indices = []
        for i in range(len(df)):
            text = df.at[i, 'text']
            for line in text.splitlines():
                line = line.lower().lstrip('*-•. ').rstrip('*-•. ')
                line = line.replace('*', '')
                if label in line:
                    indices.append(i)
        return indices
    
    
    def extract_answers(self, blocks_list):
        """ Returns a list of strings, where each string corresponds to the
        "answer" of a given category. It removes the category_name label, and
        combined multiple lines if necessary.
        """
        answers = []
        for block in blocks_list:
            answer_text = ''
            for line in block:
                line_matches = False
                for category in self.all_keywords:
                    if self.match_to_category(line, category) is True:
                        answer_text += line.split(":")[1]
                        answer_text = answer_text.lstrip('*-_•.◦ 1234567890').rstrip('*-_•.◦ 1234567890')
                        answer_text = answer_text.replace('*', '')
                        answer_text = answer_text.replace("&gt;","")
                        line_matches = True
                        break
                if line_matches is False:
                    answer_text += line
            # --Check that answer is not being parsed as missing_value:
            if answer_text.lower() == 'none':
                answer_text = 'none '
            elif answer_text.lower() == 'n/a':
                answer_text = 'n/a '
    
            answers.append(answer_text)
    
        return answers
    
    
    def create_empty_df_with_categories(self, n_rows):
        """ Returns an empty dataframe with n_rows number of rows and columns:
            all_keywords + [projects_parsed, index_]
        The column index_ is for internal development of the code. It can be
        removed at the end.
        """
        cols = list(self.all_keywords)+['projects_parsed', 'keywords_parsed', 'index_']
        df = pd.DataFrame([[np.nan]*n_rows]*len(cols)).T
        df.columns = cols
        df = df.astype('object')
        return df
    
    
    def id_sample_msg(self, df):
        """ Checks the rows of a given dataframe and adds the value 'sample' to the
        projects_parsed column if the corresponding text is a sample-message with
        the instructions of how to properly write a check-in message."""
        for i in list(df.index):
            text_i = df.at[i, 'text']
            for sample_text in self.sample_text_list:
                if sample_text in text_i:
                    df.loc[i, 'projects_parsed'] = 'sample'
                else:
                    df.loc[i, 'projects_parsed'] = ''
        return df
        """ Returns a dataframe after parsing the text of each row.
        Multiple projects are stored in multiple columns in the dataframe,
        identified with a subindex i={1,2,3}.
        """
        # --Initialize a df to collect the original and parsed information:
        parsed_df = df.copy()
        projects_parsed_in_columns = 3
        for i in range(1, projects_parsed_in_columns+1):
            for feature in self.all_keywords:
                text = f"{feature}_{i}"
                parsed_df.at[0, text] = np.nan
    
        # --Fill the columns with the check-in categories (initialized to NA) with
        # --the parsed check-in message:
        for i in range(len(df)):
            text = df.at[i, 'text']
            indices_start_of_category, category_names = self.get_indices_of_lines_with_category_name(text)
            projects_parsed = self.count_label_in_df(category_names)
            parsed_df.at[i, 'projects_parsed'] = int(projects_parsed)
            if len(indices_start_of_category) == 0:
                continue
            else:
                blocks_list = self.group_lines(text, indices_start_of_category)
                answers = self.extract_answers(blocks_list)
                parsed_df = self.checkin_categories_to_df_nrows(
                    parsed_df, i, text, indices_start_of_category,
                    category_names, answers
                    )
            parsed_df = clean.handle_missing_values(parsed_df)
    
        # --Identify rows with 'sample' text:
        parsed_df = self.id_sample_msg(parsed_df)
    
        return parsed_df
    
    
    def projects_parsed_to_fraction(self, projects_parsed, i, sample_msg_indices):
        """ Count projects_parsed as fractions of total projects """
        if i in sample_msg_indices:
            projects_parsed_str = 'sample'
        if projects_parsed == 0:
            projects_parsed_str = str(0)
        elif projects_parsed == 1:
            projects_parsed_str = '1/1'
        elif projects_parsed == 2:
            projects_parsed_str = ['1/2', '2/2']
        elif projects_parsed == 3:
            projects_parsed_str = ['1/3', '2/3', '3/3']
        return projects_parsed_str
    
    
    def keywords_parsed_to_fraction(self, projects_parsed, category_names):
        keywords_parsed_str = []
        if projects_parsed == 0:
            keywords_parsed_str.append("0")
        elif projects_parsed == 1:
            keywords_parsed_str.append(
                f"{len(category_names)}/{len(self.all_keywords)}"
                )
        elif projects_parsed > 1:
            indices = []
            for index, key in enumerate(category_names):
                if key == self.index_keyword:
                    indices.append(index)
            indices.append(index+1)
            for i in range(len(indices)-1):
                keywords_parsed_str.append(
                    f"{len(category_names[indices[i]:indices[i+1]])}/{len(self.all_keywords)}"
                    )
        return keywords_parsed_str
    
    
    def checkin_categories_to_df_nrows(self, df, text,
                                       indices_start_of_category, category_names,
                                       answers):
        """ Takes the empty dataframe created with the function
        "create_empty_df_with_categories(n_rows)" and fills the cells with the
        "answers" to the categories that were correctly identified in the text.
        Multiple projects are stored in multiple rows, while preserving the
        identification information of the message.
        """
        project_counter = -1
        for i in range(len(category_names)):
            if category_names[i] == self.index_keyword:
                project_counter += 1
            df.at[project_counter, category_names[i]] = answers[i]
        return df
    
    
    def parse_nrows(self, df):
        """ Returns a dataframe with the parsed text.
        Messages with weekly reports of more than one project are splitted in as
        many rows as projects in the report.
        """
        # --Identify rows with sample text:
        df_tmp = self.id_sample_msg(df.copy())
        sample_msg_indices = df_tmp[df_tmp['projects_parsed'] == 'sample'].index
    
        for i in range(len(df)):
            # --Build the dataframe with the columns pertaining to a check-in
            # --message.
            text = df.at[i, 'text']
            indices_start_of_category, category_names = self.get_indices_of_lines_with_category_name(text)
            projects_parsed = self.count_index_keyword(category_names, self.index_keyword)
            # --Create empty df with categories:
            if projects_parsed == 0 or i in sample_msg_indices:
                df_i_blocks = self.create_empty_df_with_categories(1)
            else:
                df_i_blocks = self.create_empty_df_with_categories(projects_parsed)
            # --Fill the empty df with categories:
            if len(indices_start_of_category) > 0 and projects_parsed > 0:
                blocks_list = self.group_lines(text, indices_start_of_category)
                answers = self.extract_answers(blocks_list)
                df_i_blocks = self.checkin_categories_to_df_nrows(
                    df_i_blocks, text, indices_start_of_category,
                    category_names, answers
                    )
            # --Count the number of projects and keywords parsed:
            df_i_blocks['projects_parsed'] = self.projects_parsed_to_fraction(
                projects_parsed, i, sample_msg_indices
                )
            df_i_blocks['keywords_parsed'] = self.keywords_parsed_to_fraction(
                projects_parsed, category_names
                )
    
            df_i_blocks['index_'] = i
            df_i_blocks = clean.handle_missing_values(df_i_blocks, self.missing_value)
    
            # --Build the dataframe with the original information and full text.
            # --Rows are dublicated as many times as projects in the check-in msg:
            df_i_text = pd.DataFrame([list(df.loc[i].values)]*len(df_i_blocks))
            df_i_text.columns = df.columns
            df_i_text['index'] = i
            df_i_text = clean.handle_missing_values(df_i_text, self.missing_value)
    
            # --Horizontally concatenate df_i_text and df_i_blocks for i-th
            # --message:
            df_i_all = pd.concat(
                [df_i_text, df_i_blocks], axis=1, ignore_index=True
                )
            df_i_all.columns = list(df_i_text.columns) + list(df_i_blocks.columns)
            df_i_all = clean.handle_missing_values(df_i_all, self.missing_value)
    
            # --Initialized the parsed_df dataframe if parsing the first row:
            if i == 0:
                parsed_df = df_i_all.copy()
            # --Vertically concatenate parsed_df and df_i_all:
            else:
                parsed_df = pd.concat(
                    [parsed_df, df_i_all], axis=0, ignore_index=True
                    )
    
        return parsed_df
