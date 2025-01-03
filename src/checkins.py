import pandas as pd
import numpy as np
import clean

# #-- Introduce expected/possible keywords per report's category:
keywords_dictionary = {
    'header': ['weekly report', 'report', "week's report"],
    'project_name': ['project name'],
    'working_on': ['working on', 'working', 'what you are working on',
                   'worked on'],
    'progress_and_roadblocks': ['progress and roadblocks',
                                'progress and roadblock',
                                'progress &amp; roadblocks',
                                'Progress/Roadblocks',
                                'progress and challenges'],
    'progress': ['progress'],
    'roadblocks': ['roadblocks', 'roadblock'],
    'plans_for_following_week': ['plans for the following week',
                                 'plans for next week', 'following week',
                                 'next week', 'plans for the upcoming week'],
    'meetings': ['meetings', 'meet', 'met', "meetings you've attended",
                 'upcoming meetings', 'meeting', 'Meeting attended',
                 'Meetings attended']
}
all_keywords = ['project_name', 'working_on', 'progress_and_roadblocks',
                'progress', 'roadblocks', 'plans_for_following_week',
                'meetings']


def match_to_category(line, category_name):
    """
    Returns True if the category_name, followed by a semicolon, was found in
    a line of the message.
    """
    line = line.lower().lstrip('*-•. ').rstrip('*-•. ')
    line = line.replace('*', '').replace(' ', '')
    out = False
    for keyword in keywords_dictionary[category_name]:
        if keyword.lower().replace(' ', '')+':' in line:
            out = True
    return out


def review_format(text):
    """
    Returns a dictionary with keys:
        project_name, working_on, progress_and_roadblocks, progress,
        roadblocks, plans_for_following_week and meetings.
    And values={0,1,missing_value} depending if the above keywords were found
    in the text as a whole.
    """
    is_format_correct_list = [0]*len(all_keywords)
    is_format_correct_dict = {}

    if text != '':
        text_to_lines = text.splitlines()
        for i_line in range(len(text_to_lines)):
            line = text_to_lines[i_line]
            for i in range(len(all_keywords)):
                category_name = all_keywords[i]
                if match_to_category(line, category_name) is True:
                    is_format_correct_list[i] = 1
                    break
                # #-- Double check 'roadblocks:' vs. 'progress and roadblocks:'
                if category_name == 'roadblocks' and match_to_category(line, 'roadblocks') is True:
                    is_format_correct_list[i] = 0

    for i in range(len(all_keywords)):
        is_format_correct_dict[all_keywords[i]] = is_format_correct_list[i]

    return is_format_correct_dict


def get_indices_of_lines_with_category_name(text):
    """
    Returns two lists. One with the number of the line in the text where a
    keyword was identified and the other list with the corresponding
    category_names.
    """
    indices_start_of_category = []
    category_names = []
    if text != '':
        text_to_lines = text.splitlines()
        for i_line in range(len(text_to_lines)):
            line = text_to_lines[i_line]
            for i in range(len(all_keywords)):
                category_name = all_keywords[i]
                if match_to_category(line, category_name) is True:
                    indices_start_of_category.append(i_line)
                    category_names.append(category_name)
                    break
                # #-- Double check 'roadblocks:' vs. 'progress and roadblocks:'
                if category_name == 'roadblocks' and match_to_category(line, 'roadblocks') is True:
                    indices_start_of_category = indices_start_of_category[:-1]
                    category_names = category_names[:-1]

    return indices_start_of_category, category_names


def group_lines(text, indices_start_of_category):
    """
    Returns a list of lists, where each elements collects the content
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


def count_projects(category_names):
    """
    Returns an integer with the number of identified projects in the text.
    A project is identified in the category label is:
        "Project name:"
    independently of lowercase or uppercase letters.
    """
    counter = 0
    for name in category_names:
        if name == 'project_name':
            counter += 1
    return counter


def count_weekly_report_label(df):
    """
    Retunrs a list of dataframe indices such that the label "Weekly report:"
    was found in the corresponding text.
    """
    indices = []
    for i in range(len(df)):
        text = df.at[i, 'text']
        for line in text.splitlines():
            line = line.lower().lstrip('*-•. ').rstrip('*-•. ')
            line = line.replace('*', '')
            if 'weekly report' in line or 'weekly update' in line:
                indices.append(i)
    return indices


def extract_answers(blocks_list):
    """
    Returns a list of strings, where each string corresponds to the "answer"
    of a given category.
    It removes the category_name label, and combined multiple lines if
    necessary.
    """
    answers = []
    for block in blocks_list:
        answer_text = ''
        for line in block:
            line_matches = False
            for category in all_keywords:
                if match_to_category(line, category) is True:
                    answer_text += line.split(":")[1]
                    answer_text = answer_text.lstrip('*-•. ').rstrip('*-•. ')
                    answer_text = answer_text.replace('*', '')
                    line_matches = True
                    break
            if line_matches is False:
                answer_text += line
        answers.append(answer_text)
    return answers


def create_empty_df_with_categories(n_rows):
    """
    Returns an empty dataframe with n_rows number of rows and columns:
        project_name, working_on, progress_and_roadblocks, progress,
        roadblocks, plans_for_following_week, meetings, projects_parsed, index_
    The column index_ is for internal development of the code. It can be
    removed at the end.
    """
    columns = list(all_keywords)+['projects_parsed', 'index_']
    df = pd.DataFrame([[np.nan]*n_rows]*len(columns)).T
    df.columns = columns
    df = df.astype('object')
    return df


def id_sample_msg(df):
    sample_text_1 = """We’re working on a volunteer tracking project for HR,
        centralizing all information in Salesforce"""
    sample_text_2 = """<!channel> reposting <@U07FCQXU7Q9>'s message.
        Please adhere to it. THANK YOU."""
    sample_text_3 = "please follow this structure when posting updates"
    for i in list(df.index):
        text_i = df.at[i, 'text']
        if sample_text_1 in text_i or sample_text_2 in text_i or sample_text_3 in text_i:
            df.loc[i, 'projects_parsed'] = 'sample'
    return df


def checkin_categories_to_df_1row(df, row_entry, text,
                                  indices_start_of_category, category_names,
                                  answers):
    project_counter = 0
    if len(indices_start_of_category) > 0:
        for i in range(len(category_names)):
            category_name = category_names[i]
            if category_name == 'project_name':
                project_counter += 1
            if project_counter > 0:
                for feature in all_keywords:
                    if category_name == feature:
                        column_name = f"{category_name}_{project_counter}"
                        df.at[row_entry, column_name] = answers[i]
                        break
    return df


def parse_1row(df):
    """
    Returns a dataframe with the parsed text.
    The columns are:
       'user', 'client_msg_id', 'ts', 'json_name', 'text',
       'format_project_name', 'format_working_on',
       'format_progress_and_roadblocks', 'format_progress',
       'format_roadblocks', 'format_plans_for_following_week',
       'format_meetings', 'status', 'project_name', 'working_on',
       'progress_and_roadblocks', 'progress', 'roadblocks',
       'plans_for_following_week', 'meetings', 'projects_parsed', 'index_',
       'index','progress_and_roadblocks_combined'
    """
    # #-- Initialize a df to collect the original and parsed information:
    projects_parsed_in_columns = 3
    parsed_df = df.copy()
    for i in range(1, projects_parsed_in_columns+1):
        for feature in all_keywords:
            text = f"{feature}_{i}"
            parsed_df.at[0, text] = np.nan

    for i in range(len(df)):
        text = df.at[i, 'text']
        indices_start_of_category, category_names = get_indices_of_lines_with_category_name(text)
        projects_parsed = count_projects(category_names)
        parsed_df.at[i, 'projects_parsed'] = int(projects_parsed)
        if len(indices_start_of_category) == 0:
            continue
        else:
            blocks_list = group_lines(text, indices_start_of_category)
            answers = extract_answers(blocks_list)
            parsed_df = checkin_categories_to_df_1row(
                parsed_df, i, text, indices_start_of_category,
                category_names, answers)
        parsed_df = clean.handle_missing_values(parsed_df)

    parsed_df = id_sample_msg(parsed_df)
    return parsed_df


def checkin_categories_to_df_nrows(df, text, indices_start_of_category,
                                   category_names, answers):
    """
    Takes the empty dataframe created with the function
    "create_empty_df_with_categories(n_rows)" and fills the cells with the
    "answers" to the categories that were correctly identified in the text.
    """
    project_counter = -1
    for i in range(len(category_names)):
        if category_names[i] == 'project_name':
            project_counter += 1
        df.at[project_counter, category_names[i]] = answers[i]
    return df


def parse_nrows(df):
    """
    Returns a dataframe with the parsed text.
    Messages with weekly reports of more than one project are splitted in as
    many rows as projects in the report.
    The columns are:
        'user', 'client_msg_id', 'ts', 'json_name', 'text',
        'format_project_name', 'format_working_on',
        'format_progress_and_roadblocks', 'format_progress',
        'format_roadblocks', 'format_plans_for_following_week',
        'format_meetings', 'status', 'project_name', 'working_on',
        'progress_and_roadblocks', 'progress', 'roadblocks',
        'plans_for_following_week', 'meetings',
        'projects_parsed', 'index_', 'index','progress_and_roadblocks_combined'
    """

    # #-- Identify messages with sample text:
    df_tmp = id_sample_msg(df.copy())
    sample_msg_indices = df_tmp[df_tmp['projects_parsed'] == 'sample'].index

    for i in range(len(df)):
        text = df.at[i, 'text']
        indices_start_of_category, category_names = get_indices_of_lines_with_category_name(text)
        if i in sample_msg_indices:
            df_i_blocks = create_empty_df_with_categories(1)
            projects_parsed = 'sample'
        elif len(indices_start_of_category) == 0:
            # #-- If no keywords were identified in the non-empty text:
            df_i_blocks = create_empty_df_with_categories(1)
            projects_parsed = '0'
        elif len(indices_start_of_category) > 0:
            blocks_list = group_lines(text, indices_start_of_category)
            answers = extract_answers(blocks_list)
            projects_parsed = count_projects(category_names)
            if projects_parsed == 0:
                # #-- If project_name was not identified:
                df_i_blocks = create_empty_df_with_categories(1)
                projects_parsed = '0'
            else:
                df_i_blocks = create_empty_df_with_categories(projects_parsed)
                df_i_blocks = checkin_categories_to_df_nrows(
                    df_i_blocks, text, indices_start_of_category,
                    category_names, answers
                    )
                if projects_parsed == 1:
                    projects_parsed = '1/1'
                elif projects_parsed == 2:
                    projects_parsed = ['1/2', '2/2']
                elif projects_parsed == 3:
                    projects_parsed = ['1/3', '2/3', '3/3']
        df_i_blocks['projects_parsed'] = projects_parsed
        df_i_blocks['index_'] = i
        df_i_blocks = clean.handle_missing_values(df_i_blocks)

        # #-- Dataframe with the original text. Rows are dublicated as many
        # times as projects in the checkin:
        df_i_text = pd.DataFrame([list(df.loc[i].values)]*len(df_i_blocks))
        df_i_text.columns = df.columns
        df_i_text['index'] = i
        df_i_text = clean.handle_missing_values(df_i_text)

        # #-- Concatenate df_i_text and df_i_blocks for i-th message:
        df_i_all = pd.concat(
            [df_i_text, df_i_blocks], axis=1, ignore_index=True
            )
        df_i_all.columns = list(df_i_text.columns) + list(df_i_blocks.columns)
        df_i_all = clean.handle_missing_values(df_i_all)

        # #-- Collect df_i_all in parsed_df:
        if i == 0:
            parsed_df = df_i_all.copy()
        else:
            parsed_df = pd.concat(
                [parsed_df, df_i_all], axis=0, ignore_index=True
                )

    return parsed_df
