#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# --Syntax to use for missing values:
missing_value = 'n/d'

# --Specify time-zone region to express the date:
timezone = 'US/Central'


# ############################################################################
# ############################################################################
# ### DIRECTORIES AND FILE NAMES:
# ##################################
#
# 1. Name of file with the channels information located in the source directory:
channels_json_name = 'channels.json'
#
# -----------------------------------------------------------------------------
# 2. Name of file with the users information located in the source directory:
users_json_name = 'users.json'
#
# -----------------------------------------------------------------------------
# 3. Name extension to the destination folder:
dest_name_ext = '_JSONs_converted'


# ############################################################################
# ############################################################################
# ### KEYWORDS OF CHECK-IN MESSAGES:
# ##################################
#
# 1. List of keywords to look for in each check-in message:
# (It is assumend that the text follows the convention "keyword : ")
all_keywords = ['project_name', 'working_on', 'progress_and_roadblocks',
                'progress', 'roadblocks', 'plans_for_following_week',
                'meetings']
#
# -----------------------------------------------------------------------------
# 2. Which keyword do you want to use as trigger for distinguishing different
# projects?
# NOTE: currently chosing project_name since its use is more consistent than
# 'weekly report:'
index_keyword = 'project_name'
#
# -----------------------------------------------------------------------------
# 3. To maximize the parsing of messages, enter words/phrases that users are
# likely to use as keywords in their check-in messages.
# Enter your inputs in the form of a Python dictionary as shown bellow:
keywords_dictionary = {
    # 'keyword': ['keyword/phrase 1', ..., 'keyword/phrase n'],
    'header':
        ['weekly report', 'report', "week's report"],
    'project_name':
        ['project name', 'Current Project Name', 'project'],
    'working_on':
        ['working on', 'working', 'what you are working on', 'worked on'],
    'progress_and_roadblocks':
        ['progress and roadblocks', 'progress and roadblock',
        'progress &amp; roadblocks', 'Progress/Roadblocks',
        'progress and challenges'],
    'progress':
        ['progress'],
    'roadblocks':
        ['roadblocks', 'roadblock'],
    'plans_for_following_week':
        ['plans for the following week', 'plans for next week',
        'following week', 'next week', 'plans for the upcoming week'],
    'meetings':
        ['meetings', 'meet', 'met', "meetings you've attended",
        'upcoming meetings', 'meeting', 'Meeting attended', 'Meetings attended']
}
# -----------------------------------------------------------------------------
# 4. Some of the checkin messages are sent to the Slack channel as a sample on
# how to properly format the expected message. In the list sample_text_list
# bellow, enter a fragment of the text of such "sample" messages. This list
# will be use to explicitely identifying the sample messages in the final
# Excel sheet:
sample_text_list = [
    "We’re working on a volunteer tracking project for HR, centralizing all information in Salesforce",
    "<!channel> reposting <@U07FCQXU7Q9>'s message. Please adhere to it. THANK YOU.",
    "please follow this structure when posting updates",
    "Reminder: Hey everyone! If you haven't already, don't forget to submit your weekly check-in no later than Saturday 11:59pm local time"
]

# ############################################################################
# ############################################################################
# ### EXCEL FORMATTING:
# #####################
#
# 1. Change the width of the Excel cells:
w_date = 19
w_name = 19
w_text = 30
w_bool = 7
w_count = 8
column_widths = {
    'msg_id': 12, 'msg_date': w_date, 'user': 15, 'name': w_name,
    'display_name': w_name, 'deactivated': w_bool, 'is_bot': w_bool,
    'type': 8, 'text': w_text, 'reply_count': w_count,
    'reply_users_count': w_count, 'latest_reply_date': w_date,
    'thread_date': w_date, 'parent_user_name': w_name,
    'URL(s)': w_text, 'projects_parsed': w_count, 'keywords_parsed': w_count,
    'project_name': w_text, 'working_on': w_text,
    'progress_and_roadblocks': w_text, 'progress': w_text,
    'roadblocks': w_text, 'plans_for_following_week': w_text,
    'meetings': w_text
}
# -----------------------------------------------------------------------------
# 2. Set the alignment of cells containing long text:
text_type_cols = ['I', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']
#
# -----------------------------------------------------------------------------
# 3. Format the height, color, font alignment and size of the first column:
height_1strow = 43
alignment_vert_1strow = "top"
alignment_horiz_1strow = "left"
font_size_1strow = 9
font_bold_1strow = True
cell_color_1strow = [
    # ("color_code", [list the columns' letter])
    ("e7c9fb", ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']),
    ("CDB5B7", ['P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X'])
]
# -----------------------------------------------------------------------------
# 4. Distinguish special columns by changing the color of their font:
font_color_in_column = [
    # ["column_letter', 'color_code']
    ('I', "0707C5"),  # "text" to blue
    ('N', "c10105"),  # "parent_user_name" to red
    ('E', "c10105")   # "display_name" to red
]
# -----------------------------------------------------------------------------
# 5. Change the color and/or font on chosen cells based on a certain condition
# you wish to highlight:
highlights = [
    # 5.0. Highlight format:
    # {"activate": True/False,
    # "trigger": ["column_letter", "condition", value],
    # "columns": [list of columns to highlight],
    # "cell_color": "color_code",
    # "font_size": #,
    # "font_bold": True/False,
    # "font_horiz_alignment": "alignment_type"
    # },
    # 5.1. Highlight cases where the user is a bot.
    {"activate": True,
    "trigger": ["G", "==", True],
    "columns": ["C", "D", "E", "F", "G", "P", "Q", "R", "S", "T", "U", "V", "W", "X"],
    "cell_color": "FBBF8F",
    "font_size": 11,
    "font_bold": False,
    "font_horiz_alignment": "left"
    },
    # 5.2. Highlight cases there the message is part of a thread.
    {"activate": True,
    "trigger": ["H", "==", "thread"],
    "columns": ["H", "I", "P", "Q", "R", "S", "T", "U", "V", "W", "X"],
    "cell_color": "FBFB99",
    "font_size": 11,
    "font_bold": False,
    "font_horiz_alignment": "left"
    },
    # 5.3. Highlight the cases where there is reply_counts.
    {"activate": True,
    "trigger": ["J", "!=", "n/d"],
    "columns": ["J", "K"],
    "cell_color": "No Fill",
    "font_size": 12,
    "font_bold": True,
    "font_horiz_alignment": "right"
    }
]
#


# --Ready to start the analysis?:
continue_analysis = True