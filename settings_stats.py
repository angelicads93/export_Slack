#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Feb 25 16:34:14 2025

@author: agds
"""

# --Syntax to use for missing values:
missing_value = 'n/d'


# ############################################################################
# ############################################################################
# ### DIRECTORIES AND FILE NAMES:
# ##################################
#
# 0. Path where the exported JSON files are:
jsons_source_path = '/home/agds/Documents/RET/Source/Jan_1_2025_to_Jan_31_2025/'
#
# 1. Path where the exported Excel files are:
excel_channels_path = '/home/agds/Documents/RET/Source/Jan_1_2025_to_Jan_31_2025/_JSONs_converted'
#
# 2. Name of the Slack channel dedicated to the weekly reports:
reports_channel_name = 'think-biver-weekly-checkins'
#
# 3. Name of file to be saved with the compilation of check-in weekly reports:
compilation_reports_file_name = 'test_reports_v4.xlsx'
#
# 4. Path where the file compilation_reports_file_path will be saved:
compilation_reports_path = '/home/agds/Desktop'



# ############################################################################
# ############################################################################
# ### EXCEL FORMATTING:
# #####################
#
# 0. Set the order of the columns in the Excel table:
columns_order = ['channel', 'export_dates', 'parsed_reports_in_channel',
                 'user', 'name', 'display_name', 'deactivated', 'is_bot',
                 'msg_id', 'msg_date', 'type', 'text',
                 'reply_count', 'reply_users_count', 'latest_reply_date',
                 'thread_date', 'parent_user_name', 'URL(s)',
                 'number_msgs_in_channel', 'projects_parsed',
                 'keywords_parsed', 'latest_report_date', 'project_name',
                 'working_on', 'progress_and_roadblocks', 'progress',
                 'roadblocks', 'plans_for_following_week', 'meetings']
#
# 1. Change the width of the Excel cells:
w_date = 19
w_name = 19
w_text = 30
w_bool = 7
w_count = 8
column_widths = {
    'channel': w_text, 'export_dates': 25, 'parsed_reports_in_channel': w_count,
    'msg_id': 12, 'msg_date': w_date, 'user': 15, 'name': w_name,
    'display_name': w_name, 'deactivated': w_bool, 'is_bot': w_bool,
    'type': 8, 'text': w_text, 'reply_count': w_count,
    'reply_users_count': w_count, 'latest_reply_date': w_date,
    'thread_date': w_date, 'parent_user_name': w_name,
    'URL(s)': w_text, 'number_msgs_in_channel': w_count,
    'projects_parsed': w_count, 'keywords_parsed': w_count,
    'project_name': w_text, 'working_on': w_text,
    'progress_and_roadblocks': w_text, 'progress': w_text,
    'roadblocks': w_text, 'plans_for_following_week': w_text,
    'meetings': w_text, 'latest_report_date': w_date
}
# -----------------------------------------------------------------------------
# 2. Set the alignment of cells containing long text:
text_type_cols = ['L', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC']
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
    ("e7c9fb", ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']),
    ("CDB5B7", ['C', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC'])
]
# -----------------------------------------------------------------------------
# 4. Distinguish special columns by changing the color of their font:
font_color_in_column = [
    # ["column_letter', 'color_code']
    ('L', "0707C5"),  # "text" to blue
    ('Q', "c10105"),  # "parent_user_name" to red
    ('F', "c10105")   # "display_name" to red
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
    "trigger": ["H", "==", True],
    "columns": ["D", "E", "F", "G", "H", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC"],
    "cell_color": "FBBF8F",
    "font_size": 11,
    "font_bold": False,
    "font_horiz_alignment": "left"
    },
    # 5.2. Highlight cases where the message is part of a thread.
    {"activate": True,
    "trigger": ["K", "==", "thread"],
    "columns": ["K", "L", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC"],
    "cell_color": "FBFB99",
    "font_size": 11,
    "font_bold": False,
    "font_horiz_alignment": "left"
    },
    # 5.3. Highlight the cases where there is reply_counts.
    {"activate": True,
    "trigger": ["M", "!=", "n/d"],
    "columns": ["M", "N"],
    "cell_color": "No Fill",
    "font_size": 12,
    "font_bold": True,
    "font_horiz_alignment": "right"
    }
]
# -----------------------------------------------------------------------------
# 6. Draw vertical lines at the right of the given columns:
draw_vert_line = {
    "columns": ["C", "H", "R"],
    "thickness": "medium"
    }
