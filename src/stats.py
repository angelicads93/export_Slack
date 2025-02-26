#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb 23 10:59:17 2025

@author: agds
"""
import numpy as np
import pandas as pd
import sys
import os

parent_dir = os.path.dirname(os.getcwd())
sys.path.append(parent_dir)
sys.path.append(os.path.join(parent_dir, 'src'))
import inputs
import settings
import settings_stats
import excel


missing_value = settings.missing_value



def apply_excel_adjustments(file_path, settings_mod):
    """ Defines the sequence of changes to be done in the Excel file 
    given the user's inputs in the module settings_mod.
    """
    xl = excel.ExcelFormat(file_path)
    xl.set_cell_width(settings_mod.column_widths)
    xl.set_allignment('top')
    xl.format_first_row(
            settings_mod.height_1strow,
            settings_mod.alignment_vert_1strow,
            settings_mod.alignment_horiz_1strow,
            settings_mod.font_size_1strow, 
            settings_mod.font_bold_1strow,
            settings_mod.cell_color_1strow            
            )
    for cc in settings_mod.font_color_in_column:
        xl.set_font_color_in_column(cc)
    for highlight in settings_mod.highlights:
        xl.format_highlight(highlight)
    for column in settings_mod.text_type_cols:
        xl.format_text_cells(column)
    xl.save_changes()
    
    
path_converted = f"{inputs.converted_directory}/{settings.dest_name_ext}"    
for file in os.listdir(path_converted):
    channel_path = f"{path_converted}/{file}"
    if "/.~lock." in channel_path:  # avoid hidden files, if any.
        continue
    else:
        channel_df = pd.read_excel(channel_path, engine='openpyxl')
    
        channel_name = "_".join(channel_path.split("/")[-1].split(".")[0].split('_')[:-3])
        channel_date = " ".join(channel_path.split('/')[-1].split(".")[0].split('_')[-3:])
    
        df_ = channel_df['projects_parsed'].astype('string')
        reports_in_channel = f"{len(df_[df_ != '0'])}/{len(channel_df)}"
    
        channel_df['channel'] = [channel_name]*len(channel_df)
        channel_df['export_dates'] = [channel_date]*len(channel_df)
        channel_df['reports_in_channel'] = [reports_in_channel]*len(channel_df)
    
    
        print('--------------------------------------------------------------')
        print(channel_path)
        print(channel_name)
        print(channel_date)
        print(reports_in_channel)   
    
        users = channel_df['user'].unique()
        for user in users:
            user_df = channel_df[channel_df['user'] == user].sort_values(
                by='msg_date', inplace=False, ignore_index=True
                )
            latest_report_date = user_df['msg_date'].to_list()[-1]
            user_df['latest_report_date'] = [latest_report_date]*len(user_df)
            user_df['number_msgs_in_channel'] = len(user_df)
            if user == users[0]:
                channel_df_ = user_df.copy()
            else:
                channel_df_ = pd.concat(
                    [channel_df_, user_df], axis=0, ignore_index=False
                    )
    
        # --Handle missing values:
        channel_df_ = channel_df_.replace(pd.NaT, missing_value)
        channel_df_ = channel_df_.replace(np.nan, missing_value)
        channel_df_ = channel_df_.fillna(missing_value)
    
        # --Sort columns:
        channel_df_.sort_values(
            by=['display_name', 'msg_date'], inplace=True, ignore_index=True
            )
    
        # --Reorder columns:
        columns_order = ['channel', 'export_dates', 'reports_in_channel',
                         'user', 'name', 'display_name', 'deactivated', 'is_bot',
                         'msg_id', 'msg_date', 'type', 'text',
                         'reply_count', 'reply_users_count', 'latest_reply_date',
                         'thread_date', 'parent_user_name', 'URL(s)',
                         'number_msgs_in_channel', 'projects_parsed',
                         'keywords_parsed', 'latest_report_date', 'project_name',
                         'working_on', 'progress_and_roadblocks', 'progress',
                         'roadblocks', 'plans_for_following_week', 'meetings']
        channel_df_ = channel_df_[columns_order]
    
        if file == os.listdir(path_converted)[0]:
            df = channel_df_.copy()
        else:
            df = pd.concat([df, channel_df_], axis=0, ignore_index=False)
    
# --Sort columns:
df.sort_values(
    by=['channel', 'display_name', 'msg_date'],
    inplace=True, ignore_index=True
    )

# --Set columns types:
df['projects_parsed'] = df['projects_parsed'].astype('string')
df['keywords_parsed'] = df['keywords_parsed'].astype('string')

# --Select rows with parsed projects:
df = df[df['projects_parsed'] != '0']
df = df.reset_index().drop(columns=['index'])
    
print('Information collected')

# --Save Excel file:
df.to_excel(f"{path_converted}/../test_excel_v3.xlsx", index=False)
apply_excel_adjustments(f"{path_converted}/../test_excel_v3.xlsx", settings_stats)