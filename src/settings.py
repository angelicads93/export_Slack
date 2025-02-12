#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# --Syntax to use for missing values:
missing_value = 'n/d'

# --Specify time-zone region to express the date:
timezone = 'US/Central'

# --Name of file with the channels information located in the source directory:
channels_json_name = 'channels.json'

# --Name of file with the users information located in the source directory:
users_json_name = 'users.json'

# --Name extension to the destination folder:
dest_name_ext = '_JSONs_converted'

# --Do you wish to show keywords in the cells with separated weekly-report's
# --parts?:
key_wrd_text_show = False


# --List of keywords to look for in each checkin message:
# --(It is assumend that the text follows the convention "keyword : ")
all_keywords = ['project_name', 'working_on', 'progress_and_roadblocks',
                'progress', 'roadblocks', 'plans_for_following_week',
                'meetings']

# --Which keyword do you want to use as trigger for distinguishing different
# --projects?
# --NOTE: currently chosing project_name since its use is more consistent than
# --'weekly report:'
index_keyword = 'project_name'

# --To maximize the parsing of messages, enter words/phrases that users are
# --likely to use as keywords in their checkin messages. 
# --Enter your inputs in the form of a Python dictionary as shown bellow:
keywords_dictionary = {
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


# --Some of the checkin messages are sent to the Slack channel as a sample on
# --how to properly format the expected message. In the list sample_text_list
# --bellow, enter a fragment of the text of such "sample" messages. This list
# --will be use to explicitely identifying the sample messages in the final
# --Excel sheet:
sample_text_list = [
    "We’re working on a volunteer tracking project for HR, centralizing all information in Salesforce",
    "<!channel> reposting <@U07FCQXU7Q9>'s message. Please adhere to it. THANK YOU.",
    "please follow this structure when posting updates",
    "Reminder: Hey everyone! If you haven't already, don't forget to submit your weekly check-in no later than Saturday 11:59pm local time"
]


# --Ready to start the analysis?:
continue_analysis = True