# Slack Workspace to Excel Databases
This work was developed to assist the internal functioning of an organization 
with more than 1400 users and 114 different Slack channels. The goal was to 
collect, clean, and arrange all the information in the Slack workspace 
into Excel tables that project managers can then use to efficiently lead and 
guide the progress of the multiple projects being developed. 

Furthermore, it is customary for users of the Slack workspace to post their 
progress every week. Such weekly reports are parsed into categories and 
collected in the Excel databases.


## 1. Excel database per Slack channel
The messages on each Slack channel are exported into a unique Excel workbook.
The name of each Excel file consists of the name of the corresponding channel 
and the start and end date of the exported messages: 
<channel_name>_<YYYY_MM_DD>_to_<YYYY_MM_DD>.xlsx

The messages of the Slack channels are pruned based on their value and relevance 
into three worksheets or tabs: "Relevant messages", "Filtered-out messages", 
and "All messages". The "relevant messages" do not contain messages sent by a
bot user, automatic messages sent after an action has been taken on the Slack 
channel (changed on name, if a user joined or left the channel, etc), nor short 
messages usually associated with short greetings and polite replies (e.g. "Hi 
everyone", "Ok", "Thank you"). Emojis were also removed (e.g. ":simple_smile:").

The information of each Slack channel is arranged into the columns:

* msg_id: Identification code of each message (e.g. 7ca6388d-a828-4614-bc60-4b3661fa2336).
* msg_date: Date when the message was sent (e.g. 2025-01-25 18:13:46).
* user: Unique user identification code (e.g. U02063W7Z1V).
* name: User's name.
* display_name: User's name as display in the Slack profile.
* deactivated: Whether the user is active or not in the Slack workspace (e.g. True, False).
* is_bot: Whether the user is a bot (e.g. True, False).
* type: Type of correspondence, either a new message or a conversation thread (e.g. message, thread).
* text: User's message.
* contained_emoji: Whether the user's message contained any emojis (e.g. True, False).
* reply_count: Number of replies on a message thread.
* reply_users_count: 
* latest_reply_date: Date of the latest reply to the message thread (e.g. 2025-01-20 16:00:00).
* thread_date: 
* parent_user_name: The user's name that started the corresponding thread.
* URL(s): Lists any links present in the user's message.
* projects_parsed: Count the number of projects parsed in the user's message.
* keywords_parsed: Count the number of categories correctly parsed in the user's message.
* project_name: Name of the project being reported.
* working_on: Short statement with the task at hand.
* progress_and_roadblocks: List of accomplishments and difficulties found on the task at hand. 
* progress: List of accomplishments on the task at hand. Will substitute progress_and_roadblocks in the future.
* Roadblocks: List of difficulties found in the task at hand. Will substitute progress_and_roadblocks in the future.
* plans_for_following_week: List of things to do in the following week.
* meetings: List meetings and any communication with other project members. 

In addition to the Excel workbooks of each Slack channel, it
can be chosen to generate an Excel workbook with the compiled information of 
all the channels and all the users in the Slack workspace. 

### 1.1 Provide your inputs
The first step is to provide the information requested in the
files `inputs.py` and `settings_messages.py`. The `inputs.py` file contains the
 minimal information necessary to generate the Excel database of the Slack 
 channel(s) of choice. 
|INPUTS | Description| Example/Format |
|---|---|---|
|slackexport_folder_path | Path to the source directory containing all the information exported from the Slack workspace. | r"C:\Users\user_name\Documents\slackSource\January" |
|converted_directory| Path to the destination directory where the analysis results will be saved. | r"C:\Users\user_name\Documents\slackSource\Excel" |
|chosen_channel_name | 'name-of-slack-channel', if analyzing only one Slack channel, or '' if analyzing all the Slack channels in the source directory. | "general" |
|write_all_channels_info | True/False to generate a file with information on all the Slack channels. | True |
|write_all_users_info | True/False to generate a file with the information of all the Slack users. | True |

The `settings_messages.py` file contains variables necessary for formatting the 
information in the databases and the style of the Excel tables. 
|SETTINGS_MESSAGES | Description |  Example/Format |
|---|---|---|
|timezone | Time zone to express the dates in the Excel tables. | "US/Central" |
|missing_value | String that identifies missing values in the Excel tables. | "n/d" |
|channels_json_name | Name of the JSON file exported from the Slack workspace containing the channels' information.| "channels.json" |
|users_json_name | Name of the JSON file exported from the Slack workspace containing the user's information. | "users.json" |
|dest_name_ext | Name of the folder created in the destination directory where all the files will be stored. | "_JSON_converted"|
|channels_excel_name | Name of the Excel file where the channels' information will be saved.| "_channels.xlsx" |
|users_excel_name | Name of the Excel file where the user's information will be saved. | "_users.xlsx" |
|all_keywords | Keywords used to parse the check-in messages. | ["project_name", "working_on", "progress"] |
|index_keyword | Keyword used to identify each check-in message. | "project_name" |
|keywords_dictionary | Words/phrases that users are likely to use as keywords in their check-in messages.| {"working_on": ["working on", "worked on"]} |
|sample_text_list | Fragments of messages sent as samples showing the format of the weekly check-in messages. | ["please follow this structure when posting updates"] 
|columns_order | List the names of the columns in the order in which they will appear in the Excel table. | columns_order = [column_A_name, column_B_name, ...] |
|column_widths | Set the width of the Excel columns.| {"col1_name": col1_width, "col2_name": col2_width} |
|text_type_cols | List the letters of the Excel columns containing long text to set their cell alignment. | ["A", "B", "C"] |
|header_row | Set the format of the header row. |{"height": 43, "alignment_vert": "top", "alignment_horiz": "left", "font_size": 9, "font_bold": True, "cell_color": [("e7c9fb", ['A', 'B', 'C']), ("CDB5B7", ['S', 'T', 'U'])]}|
|font_color_in_column | Distinguish special columns by changing the color of their font.| [("A", "0707C5"), ("E", "c10105")] |
|highlights | Change the color and/or font on chosen cells based on a condition you wish to highlight.| {"activate": True/False, "trigger": ["column_letter", "condition", value], "columns": [list of columns to highlight], "cell_color": "color_code", "font_size": #, "font_bold": True/False, "font_horiz_alignment": "alignment_type"} |
|draw_vert_line | Draw vertical lines with the given thickness on the right side of the specified columns. | {"columns": ["C", "H", "R"], "thickness": "medium"}|

### 1.2 Execute the analysis
After you have provided the required information in the files `inputs.py` and
`settings_messages.py`, proceed to execute the routine that generates
the Excel databases, either through Visual Studio Code, Jupyter Notebook, or an 
executable Windows application. 

To use the variants that require compiling the 
Python code when running the analysis, you must first set up a virtual 
environment and install the necessary [dependencies](dependencies/requirements.txt). 
The first part of the analysis can then be executed from the terminal with the 
command
```{script}
cd VScode
python main.py --inputs_file_path='../inputs.py' --settings_file_path='../settings_messages.py'

```

For detailed instructions in how to set up and use the different variants of the 
analysis, refer to:

**Visual Studio Code:** Navigate to the directory [VScode](VScode) to find 
detailed documentation on how to download the git repo, set up the virtual 
environment, install the dependencies, and execute the analysis in VScode. 

**Jupyter Notebook:** Navigate to the directory [JupyterNotebook](JupyterNotebook)
 to access the file you can use in the Jupyter Notebook web app. This analysis
  variant also requires downloading the code and setting up a virtual 
  environment, on which the notebook provides detailed instructions. 

**Graphic User Interface:** Navigate to the directory [GUI](GUI) to access the 
[executable file](GUI/slack2excel.exe) of the Windows visual interface. To use 
this application, download the executable file, run the app, and input the 
information prompted by the app. This is a standalone application, so you do 
not need to set up a virtual environment or manually specify your inputs into
 the files `inputs.py` or `settings_messages.py`. In the current version of the 
 visual interface, the user only provides the information listed in the INPUTS 
 table above. 




## 2. Excel database with compiled weekly reports

There is a dedicated Slack channel (named 'think-biver-weekly-checkins') for
users of the organization to post their weekly updates on the project(s) they 
are working on. However, and most generally, weekly reports can also be found 
in the dedicated Slack channels of the corresponding projects. 
The second part of this analysis takes the Excel files with the messages of
every Slack channel (previously generated in Section 1). It collects the messages fully
or partially parsed as weekly reports into a new Excel workbook. 
In addition, messages from the Slack channel "think-biver-weekly-reports" that
were not correctly parsed as weekly reports are also collected in a different
Excel sheet to facilitate their further inspection. The Excel workbook 
contains filters in the header columns to conveniently sort the information by
 channel, date, user name, etc.
 
### 2.1 Provide your inputs
You must first provide the information prompted in the file `settings_stats.py`.
Most of the variables are common to the variables in the SETTINGS_MESSAGES 
table above, except for:

|SETTINGS_STATS | Description |  Example/Format |
|---|---|---|
|jsons_source_path | Path where the exported JSON files are.| r"C:\Users\user_name\Documents\slackSource\January" |
|excel_channels_path | Path where the exported Excel files are. | r"C:\Users\user_name\Documents\slackSource\Excel" |
|reports_channel_name | Name of the Slack channel dedicated to the weekly reports. | "think-biver-weekly-checkins"|
|compilation_reports_file_name | Name of the file to be saved with the compilation of check-in weekly reports.| "compiled_weekly_reports_Jan.xlsx" |
|compilation_reports_path | Path where the file compilation_reports_file_path will be saved. | r"C:\Users\user_name\Documents\slackSource\Excel" |

### 2.2 Execute the analysis
After you have provided the required information in the file `settings_stats.py`, proceed to run the following command:
```{script}
cd VScode
python stats.py --settings_file_path="..\settings_stats.py"
```
Refer to the [VScode documentation](VScode) for detailed instructions.


