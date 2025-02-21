# Slack Workspace to Excel Databases
This work was developed to assist the internal functioning of an organization with more than 1400 users and 114 different Slack channels.
The goal was to collect, clean, and arrange all the information in the various Slack channels into Excel tables. Project managers can then use such databases to efficiently lead and guide the progress of the multiple projects being developed.
The code presented here can generate:
* Database of all the channels in the Slack workspace
* Database of all the users in the Slack workspace
* Database of all the messages in each Slack channel
  
In addition, it is customary for users of the Slack workspace to post their progress every week. Such weekly reports are parsed into categories and collected in the Excel databases. The default categories of the check-in messages are:
* Project Name
* Working On
* Progress
* Roadblocks
* Plans for the following week
* Meetings


## User's inputs
The conditions or inputs to the code are specified in the files `inputs.py` and `settings.py`. 

The `inputs.py` file contains the minimal information necessary to generate the Excel database of one or all Slack channels. 
|INPUTS | Description| Example/Format |
|---|---|---|
|slackexport_folder_path | Path to the source directory containing all the information exported from the Slack workspace. | r"C:\Users\user_name\Documents\slackSource\January" |
|converted_directory| Path to the destination directory where the analysis results will be saved. | r"C:\Users\user_name\Documents\slackSource\Excel" |
|chosen_channel_name | 'name-of-slack-channel', if analyzing only one Slack channel, or '' if analyzing all the Slack channels in the source directory. | "general" |
|write_all_channels_info | True/False to generate a file with information on all the Slack channels. | True |
|write_all_users_info | True/False to generate a file with the information of all the Slack users. | True |

The `settings.py` file contains variables that define the formatting of the check-in messages and the Excel tables:
|SETTINGS | Description |  Example/Format |
|---|---|---|
| timezone | Time zone to express the dates in the Excel tables. | "US/Central" |
| missing_value | String that identifies missing values in the Excel tables. | "n/d" |
|channels_json_name | Name of the JSON file exported from the Slack workspace containing the channels's information.| "channels.json" |
|users_json_name | Name of the JSON file exported from the Slack workspace containing the user's information. | "users.json" |
|dest_name_ext | Name of the folder created in the destination directory where all the files will be stored. | "_JSON_converted"|
|all_keywords | Keywords used to parse the check-in messages. | ["project_name", "working_on", "progress"] |
|index_keyword | Keyword used to identify each check-in message. | "project_name" |
|keywords_dictionary | Words/phrases that users are likely to use as keywords in their check-in messages.| {"working_on": ["working on", "worked on"]} |
|sample_text_list | Fragments of messages sent as samples showing the format of the weekly check-in messages. | ["please follow this structure when posting updates"] |
|column_widths | Set the width of the Excel columns.| {"col1_name": col1_width, "col2_name": col2_width} |
|text_type_cols | List the letters of the Excel columns containing long text to set their cell alignment. | ["A", "B", "C"] |
|height_1strow | Height of the header row in the Excel table. | 43 |
|alignment_vert_1strow | Vertical alignment of the header row in the Excel table. | "top" |
|alignment_horiz_1strow | Horizontal alignment of the header row in the Excel table. | "left" |
|font_size_1strow | Font size of the header row in the Excel table. | 9 |
|font_bold_1strow | Set the font of the header row in the Excel table to bold font. | True |
|cell_color_1strow | Cell color of the header row in the Excel table. | [("FFFFF", ["A", "B", "C"]), ("FFFFF", ["D", "E", "F"])] |
|font_color_in_column | Distinguish special columns by changing the color of their font.| [("A", "0707C5"), ("E", "c10105")] |
|highlights | Change the color and/or font on chosen cells based on a condition you wish to highlight.| {"activate": True/False, "trigger": ["column_letter", "condition", value], "columns": [list of columns to highlight], "cell_color": "color_code", "font_size": #, "font_bold": True/False, "font_horiz_alignment": "alignment_type"} |


## Executing the analysis

The options available for using this analysis to generate the Excel databases of your Slack workspace are using Visual Studio Code, Jupyter Notebook, or an executable Windows application. To use the variants that require compiling the Python code when running the analysis, it is necessary to set up a virtual environment and install the required [dependencies](dependencies/requirements.txt). 

**Visual Studio Code:** Navigate to the directory [VScode](VScode) to find detailed documentation on how to download the git repo, set up the virtual environment, install the dependencies, and execute the analysis in VScode. 

**Jupyter Notebook:** Navigate to the directory [JupyterNotebook](JupyterNotebook) to access the file you can use in your Jupyter Notebook web app. This analysis variant also requires downloading the code and setting up a virtual environment. The notebook provides detailed instructions on how to do so. 

**Graphic User Interface:** Navigate to the directory [GUI](GUI) to access the [executable file](GUI/slack2excel.exe) of the Windows visual interface. To use this application, download the executable file, run the app, and input the information prompted by the app. This is a standalone application, so you do not need to set up a virtual environment or manually specify your inputs into the files `inputs.py` or `settings.py`. In the current version of the visual interface, the user only provides the information listed in the INPUTS table above. 
