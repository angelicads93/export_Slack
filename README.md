# Slack Workspace to Excel Databases
This work was developed to assist the internal functioning of an organization 
with more than 114 different Slack channels. The goal was to collect, clean,
and arrange all the information in the Slack workspace into Excel tables that
project managers can then use to efficiently lead and guide the progress of 
the multiple projects being developed. 

Furthermore, it is customary for users of the Slack workspace to post their 
progress every week. Such weekly reports are parsed into categories and 
collected in the Excel databases.



## Repository Content

### 1. VScode

Contains the main Python scripts to perform the various analyses:
* [extract_messages.py](VScode/extract_messages.py):
  Python script to export all the Slack messages into Excel workbooks. 
* [extract_weekly_reports.py](VScode/extract_weekly_reports.py):
  Python script to export all the weekly reports of the Slack users into an Excel workbook.
* [extract_urls.py](VScode/extract_urls.py):
  Python script to export all the messages containing URL links into an Excel workbook.
  
These Python scripts can be executed through a command line interface or an integrated development environment (IDE), like VScode
for Windows users. A detailed documentation on how to use this repository in VScode is also provided.

### 2. User's input files

Contains the various text files necessary to run the multiple analyses. They are ".txt" files, however, their content
should be written with Python syntax and conventions, especially for variable definitions, strings, lists and
dictionaries. Each file is first examined to check that each variable definition is a correct Python command, as
well as checking that any path or directory exists.
* [inputs.txt](inputs.txt) and [settings_messages.txt](settings_messages.txt): Input text files for the Python script 'extract_messages.py'.
* [settings_weekly_reports.txt](settings_weekly_reports.txt): Input text file for the Python script `extract_weekly_reports.py`.
* [settings_urls.txt](settings_urls.txt): Input text file for the Python script `extract_urls.py`.

All the input files are located in the main repository directory (`export_Slack/`), such that they are equally accessible
to the various ways of executing the analyses (VScode, JupyterNotebook or GUI). More details on the variables defined in each of these
input files are provided later.

### 2. src/

Contains the custom Python modules built to target the tasks necessary to complete each analysis. 
* [sparser.py](src/sparser.py):
  Python module used to parse the information provided by the user in both the command-line and the txt files
  (`inputs.txt`, `settings_messages.txt`, `settings_weekly_reports.txt`, `settings_urls.txt`).
* [slack.py](src/slack.py):
  Python module used to extract, process, and format information from the Slack workspace. The functions defined here
  are used across the three main Python scripts `extract_messages.py`, `extract_weekly_reports.py` and
  `extract_urls.py`.
* [checkins.py](src/checkins.py):
  Python module used to parse the Slack messages for weekly reports.
* [clean.py](src/clean.py):
  Python module used to clear the format of Pandas dataframes. 
* [excel.py](src/excel.py):
  Python module used to format the Excel workbooks.

These modules are imported at the beginning of the main Python scripts (`extract_messages.py`, `extract_weekly_reports.py` and `extract_urls.py`)
and their classes and functions are accessed through the module's namespace. For example,
```{python}
import checkins
df_msgs_parsed = checkins.parse_reports(df_msgs, setts)
```
To inspect/debug a line of code on one of the main Python scripts, first identify the custom Python module containing the function, and then
inspect the definition of the function in the Python module (for example, on `checkins.py`). Each class and function in the custom Python
modules has its docstring documentation, containing the description of the function, parameters and what does the function returns.
Such documentation can be conveniently accessed by hovering over the function's name from your VScode interface. 

### 3. JupyterNotebook/
 Contains the Jupyter Notebook [main.ipynb](JupyterNotebook/main.ipynb) that performs the analysis of exporting all the Slack messages into Excel workbooks (it is
 the Jupyter notebook equivalent of `extract_messages.py`). 
 This variant requires downloading the code and setting up a virtual environment, on which the notebook provides detailed instructions.
 In the present version of export_Slack, the Jupyter Notebook variant does not support the generation of Excel workbooks with the compilation of the weekly reports or the URLs found in the messages.

### 4. GUI/
* [GUI_tkinter.py](GUI/GUI_tkinter.py):
  Python script to export all the Slack messages into Excel workbooks through a graphical user interface (GUI).
  The visual interface prompts the user for:
  * The source path containing the JSON files exported from the Slack workspace.
  * The destination path where the outputs of the analysis will be saved.
  * The name of the channel(s) to be analyze.
  * Whether to generate Excel workbooks with information on all the channels and all the users in the Slack workspace.

  Notice that this information is precisely what is contained in the `inputs.txt` file when executing the `extract_messages.py` through VScode.
  The Python script `GUI_tkinter.py` then generates the corresponding text file with the information provided by the user through the visual interface and
  proceeds to execute the main routine. The variables of the text file `settings_messages.txt` are not prompted
  to the user of the visual interface, but instead, the values used are their default ones. This is the main reason why `inputs.txt` and `settings_messages.txt`
  are two different files. In the present version of export_Slack, the visual interface does not support the generation of Excel workbooks
  with the compilation of the weekly reports or the URLs found in the messages. 

  The Python script `GUI_tkinter.py` can be used to generate a standalone executable file suitable for Windows operating system.
  To generate the executable file, first ensure that the virtual environment is activated, then run the following command:
  ```{script}
  cd GUI
  pyinstaller --one-file --windowed GUI_tkinter.py
  ```
  The option `--one-file` bundles all the dependencies into a single executable file. And the option `--windowed` prevents a terminal window
  from popping when running the executable file. The output of this command is the directory `tkGUI/`.
  
* [tkGUI](GUI/tkGUI)
  Directory created from running the `pyinstaller` command above.
  
* [slack2excel.exe](GUI/slack2excel.exe):
  Executable file of the Graphical User Interface. It is a renamed copy of the file `tkGUI/dist/RET_GUI_tkinter.exe`.

  To use this application, download the executable file, run the app, and input the information prompted by the app. This is a standalone application, so you do
  not need to set up a virtual environment or manually specify your inputs in the files `inputs.txt` or `settings_messages.txt` to run the app.

### 5. dependencies/
* [requirements.txt](dependencies/requirements.txt):
  Text file containing the list of required Python libraries. This file is used to install the specified libraries in a virtual environment.

  To create a Windows virtual environment for the first time, run the following commands in your terminal:
  ```{script}
  python -m venv venv
  .\venv\Script\activate
  pip install -r dependencies\requirements.txt
  ```

  To deactivate your virtual environment once you have finished working on it, run the command:
  ```{script}
  deactivate
  ```

  If you created the virtual environment and installed the dependencies in a previous session, and you want to reuse it, you just need to
  reactivated with
  ```{script}
  .\venv\Script\activate
  ```


  
* [setup_venv_windows.bat](dependencies/setup_venv_windows.bat) and [setup_venv_linux.sh](dependencies/setup_venv_linux.sh):
  Windows and Linux files with the sequence of the command-line instructions necessary to create and activate a virtual environment and to
  install the required Python libraries from `requirements.txt`. If used, they should be executed from the main repository path `export_Slack/`.

### 6. images/
 * List of pictures used on the VScode documentation.


   
## Analyses

### 1. Excel database of messages in the Slack workspace
The messages on each Slack channel are exported into a unique Excel workbook.
The name of each Excel file consists of the name of the corresponding channel,
the start date, and the end date of the exported messages: 
<channel_name>\_<YYYY_MM_DD>\_to\_<YYYY_MM_DD>.xlsx

The messages of the Slack channels are pruned based on their value and relevance 
into three worksheets or tabs: "Relevant messages", "Filtered-out messages", 
and "All messages". The "relevant messages" do not contain messages sent by a
bot user, automatic messages sent after an action has been taken on the Slack 
channel (changed on name, if a user joined or left the channel, etc), nor short 
messages usually associated with short greetings and polite replies (e.g. "Hi 
everyone", "Ok", "Thank you"). Emojis were also removed (e.g., ":simple_smile:").
In addition to the Excel workbooks of each Slack channel, it
can be chosen to generate an Excel workbook with the compiled information of 
all the channels and all the users in the Slack workspace. 

The information of each Slack channel is arranged into the columns:

* msg_id: Identification code of each message (e.g., 7ca6388d-a828-4614-bc60-4b3661fa2336).
* msg_date: Date when the message was sent (e.g., 2025-01-25 18:13:46).
* user: Unique user identification code (e.g., U02063W7Z1V).
* name: User's name.
* display_name: User's name as displayed in the Slack profile.
* deactivated: Whether the user is active or not in the Slack workspace (e.g., True, False).
* is_bot: Whether the user is a bot (e.g., True, False).
* type: Type of correspondence, either a new message or a conversation thread (e.g., message, thread).
* text: User's message.
* contained_emoji: Whether the user's message contained any emojis (e.g., True, False).
* reply_count: Number of replies on a message thread.
* reply_users_count: 
* latest_reply_date: Date of the latest reply to the message thread (e.g., 2025-01-20 16:00:00).
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

#### 1.1 Provide your inputs
The first step is to provide the information requested in the
files `inputs.txt` and `settings_messages.txt`. The `inputs.txt` file contains the
 minimal information necessary to generate the Excel database of the Slack 
 channel(s) of choice. 
|INPUTS | Description| Example/Format |
|---|---|---|
|slackexport_folder_path | Path to the source directory containing all the information exported from the Slack workspace. | r"C:\Users\user_name\Documents\slackSource\January" |
|converted_directory| Path to the destination directory where the analysis results will be saved. | r"C:\Users\user_name\Documents\slackSource\Excel" |
|chosen_channel_name | 'name-of-slack-channel', if analyzing only one Slack channel, or '' if analyzing all the Slack channels in the source directory. | "general" |
|write_all_channels_info | True/False to generate a file with information on all the Slack channels. | True |
|write_all_users_info | True/False to generate a file with the information of all the Slack users. | True |

The `settings_messages.txt` file contains variables necessary for formatting the 
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

#### 1.2 Execute the analysis
After you have provided the required information in the files `inputs.txt` and
`settings_messages.txt`, proceed to execute the routine that generates
the Excel databases, either through the command-line interface with
```{script}
cd VScode
python extract_messages.py --inputs_file_path='../inputs.txt' --settings_file_path='../settings_messages.txt'

```
or through Visual Studio Code ([see documentation](VScode)), Jupyter Notebook ([main.ipynb](JupyterNotebook/main.ipynb)), or
the executable Windows application ([slac2excel.exe](GUI/slack2excel.exe)).

### 2. Excel database with compiled weekly reports

There is a dedicated Slack channel (named 'think-biver-weekly-checkins') for
users of the organization to post their weekly updates on the project(s) they 
are working on. However, and most generally, weekly reports can also be found 
in the dedicated Slack channels of the various projects. 
The second part of this analysis takes the Excel files with the messages of
every Slack channel (previously generated in Section 1), and it collects into a 
new Excel workbook the messages that were fully or partially parsed as weekly reports. 
In addition, messages from the Slack channel "think-biver-weekly-reports" that
were not correctly parsed as weekly reports are also collected in a different
Excel sheet to facilitate their further inspection. The Excel workbook 
contains filters in the header columns to conveniently sort the information by
 channel, date, user name, etc.
 
#### 2.1 Provide your inputs
You must first provide the information prompted in the file `settings_weekly_reports.txt`.
Most of the variables are common to the variables in the SETTINGS_MESSAGES 
table above, except for:

|SETTINGS_STATS | Description |  Example/Format |
|---|---|---|
|jsons_source_path | Path where the exported JSON files are.| r"C:\Users\user_name\Documents\slackSource\January" |
|excel_channels_path | Path where the exported Excel files are. | r"C:\Users\user_name\Documents\slackSource\Excel" |
|reports_channel_name | Name of the Slack channel dedicated to the weekly reports. | "think-biver-weekly-checkins"|
|compilation_reports_file_name | Name of the file to be saved containing the compilation of the weekly reports.| "compiled_weekly_reports_Jan.xlsx" |
|compilation_reports_path | Path where the file compilation_reports_file_path will be saved. | r"C:\Users\user_name\Documents\slackSource\Excel" |

#### 2.2 Execute the analysis
After you have provided the required information in the file `settings_weekly_reports.txt`, proceed to run the following command:
```{script}
cd VScode
python extract_weekly_reports.py --settings_file_path="..\settings_weekly_reports.txt"
```
or refer to the [VScode documentation](VScode) for detailed instructions.


### 3. Excel database with Slack messages containing specific URL(s)

Finally, the executable Python script `extract_urls.py` can be used to generate an Excel workbook
with tabs "All URLs" and "Selected URLs". The former contains all the messages on the Slack
workspace that explicitly had a URL link in their text. The cases containing URLs of interest,
as specified in the variable "urls_to_show", are collected in the "Selected URLs" tab of the Excel
workbook. The columns displayed are:
* display_name: User's name as displayed in the Slack profile.
* URL(s): Lists any links present in the user's message.
* msg_date: Date when the message was sent (e.g. 2025-01-25 18:13:46).
* channel: Name of the Slack channel.

#### 3.1 Provide your inputs
You must first provide the information prompted in the file `settings_urls.txt`.
Most of the variables are common to the variables in the SETTINGS_MESSAGES 
table above, except for:

|SETTINGS_STATS | Description |  Example/Format |
|---|---|---|
|jsons_source_path | Path where the exported JSON files are.| r"C:\Users\user_name\Documents\slackSource\January" |
|excel_channels_path | Path where the exported Excel files are. | r"C:\Users\user_name\Documents\slackSource\Excel" |
|reports_channel_name | Name of the Slack channel dedicated to the weekly reports. | "think-biver-weekly-checkins"|
|urls_to_show | List with the name of URLs of interest. | ["docs.google.com", "figma.com"]|
|compilation_urls_file_name | Name of the file to be saved containing the compilation of the URL links.| "compiled_urls_Jan.xlsx" |
|compilation_urls_path | Path where the file compilation_urls_file_path will be saved. | r"C:\Users\user_name\Documents\slackSource\Excel" |

#### 3.2 Execute the analysis
After you have provided the required information in the file `settings_urls.txt`, proceed to run the following command:
```{script}
cd VScode
python extract_urls.py --settings_file_path="..\settings_urls.txt"
```
or refer to the [VScode documentation](VScode) for detailed instructions.

