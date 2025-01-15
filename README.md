# Slack channels to Excel databases
This work was developed to assist the internal functioning of an organization with more than 1400 users and 114 different Slack channels.
The goal was to collect, clean, and arrange all the information in the various Slack channels into Excel databases that the project managers could use.
The results include:
* Database of all the channels in the Slack workspace
* Database of all the users in the Slack workspace
* Database of all the messages in each Slack channel
  
One of the Slack channels is used to register each user's weekly progress. These reports are parsed into the categories below and included in its corresponding database.
* Project Name
* Working On
* Progress
* Roadblocks
* Plans for the following week
* Meetings


## Input:
The conditions or inputs to the code can be specified in the file `inputs.py`.

|Variable | Purpose|
|---|---|
|slacexport_folder_path | Path to the source directory containing all the information exported from the Slack workspace |
|converted_directory| Path to the destination directory where the results of the analysis will be saved |
|chosen_channel_name | 'name-of-slack-channel', if analyzing only one Slack channel, or '' if analyzing all the Slack channels present in the source directory | 
| write_all_channels_info | True/False to generate a file with the information of all the Slack channels |
| write_all_users_info | True/False to generate a file with the information of all the Slack users |
| timezone | Choice of the time zone to express the dates in the Excel databases |
| missing_value | Choice of string to identify missing values in the Excel databases |
    
## Dependencies:
The dependencies used are specified in the file `requirements.txt`. You can create a virtual environment and install such dependencies by running the following command from the directory where you wish to create the virtual environment:
```{bash}
python3 -m venv ./venv
```
To activate the virtual environment _venv_ in Linux, run the command,
```{bash}
source venv/bin/activate
```
or the command below if working in Windows,
```{bash}
venv\Scripts\activate
```
The dependencies can then be installed with,
```{bash}
pip install -r requirements.txt
```

## Analysis:
Given the information supplied in the file `inputs.py`, you can execute the analysis with the following command,
```{bash}
python3 src/main.py
```
After completing the analysis, the virtual environment can be deactivated with,
```{bash}
deactivate
```

## Usage on Visual Studio Code:
Check the file `instructions_VScode.md` for detailed instructions on how to get and load the repository into Visual Studio Code, create and activate the virtual environment, and set up and run the analysis.
