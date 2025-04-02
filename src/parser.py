#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on 30/03/2025
@author: Angelica Goncalves

This Python module defines the routine to parse the user's input from a .txt.

The txt file is expected to be written with the Python syntax for comments,
strings, lists and dictionaries. If a syntax error is encountered when parsing
the txt file, an error message is printed in the terminal and the code
terminates.

To use parser.py:

1. Import the module parser.py in your Python code:
    import parser

2. Initialize an object of the class "settings" given as argument a Python
string with the absolute path to the txt file you wish to parse:
    inputs_parser = parser.Parser(path_to_txt_file)

3. Retrieve the value of the a variable from the txt file:
    var = inputs_parser.get("variable_name")
"""

# Import standard Python libraries:
import os
from pathlib import Path
import shutil
import pandas as pd
import re


class Parser:
    def __init__(self, txt_path):
        self.txt_path = txt_path
        self.file_name = os.path.basename(self.txt_path).split('.')[0]

        # --Extract the definitions of each variables and store them as
        # --one-line-string into the list variables_definitions:
        self.var_defs = self.vars_def_in_one_line()

        # --Build a dictionary with the variables names as keys and the Python
        # --evaluation as values:
        self.var_dict = self.evaluate_variables(self.var_defs)

        # --List the name of all the variables defined in txt_path:
        self.var_names = self.var_dict.keys()

    def vars_def_in_one_line(self):
        """ Each variable definition in the txt file is compressed to one line
        and store in the list "variables".
        Outputs the list "variables".
        """
        variables = []

        with open(self.txt_path, 'r') as file:
            for line in file:
                # Ignore fully commented lines:
                if line.lstrip(' ')[0] != "#" and line.lstrip(' ')[0] != "\n":

                    # Ignore tabs:
                    line = line.lstrip(' ')

                    # Ignore in-line comments:
                    line = line.split('#')[0]

                    # Replace "'" with "\'" to avoid syntax errors later when
                    # using eval():
                    line = line.replace("'", "\\'")

                    # Create a new entry on the list when a new variable is
                    # starting to be defined:
                    if ' = ' in line:
                        variables.append("")

                    # Add the current line to the corresponding item of the
                    # list "variables":
                    variables[-1] += line.rstrip(' \n')

        return variables

    def evaluate_variables(self, variables):
        """ Evaluate the Python validity of each of the commands in the list
        variables.
        Stores in a Python dictionary the name of the variable and the
        evaluated Python value.
        Extra checks on strings and paths are made. Pending to add more as
        they come.
        """
        # --Initialize flag to control the final output of the function:
        flag = True

        # --Initialize dictionary to store the variables and their values:
        var_dict = {}

        # --Iterate through all the variables to store them in a dictionary:
        for i in range(len(variables)):
            name = variables[i].split(' = ')[0]
            value = variables[i].split(' = ')[1]

            # --Substitute previously defined variables into current variable:
            for var in var_dict.keys():
                if var in value:
                    value = value.replace(var, str(var_dict[var]))

            # --Replace "\" with "\\" to avoid syntax errors in Windows:
            if "path" in name:
                value = value.replace('\\', '\\\\')

            # --Replace single-quotation to double-quatation
            # --(for syntax uniformity):
            value = value.replace("'", '"')

            # --Test the Python validity:
            try:
                value = eval(value)
                # --Add variable and evaluated-value to the dictionary:
                if name not in var_dict.keys():
                    var_dict[name] = value
            except Exception as e:
                flag = False
                #print(f'ERROR: {str(e).split("(")[0]}')
                print('ERROR: Please review the definition of the variable '
                      + f'{name} in the file {self.file_name}')
                break

        if flag is True:
            print(f'Configuration file "{self.file_name}" correctly parsed')
            return var_dict
        else:
            return None

    def get(self, var_name):
        """ Method to retrieve the Python value of the given var_name.
        If var_name was not found in the input txt file, then it return None.
        """
        if var_name in self.var_names:
            return self.var_dict[var_name]
        else:
            return None


class InspectSource:
    """
    Class to check the validity of the source directory given by the user.

    ...

    Attributes
    ----------
    inputs : parser.Parser(txt_path)
        The parsed user's inputs from the file inputs.txt.
        Variables defined in inputs.txt are retrieved as inputs.get(var_name).

    settings : parser.Parser(txt_path)
        The parsed user's inputs from the file settings_messages.txt
        Variables defined in settings_messages.txt are retrieved as
        settings.get(var_name).

    Methods
    -------
    update_cont_analysis(value)
        Update the value of the variable 'continue_analysis'.

    get_cont_analysis()
        Retrieve the value of the variable 'continue_analysis'.

    set_flag_analyze_all_chs()
        Return a boolean specifying if all the Slack channels must be analyzed

    check_src_path_exists()
        Checksthat the path to the source data exists.

    check_channels_json_exists()
        Check if the JSON files with the channels' information exist.

    check_users_json_exists()
        Check if the JSON files with the user's information exist.

    get_dest_path()
        Return the absolute path of the destination directory.

    make_dest_path()
        Create the path where all the files will be saved.

    get_chs_dir()
        Return list with name(s) of the Slack channel(s) to be converted.

    check_missing_chs()
        Return a list with the name of the channels that are absent.

    get_jsons_in_ch(list_names)
        Return a list with the names of a channels JSON files.

    get_jsons_in_all_chs()
        Return a list with the names of the JSON files of all the channels.

    """

    def __init__(self, inputs, settings):

        # Retrieve users inputs from inputs.txt:
        self.inputs = inputs
        self.chosen_channel_name = self.inputs.get('chosen_channel_name')
        self.slackexport_folder_path = self.inputs.get('slackexport_folder_path')
        self.converted_directory = self.inputs.get('converted_directory')

        # Retrieve users inputs from settings.txt:
        self.settings = settings
        self.dest_name_ext = self.settings.get('dest_name_ext')
        self.channels_json_name = self.settings.get('channels_json_name')
        self.users_json_name = self.settings.get('users_json_name')

        # Initialize flag to track validity of the inputs:
        self.update_cont_analysis(True)
        #
        # Perform some checks and retrieve some variables:
        self.dest_path = self.get_dest_path()
        self.flag_all_chs = self.set_flag_analyze_all_chs()
        self.present_chs = self.get_chs_dir()
        self.missing_chs = self.check_missing_chs()
        self.chs_jsons = self.get_jsons_in_all_chs()

    def update_cont_analysis(self, value):
        """
        Update the value of the variable 'continue_analysis'.

        Arguments
        ---------
        value : bool
            Boolean value (True/False).

        """
        self.continue_analysis = value

    def get_cont_analysis(self):
        """Retrieve the value of the variable 'continue_analysis'."""
        return self.continue_analysis

    def set_flag_analyze_all_chs(self):
        """
        Return a boolean specifying if all the Slack channels must be analyzed.

        The input variable "chosen_channel_name" is expected to be an empty
        string "" if all the Slack channels must be analyzed. Otherwise,
        it is expected to be a string with the name of the Slack channel as
        written in the source directory.

        """
        if len(self.chosen_channel_name) < 1:
            analyze_all_channels = True
            print('Channel(s) to analyze: All')
        else:
            analyze_all_channels = False
            print('Channel(s) to analyze: ', self.chosen_channel_name)
        return analyze_all_channels

    def check_src_path_exists(self):
        """Verify that the path to the source data exists."""
        if os.path.exists(self.slackexport_folder_path) is False:
            print('Please enter a valid path to the source directory')
            self.update_cont_analysis(False)

    def check_channels_json_exists(self):
        """Check if the JSON files with the channels information exist."""
        if os.path.exists(
            f"{self.slackexport_folder_path}/{self.channels_json_name}"
        ) is False:
            print(f'File {self.channels_json_name} was not found in the '
                  + 'source directory')
            self.update_cont_analysis(False)

    def check_users_json_exists(self):
        """Check if the JSON files with the channels information exist."""
        if os.path.exists(
            f"{self.slackexport_folder_path}/{self.users_json_name}"
        ) is False:
            print(f'File "{self.users_json_name}" was not found in the '
                  + 'source directory')
            self.update_cont_analysis(False)

    def get_dest_path(self):
        """Return the absolute path of the destination directory."""
        return f"{self.converted_directory}/{self.dest_name_ext}"

    def make_dest_path(self):
        """
        Create the path where all the files will be saved.

        If the path already exists, it deletes it and creates a fresh one.

        """
        # Get the absolute path where all the files will be saved:
        path = self.get_dest_path()

        # If the path already exists, remove it:
        if os.path.exists(path) is True:
            exprt_folder_path = Path(path)
            if exprt_folder_path.is_dir():
                print(f"WARNING: The path '{path} already exists and it will "
                      + "be replaced.")
                shutil.rmtree(exprt_folder_path)

        # Create a fresh path:
        Path(f"{path}").mkdir(parents=True, exist_ok=True)

    def get_chs_dir(self):
        """
        Return list with name(s) of the Slack channel(s) to be converted.

        If analysing one channel, checks that its directory exists, and default
        to the 0-th element of channels_names:
        channels_names = [ chosen_channel_name ] for one channel
        channels_names = [channel0, channel1, ...] for all the channels

        """
        if self.flag_all_chs is False:
            # Check that chosen_channel_name is correct and add to list:
            if os.path.exists(
                f"{self.slackexport_folder_path}/{self.chosen_channel_name}"
            ) is False:
                chs_names = []
                print("ERROR: The source directory for the channel "
                      + f"'{self.chosen_channel_name}' was not found in "
                      + f"{self.slackexport_folder_path}")
                self.update_cont_analysis(False)
            else:
                chs_names = [self.chosen_channel_name]
        else:
            # Add all directories in the source path to the list:
            lst_src = os.listdir(self.slackexport_folder_path)
            chs_names = [lst_src[i]
                         for i in range(len(lst_src))
                         if os.path.isdir(f"{self.slackexport_folder_path}/{lst_src[i]}") is True]

        return chs_names

    def check_missing_chs(self):
        """Return a list with the name of the channels that are absent."""
        # Get names of channels in channels.json:
        expected_chs_names = pd.read_json(
            f"{self.slackexport_folder_path}/{self.channels_json_name}"
            )['name'].values

        # Get names of expected channels that are not in the source directory:
        missing_channels = []
        for ch in self.present_chs:
            if ch not in expected_chs_names:
                missing_channels.append(ch)

        # Print message with missing channel(s) in the terminal:
        if self.flag_all_chs is True:
            if missing_channels is not []:
                print("WARNING: The following channels are missing in the "
                      + "source directory:", missing_channels)

        return missing_channels

    def get_jsons_in_ch(self, ch_files):
        """
        Return a list with the names of a channels JSON files.

        Arguments
        ---------
        ch_files : list
            List with the name of all the files inside the directory of a given
            Slack channel.

        """
        list_names_dates = []
        for i in range(len(ch_files)):
            match = re.match(
                r'(\d{4})(-)(\d{2})(-)(\d{2})(.)(json)', ch_files[i]
                )
            if match is not None:
                list_names_dates.append(ch_files[i])
        return list_names_dates

    def get_jsons_in_all_chs(self):
        """
        Return a list with the names of the JSON files of all the channels.

        The name of the JSON files should the correct format 'yyyy-mm-dd.json'.

        all_channels_jsonFiles_dates = [
            [chosen_channel_name_json0, chosen_channel_name_json1, ...]
            ] for one exportchannel
        all_channels_jsonFiles_dates = [
            [channel0_json0, channel0_json1, ...],
            [channel1_json0, channel1_json1, ...], ...
            ] for all the channels
        """
        all_channels_jsonFiles_dates = []
        for ch in self.present_chs:
            channel_jsonFiles_dates = self.get_jsons_in_ch(
                os.listdir(f"{self.slackexport_folder_path}/{ch}"))
            all_channels_jsonFiles_dates.append(channel_jsonFiles_dates)
        return all_channels_jsonFiles_dates
