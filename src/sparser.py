#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: Angelica Goncalves.

Module to parse the information provided by a user through a command in the
terminal and txt files.

Classes
-------
Parser(txt_path)

Functions
---------
set_flag_analyze_all_chs(chosen_channel_name)
    Return a boolean specifying if all the Slack channels must be analyzed.

check_missing_chs(expected_chs, present_chs, analyze_all_chs=False)
    Return a list with the names of the channels that are absent.

check_path_in_user_file(file_name, var_name, var_value, kill=True)
    Verify the validity of a path input by the user in a txt file.

check_file_exists(file_path, kill=False)
    Check if a file exists.

list_dirs_in_path(source_path)
    Extract the directories in source_path.

check_ch(file_name, source_path)
    Check if file_name is indeed an expected Slack channel.

make_dest_path(path)
    Create the path where all the files will be saved.
"""

# Import standard Python libraries:
import os
import sys
from pathlib import Path
import shutil
import argparse


class Parser:
    """
    Parse the validity of the information provided by a user in a txt file.

    The txt file is expected to be written with the Python syntax for comments,
    strings, lists, and dictionaries. If a syntax error is encountered when
    parsing the txt file, an error message is printed in the terminal, and the
    code terminates.

    To use parser.py:

    1. Import the module parser.py in your Python code:
    import parser

    2. Initialize an object of the class "settings" given as argument a Python
    string with the absolute path to the txt file you wish to parse:
    inputs_parser = parser.Parser(path_to_txt_file)

    3. Retrieve the value of a variable from the txt file:
    var = inputs_parser.get("variable_name")


    Attributes
    ----------
    txt_path : str
        Path to the txt file with the user's input information.

    Methods
    -------
    vars_def_in_one_line()
        Compress variable definitions to one continuous line.

    evaluate_variables(variables)
        Evaluate the Python validity of each of the commands in "variables".

    get(var_name)
        Retrieve the Python value of given var_name.
    """

    def __init__(self, txt_path):
        self.txt_path = txt_path
        self.file_name = os.path.basename(self.txt_path).split(".")[0]

        # Extract the definitions of each variable and store them as
        # one-line-string into the list variables_definitions:
        self.var_defs = self.vars_def_in_one_line()

        # Build a dictionary with the variables' names as keys and the Python
        # evaluation as values:
        self.var_dict = self.evaluate_variables(self.var_defs)

        # List the name of all the variables defined in txt_path:
        self.var_names = self.var_dict.keys()


    def vars_def_in_one_line(self):
        """
        Compress variable definitions to one continuous line.

        Returns
        -------
        List with complete Python commands read from the txt file.

        """
        variables = []

        with open(self.txt_path, "r", encoding="utf-8") as file:
            for line in file:
                # Ignore fully commented lines:
                if line.lstrip(" ")[0] != "#" and line.lstrip(" ")[0] != "\n":

                    # Ignore tabs:
                    line = line.lstrip(" ")

                    # Ignore in-line comments:
                    line = line.split("#")[0]

                    # Replace "'" with "\'" to avoid syntax errors later when
                    # using eval():
                    line = line.replace("'", "\\'")

                    # Create a new entry on the list when a new variable is
                    # starting to be defined:
                    if " = " in line:
                        variables.append("")

                    # Add the current line to the corresponding item of the
                    # list "variables":
                    variables[-1] += line.rstrip(" \n")

        return variables

    def evaluate_variables(self, variables):
        """
        Evaluate the Python validity of each of the commands in "variables".

        Extra checks on strings and paths are made. Pending adding more.

        Arguments
        ---------
        variables : list
            List with complete Python commands read from the txt file.

        Returns
        -------
        Python dictionary with the variable names and values.

        """
        # Initialize flag to control the final output of the function:
        flag = True

        # Initialize dictionary to store the variables and their values:
        var_dict = {}

        # Iterate through all the variables to store them in a dictionary:
        for i in range(len(variables)):
            name = variables[i].split(" = ")[0]
            value = variables[i].split(" = ")[1]

            # Substitute previously defined variables into current variable:
            for var in var_dict.keys():
                if var in value:
                    value = value.replace(var, str(var_dict[var]))

            # Replace "\" with "\\" to avoid syntax errors in Windows:
            if "path" in name:
                value = value.replace("\\", "\\\\")

            # Replace single-quotation to double-quatation
            # (for syntax uniformity):
            value = value.replace("'", '"')

            # Test the Python validity:
            try:
                value = eval(value)
                # Add variable and evaluated-value to the dictionary:
                if name not in var_dict.keys():
                    var_dict[name] = value
            # except Exception as e:
            except SyntaxError:
                flag = False
                # print(f'ERROR: {str(e).split("(")[0]}')
                print("ERROR: Please review the definition of the variable "
                      + f"{name} in the file {self.file_name}")
                break

        out = None
        if flag is True:
            print(f'Configuration file "{self.file_name}" correctly parsed')
            out = var_dict

        return out

    def get(self, var_name):
        """
        Retrieve the Python value of given var_name.

        If var_name was not found in the input txt file, then it returns None.

        Arguments
        ---------
        var_name : str
            Name of a variable in a txt file.

        Returns
        -------
        The value of the variable if found, or "None" otherwise.

        """
        out = None
        if var_name in self.var_names:
            out = self.var_dict[var_name]

        return out


def init_command_parser(description, args_names):
    
    # Create instance of the ArgumentParser class:
    argparser = argparse.ArgumentParser(description=description)
    
    # Add as many arguments as variables in args_names:
    for i in range(len(args_names)):
        argparser.add_argument("--"+args_names[i], required=True, type=str)
    
    return argparser.parse_args()
    
def set_flag_analyze_all_chs(chosen_channel_name):
    """
    Return a boolean specifying if all the Slack channels must be analyzed.

    The input variable "chosen_channel_name" is expected to be an empty
    string "" if all the Slack channels must be analyzed. Otherwise,
    it is expected to be a string with the name of the Slack channel as
    written in the source directory.

    Arguments
    ---------
    chosen_channel_name : str
        String with the user's input.

    Returns
    -------
    Boolean

    """
    if len(chosen_channel_name) < 1:
        analyze_all_channels = True
        print("Channel(s) to analyze: All")
    else:
        analyze_all_channels = False
        print("Channel(s) to analyze: ", chosen_channel_name)
    return analyze_all_channels


def check_missing_chs(expected_chs, present_chs, analyze_all_chs=False):
    """
    Return a list with the names of the channels that are absent.

    Arguments
    ---------
    expected_chs : list
        List the names of the Slack channels that are expected to be present.
    present_chs : list
        List with the names of the Slack channels present in the source dir.
    analyze_all_chs : bool
        Whether all the channels are to be analyzed.

    Returns
    -------
    List with the names of the expected Slack channels that are not present
    in the source directory.

    """
    missing_channels = []
    for ch in present_chs:
        if ch not in expected_chs:
            missing_channels.append(ch)

    # Print message with missing channel(s) in the terminal:
    if analyze_all_chs is True:
        if len(missing_channels) > 0:
            print("WARNING: The following channels are missing in the "
                  + "source directory:", missing_channels)

    return missing_channels


def check_path_in_user_file(file_name, var_name, var_value, kill=True):
    """
    Verify the validity of a path input by the user in a txt file.

    Arguments
    ---------
    file_name : str
        Name of the input file.

    var_name : str
        Name of the variable to be verified.

    var_value : str
        Value of var_name inputed by the user.

    kill : bool
        Whether to terminate the code if the file is not found.

    """
    flag = True
    if os.path.exists(var_value) is False:
        flag = False
        if kill is True:
            print(f"WARNING: Path {var_value} does not exists." + "\n"
                  + "          Please review your input for the variable"
                  + f" {var_name} in {file_name}.")
        else:
            print(f"ERROR: Path {var_value} does not exists." + "\n"
                  + "       Please review your input for the variable"
                  + f" {var_name} in {file_name}.")
            sys.exit()
    else:
        print(f'Verified that path "{var_value}" exists.')

    return flag


def check_file_exists(file_path, kill=False):
    """
    Check if a file exists.

    Arguments
    ---------
    file_path : str
        Path to the file.

    kill : bool
        Whether to terminate the code if the file is not found.

    """
    flag = True
    if os.path.exists(file_path) is False:
        flag = False
        f_name = os.path.basename(file_path)
        f_parent = os.path.dirname(file_path)
        if kill is True:
            print(f"WARNING: The file {f_name} was not found in {f_parent}.")
        else:
            print(f"ERROR: The file {f_name} was not found in {f_parent}.")
            sys.exit()
    else:
        print(f'Found file "{file_path}".')

    return flag


def list_dirs_in_path(source_path):
    """
    Extract the directories in source_path.

    Arguments
    ---------
    source_path : str
        Source directory to inspect.

    Returns
    -------
    List with the name of the directories inside source_path.

    """
    # Get all non-hidden directories and files in a path:
    lst_src = os.listdir(source_path)
    # Extract items that are indeed directories:
    chs_names = [lst_src[i]
                 for i in range(len(lst_src))
                 if os.path.isdir(f"{source_path}/{lst_src[i]}") is True]
    return chs_names


def check_ch(file_name, source_path):
    """
    Check if file_name is indeed an expected Slack channel.

    Notice that the default name of a channel's directory may contain empty
    spaces (" "), although the name of the corresponding Excel file does not.
    Empty spaces were replaced by underscores in the later.

    Arguments
    ---------
    file_name : str
        Name of the Slack channel.

    source_path : str
        Path to the source directory containing all the Slack JSON files.

    Returns
    -------
    Boolean. Whether file_name is present in source_path.

    """
    # Collect the name of all channel's subdirectories after removing all type
    # of spaces and separators from their names:
    chs_path = list_dirs_in_path(source_path)
    for i, ch in enumerate(chs_path):
        ch = ch.replace(" ", "").replace("-", "").replace("_", "")
        chs_path[i] = ch

    # Extract channel_name from the full name of the Excel file
    # ("<channel_name>_<start_date>_to_<end_date>.xlsx"):
    cn = "_".join(file_name.split("_")[:-3])

    # Remove all types of spaces and separators to compare the names further:
    cn = cn.replace(" ", "").replace("-", "").replace("_", "")

    # Compare the names:
    out = True
    if cn not in chs_path:
        out = False

    return out


def make_dest_path(path):
    """
    Create the path where all the files will be saved.

    If the path already exists, it deletes it and creates a fresh one.

    Arguments
    ---------
    path : str
        Path to be created.

    """
    # If the path already exists, remove it:
    if os.path.exists(path) is True:
        exprt_folder_path = Path(path)
        if exprt_folder_path.is_dir():
            print(f'WARNING: The path "{path}" already exists and it will '
                  + "be replaced.")
            shutil.rmtree(exprt_folder_path)
    else:
        print(f'Created destination directory "{path}".')

    # Create a fresh path:
    Path(f"{path}").mkdir(parents=True, exist_ok=True)
