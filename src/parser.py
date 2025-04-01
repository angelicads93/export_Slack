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

2. Initialize an object of the class "settings_mod" given as argument a Python
string with the absolute path to the txt file you wish to parse:
    inputs_parser = parser.Parser(path_to_txt_file)

3. Retrieve the value of the a variable from the txt file:
    var = inputs_parser.get("variable_name")
"""

import os


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
                if line.lstrip(' ')[0] != "#":

                    # Ignore tabs:
                    line = line.lstrip(' ')

                    # Ignore in-line comments:
                    line = line.split('#')[0]

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
