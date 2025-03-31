import os


class settings_mod:
    def __init__(self, txt_path):
        self.txt_path = txt_path
        self.file_name = os.path.basename(self.txt_path).split('.')[0]
        
        # --Extract the definitions of each variables and store them as one-line-string into the list variables_definitions:
        self.var_defs = self.vars_def_in_one_line()

        # --Build a dictionary with the variables names as keys and the Python evaluation as values:
        self.var_dict = self.evaluate_variables(self.var_defs)

        # --List the name of all the variables defined in txt_path:
        self.var_names = self.var_dict.keys()



    def vars_def_in_one_line(self):
        """ Each variable definition in the txt file is compressed to one line and store in the list "variables".
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

                    # Create a new entry on the list when it finds ' = ':
                    if ' = ' in line:
                        variables.append("")  # start a new entry in "variables"
                    
                    # Add the line to the corresponding item of the list variables:
                    variables[-1] += line.rstrip(' \n')

        return variables


    def evaluate_variables(self, variables):
        """ Evaluate the Python validity of each of the commands in the list variables.
        Stores in a Python dictionary the name of the variable and the evaluated Python value.
        Extra checks on strings and paths are made. Pending to add more as they come.
        """
        # --Initialize flag to control the final output of the function:
        flag = True

        # --Initialize dictionary where to store the variables and their Python values:
        var_dict = {}

        # --Iterate through all the variables to store them in the dictionary var_dict:
        for i in range(len(variables)):
            name = variables[i].split(' = ')[0]
            value = variables[i].split(' = ')[1]

            # --Substitute any previously defined variable into the current variable:
            for var in var_dict.keys():
                if var in value:
                    value = value.replace(var, str(var_dict[var]))

            # --Replace "\" with "\\" to avoid potential syntax errors in Windows:
            if "path" in name:
                value = value.replace('\\','\\\\')

            # --Replace single-quotation to double-quatation (for syntax uniformity):
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
                print(f'ERROR: Please review the definition of the variable {name} in the file {self.file_name}')
                break

        if flag is True:
            print(f'Configuration file "{self.file_name}" correctly parsed')
            return var_dict
        else:
            return None

    # -- Methods to retrieve the values of the various variables:
    
    def missing_value(self):
        if 'missing_value' in self.var_names:
            return self.var_dict['missing_value']
        else:
            return None
    
    def timezone(self):
        if 'timezone' in self.var_names:
            return self.var_dict['timezone']
        else:
            return None

    def jsons_source_path(self):
        if 'jsons_source_path' in self.var_names:
            return self.var_dict['jsons_source_path']
        else:
            return None
    
    def excel_channels_path(self):
        if 'excel_channels_path' in self.var_names:
            return self.var_dict['excel_channels_path']
        else:
            return None
    
    def reports_channel_name(self):
        if 'reports_channel_name' in self.var_names:
            return self.var_dict['reports_channel_name']
        else:
            return None
    
    def compilation_reports_file_name(self):
        if 'compilation_reports_file_name' in self.var_names:
            return self.var_dict['compilation_reports_file_name']
        else:
            return None
        
    def compilation_urls_file_name(self):
        if 'compilation_urls_file_name' in self.var_names:
            return self.var_dict['compilation_urls_file_name']
        else:
            return None
        
    def urls_to_show(self):
        if 'urls_to_show' in self.var_names:
            return self.var_dict['urls_to_show']
        else:
            return None
    
    def compilation_reports_path(self):
        if 'compilation_reports_path' in self.var_names:
            return self.var_dict['compilation_reports_path']
        else:
            return None
    
    def columns_order(self):
        if 'columns_order' in self.var_names:
            return self.var_dict['columns_order']
        else:
            return None
    
    def w_date(self):
        if 'w_date' in self.var_names:
            return self.var_dict['w_date']
        else:
            return None

    def w_name(self):
        if 'w_name' in self.var_names:
            return self.var_dict['w_name']
        else:
            return None

    def w_text(self):
        if 'w_text' in self.var_names:
            return self.var_dict['w_text']
        else:
            return None

    def w_bool(self):
        if 'w_bool' in self.var_names:
            return self.var_dict['w_bool']
        else:
            return None

    def w_count(self):
        if 'w_count' in self.var_names:
            return self.var_dict['w_count']
        else:
            return None

    def column_widths(self):
        if 'column_widths' in self.var_names:
            return self.var_dict['column_widths']
        else:
            return None

    def text_type_cols(self):
        if 'text_type_cols' in self.var_names:
            return self.var_dict['text_type_cols']
        else:
            return None

    def header_row(self):
        if 'header_row' in self.var_names:
            return self.var_dict['header_row']
        else:
            return None

    def font_color_in_column(self):
        if 'font_color_in_column' in self.var_names:
            return self.var_dict['font_color_in_column']
        else:
            return None

    def highlights(self):
        if 'highlights' in self.var_names:
            return self.var_dict['highlights']
        else:
            return None

    def draw_vert_line(self):
        if 'draw_vert_line' in self.var_names:
            return self.var_dict['draw_vert_line']
        else:
            return None
        
