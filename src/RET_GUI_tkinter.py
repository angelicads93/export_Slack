#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Oct 27 19:36:24 2024

@author: agds
"""

import tkinter as tk
import os
from os.path import exists
import messages
from settings import dest_name_ext


class GUI(tk.Tk):

    def __init__(self):
        super().__init__()

        # --Define text style to be used:
        self.font_choice = ('Helvetica', 9)

        # --Defining some colors to use later for the interface:
        self.dark_gray = "#464C54"
        self.blue_gray = "#6B808E"
        self.ice_blue = "#D7DDE1"
        self.gray = "#CDCDCF"
        self.light_gray = "#E7E6E6"
        self.gold = '#FBC000'

        self.dark1 = "#1F1F1F"
        self.dark2 = "#2E2E2E"
        self.dark3 = "#4D4D4D"

        self.darkA = "#202124"
        self.darkB = "#29292D"
        self.darkC = "#686868"
        self.darkD = "#9BA0A6"
        self.white = "#FFFFFF"

        self.bkgc = self.darkA
        self.boxc = self.darkB
        self.letterc = self.white
        self.titlec = self.gold
        self.bannerc = self.darkA

        # --Initialize a default inputs.py file:
        self.build_input_file('channel', 'path_orig', 'path_dest',
                              False, False)

        # --Title and size of the GUI:
        self.wm_title = "Slack channels to Excel Databases"
        self.geometry('600x500')
        self.configure(bg=self.darkA)
        #self.resizable(0, 0)

        # --Header:
        self.frame = tk.Frame(master=self, bg=self.bkgc)

        self.label = tk.Label(
            master=self.frame, text='Slack channels to Excel Databases',
            font=('Roboto', 20),
            width=300, bg=self.bannerc
            )
        self.label.configure(fg=self.titlec)

        # --Label for the source path:
        self.labelOrig = tk.Label(
            master=self.frame, text='Path to source directory:',
            font=self.font_choice, bg=self.bkgc
            )
        self.labelOrig.configure(fg=self.letterc)

        # --Entry box for the source path:
        self.entryOrig = tk.Entry(
            master=self.frame,
            font=self.font_choice,
            width=600, bg=self.boxc, border=0, disabledbackground=self.boxc,
            highlightthickness=0
            )
        self.entryOrig.configure(fg=self.letterc)

        # --Label for the destination path:
        self.labelDest = tk.Label(
            master=self.frame, text='Save in path:',
            font=self.font_choice, bg=self.bkgc
            )
        self.labelDest.configure(fg=self.letterc)

        # --Entry box for the destination path:
        self.entryDest = tk.Entry(
            master=self.frame,
            font=self.font_choice,
            width=600, bg=self.boxc, border=0, disabledbackground=self.bkgc,
            highlightthickness=0, readonlybackground=self.boxc
            )
        self.entryDest.configure(fg=self.letterc)

        # --Label for the channel name:
        self.labelChannel = tk.Label(
            master=self.frame, text='Slack channel:',
            font=self.font_choice, bg=self.bkgc
            )
        self.labelChannel.configure(fg=self.letterc)

        # --Option Menu for the available channels in the source path:
        self.default_str_channel = '  '   # Default value
        self.channel_var = tk.StringVar(self.frame)
        self.channel_var.set(self.default_str_channel)
        self.Channel = tk.OptionMenu(
            self.frame,
            self.channel_var,
            value=[self.default_str_channel]
            )
        self.Channel.configure(
            font=self.font_choice,
            width=600, bg=self.boxc, border=0,
            highlightthickness=0,
            fg=self.letterc,
            state='disabled')

        # --Checkbox for channels_flag:
        self.channels_flag_var = tk.BooleanVar()
        self.channels_flag_var.set(False)
        self.checkbox_channels = tk.Checkbutton(
            master=self.frame, variable=self.channels_flag_var,
            text=' Export general information of all the Slack channels',
            font=self.font_choice, fg=self.letterc,
            bg=self.bkgc, disabledforeground=self.bkgc,
            bd=0, highlightthickness=0,
            selectcolor=self.bkgc,
            activebackground=self.bkgc, activeforeground=self.letterc,
            onvalue=True, offvalue=False, offrelief='flat'
            )

        # --Checkbox for users_flag:
        self.users_flag_var = tk.BooleanVar()
        self.users_flag_var.set(False)
        self.checkbox_users = tk.Checkbutton(
            master=self.frame, variable=self.users_flag_var,
            text=' Export general information of all the Slack users',
            font=self.font_choice, fg=self.letterc,
            bg=self.bkgc, disabledforeground=self.bkgc,
            bd=0, highlightthickness=0,
            selectcolor=self.bkgc,
            activebackground=self.bkgc, activeforeground=self.letterc,
            onvalue=True, offvalue=False
            )

        # --Label for showing error/update messages of the status of the code:
        self.labelError = tk.Label(
            master=self.frame, text=' ',
            font=self.font_choice, bg=self.bkgc,
            width=600, wraplength=500, height=5
            )
        self.labelError.configure(fg=self.letterc)

        # --Buttons
        self.frameButtons = tk.Frame(self.frame)
        self.buttonOK = tk.Button(
            self.frame, text='Update', command=self.check_inputs_exists,
            font=self.font_choice,
            border=0, highlightcolor=self.darkB, highlightbackground=self.bkgc
            )
        self.buttonOK.configure(bg=self.darkB, fg=self.letterc)
        self.buttonContinue = tk.Button(
            self.frame, text='Download', command=self.check_inputs_exists,
            font=self.font_choice,
            border=0, highlightcolor=self.darkB, highlightbackground=self.bkgc
            )
        self.buttonContinue.configure(bg=self.darkB, fg=self.letterc, state='disabled')

        # --Pack widgets in desired order:
        pad_x = 15
        self.frame.pack(padx=pad_x, pady=10, fill='both', expand=True, side='top')
        self.label.pack(padx=pad_x, pady=(15,25))
        self.labelOrig.pack(padx=pad_x, pady=0, after=self.label, anchor='w')
        self.entryOrig.pack(padx=pad_x, pady=(0,15), after=self.labelOrig, anchor='w')
        self.labelDest.pack(padx=pad_x, pady=0, anchor='w')
        self.entryDest.pack(padx=pad_x, pady=(0,15), after=self.labelDest, anchor='w')
        self.labelChannel.pack(padx=pad_x, pady=0, anchor='w')
        self.Channel.pack(padx=pad_x, after=self.labelChannel, anchor='w')
        self.checkbox_channels.pack(padx=pad_x, pady=(25,0), after=self.Channel, anchor='w')
        self.checkbox_users.pack(padx=pad_x, pady=5, after=self.checkbox_channels, anchor='w')
        self.labelError.pack(padx=pad_x, pady=10, after=self.checkbox_users, anchor='s')
        self.buttonOK.pack(padx=5, pady=(10,15), after=self.labelError, side='right', anchor='s')
        self.buttonContinue.pack(padx=5, pady=(10,15), after=self.labelError, side='right', anchor='s')

        self.mainloop()

    def build_input_file(self, channel, path_orig, path_dest,
                         channels_flag, users_flag):
        cwd = os.getcwd()
        self.input_file_path = f'{cwd}/..'
        f = open(f'{self.input_file_path}/inputs.py', 'w')
        f.write("# --If you wish to analyze one Slack channel, enter the channel name." + '\n')
        f.write("# --If you wish to analyze all the Slack channels, enter ''." + '\n')
        f.write(f"chosen_channel_name = '{channel}'" + '\n\n')
        f.write('# --Do you wish to generate the file with the information of all the' + '\n')
        f.write('# --Slack channels?:' + '\n')
        f.write(f"write_all_channels_info = {channels_flag}" + '\n\n')
        f.write('# --Do you wish to generate the file with the information of all the' + '\n')
        f.write('# --Slack users?:' + '\n')
        f.write(f"write_all_users_info = {users_flag}" + '\n\n')
        f.write('# --Insert path where the LOCAL copy of the GoogleDrive folder is:' + '\n')
        f.write(f"slackexport_folder_path = '{path_orig}'" + '\n\n')
        f.write('# --Insert path where the converted files will be saved:' + '\n')
        f.write(f"converted_directory = '{path_dest}'" + '\n\n')
        f.close()

    def check_inputs_exists(self):
        """ Verifies the validity of the input. Execute analysis if it's ok"""
        # -- Retrieve user's input:
        self.path_orig = str(self.entryOrig.get()).replace('"', '')
        self.path_dest = str(self.entryDest.get()).replace('"', '')
        self.channel_var_get = str(self.channel_var.get())
        self.channels_flag = str(self.channels_flag_var.get())
        self.users_flag = str(self.users_flag_var.get())
        print(f"AG: channels_flag = {str(self.channels_flag_var.get())}")
        print(f"AG: users_flag = {str(self.users_flag_var.get())}")

        # --1. Verify source path:
        if self.path_orig == '':
            self.txt = 'Please enter the path to the source directory'
            self.labelError.configure(text=self.txt)
            print(self.txt)
        elif self.path_orig != '' and exists(self.path_orig) is False:
            self.txt = 'Please enter a valid path to the source directory'
            self.labelError.configure(text=self.txt)
            print(self.txt)
        elif self.path_orig != '' and exists(self.path_orig) is True:
            self.entryDest.configure(state='normal')

            # --2. Verify destination path:
            if self.path_dest == '':
                self.txt = 'Please enter the destination path'
                self.labelError.configure(text=self.txt)
                print(self.txt)
            elif self.path_dest != '' and exists(self.path_dest) is False:
                self.txt = 'Please enter a valid destination path'
                self.labelError.configure(text=self.txt)
                print(self.txt)
            elif self.path_dest != '' and exists(self.path_dest) is True:
                if os.access(self.path_dest, os.X_OK) is False:
                    print('AG: Cannot write in the destinatin directory')
                else:
                    # --3. Verify channel:

                    # --Get names of channels in the source directory:
                    inspect_source = messages.InspectSource(
                        '', self.path_orig, self.path_dest
                        )
                    channels_names = inspect_source.get_channels_names()
                    channels_names.insert(0, 'All channels')
                    if dest_name_ext in channels_names:
                        channels_names.remove(dest_name_ext)
                    print('AG: CHANNELS NAMES = ', channels_names)

                    # --Activate menu with channels_name as choices:
                    if self.channel_var_get == self.default_str_channel:
                        self.Channel.destroy()
                        self.channel_var = tk.StringVar()
                        self.channel_var.set(self.default_str_channel)
                        self.Channel = tk.OptionMenu(
                            self.frame,
                            self.channel_var,
                            *channels_names
                            )
                        self.Channel.configure(
                            font=self.font_choice, fg=self.letterc,
                            width=600, bg=self.boxc,
                            border=0, highlightthickness=0, borderwidth=0,
                            activebackground=self.boxc,
                            activeforeground=self.letterc,
                            justify='left', direction='below'
                            )
                        self.Channel['menu'].configure(
                            bg=self.boxc, fg=self.letterc,
                            activebackground='gray',
                            activeforeground=self.letterc,
                            border=0, borderwidth=0)
                        self.Channel.pack(padx=15, pady=0, after=self.labelChannel)

                        self.txt = 'Please enter a Slack channel'
                        self.labelError.configure(text=self.txt)
                        print(self.txt)

                    else:
                        print('All input information provided')
                        print('Checking consistency')
                        self.txt = ' '
                        self.labelError.configure(text=self.txt)
                        self.buttonOK.configure(state='disabled')
                        self.buttonContinue.configure(state='normal', command=self.setup_inputs)


    def setup_inputs(self):
        """ Inspects the directories and creates the inputs.py file """
        inspect_source = messages.InspectSource(
            self.channel_var_get,
            self.path_orig, self.path_dest
            )
        save_in_path = inspect_source.save_in_path()
        inspect_source.check_save_path_exists(save_in_path)
        inspect_source.check_expected_files_exists()
        channels_names = inspect_source.get_channels_names()

        # --Generate inputs.py file:
        if self.channel_var_get == 'All channels':
            self.channel_var_get = ''
        self.build_input_file(
            self.channel_var_get,
            self.path_orig, self.path_dest,
            self.channels_flag, self.users_flag
        )
        
        # --Update visual interface label and button.
        self.Channel.configure(state='disable')
        self.entryOrig.configure(state='disable')
        self.entryDest.configure(state='disable')
        self.buttonContinue.configure(state='disable')

        # --Trigger main analysis:
        self.startDownload()

    def startDownload(self):
        """ Changes update message and executes the analysis"""
        self.labelError.configure(text="Downloading. This can take a moment.")
        self.after(1000, self.execute_analysis)

    def execute_analysis(self):
        """ Executes the main functions in the messages module. """
        try:
            scu = messages.SlackChannelsAndUsers(
                self.channel_var_get, self.channels_flag_var, self.users_flag_var,
                self.path_orig, self.path_dest
            )
            scu.get_all_channels_info()
            scu.get_all_users_info()

            # --From the class SlackMessages:
            sm = messages.SlackMessages(
                self.channel_var_get, self.channels_flag_var, self.users_flag_var,
                self.path_orig, self.path_dest
            )
            sm.get_all_messages_df()

            # --Update GUI:
            full_path = os.path.join(self.path_dest, dest_name_ext)
            self.labelError.configure(text=f'Download completed. Your files were saved in \n "{full_path}"')
            self.buttonContinue.configure(
                text="Done",
                state='normal',
                command=self.close_window
            )
        
        except:
            # --Update GUI:
            self.labelError.configure(text="An error occured")
            self.buttonContinue.configure(
                text="Exit",
                state='normal',
                command=self.close_window
            )


    def close_window(self):
        self.withdraw()
        self.quit()


gui = GUI()
