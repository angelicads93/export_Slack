#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Oct 27 19:36:24 2024
@author: agds
"""

import tkinter as tk
import sys
import os
from os.path import exists, dirname, join

parent_dir = dirname(os.getcwd())
sys.path.append(parent_dir)
sys.path.append(join(parent_dir, 'src'))
import messages
from settings import dest_name_ext, channels_json_name, users_json_name


class GUI(tk.Tk):

    def __init__(self):
        super().__init__()

        # --Define text styles:
        self.font_base = ('bahnschrift', 11)
        self.font_base_bold = ('bahnschrift', 12, 'bold')
        self.font_msg = ('bahnschrift', 12, 'bold')
        self.font_help = ('bahnschrift', 10)

        # --Defining backgraound and text colors:
        self.bkgc = '#121212'
        self.canvasc = '#1d1d1d'
        self.boxc = '#2e2e2e'
        self.letterc = 'white'
        self.letter_bkg = self.boxc
        self.titlec = '#FBC000'
        self.bannerc = self.bkgc

        # --Initialize a default inputs.py file:
        self.build_input_file('channel', 'path_orig', 'path_dest',
                              False, False)

        # --Title and size of the GUI:
        self.title("Slack2Excel")
        self.geometry('650x720')
        self.resizable(0, 0)

        # --Header:
        self.frame = tk.Frame(master=self, bg=self.bkgc)
        self.label = tk.Label(
            master=self.frame, text='Slack Channels to Excel Databases',
            font=('Roboto', 24, 'bold'),
            width=300, bg=self.bannerc, fg=self.titlec
        )

        # --Canvas where the user's input are entered:
        self.canvas = tk.Frame(master=self.frame, bg=self.canvasc)

        # --Label for the source path:
        self.labelOrig = tk.Label(
            master=self.canvas, text='Path to the source directory:',
            font=self.font_base, bg=self.canvasc, fg=self.letterc
        )

        # --Entry box for the source path:
        self.entryOrig = tk.Entry(
            master=self.canvas,
            font=self.font_base, bg=self.boxc, fg=self.letterc,
            disabledbackground=self.boxc,
            width=600, border=0, highlightthickness=0
        )

        # --Label for the destination path:
        self.labelDest = tk.Label(
            master=self.canvas, text='Save in path:',
            font=self.font_base, bg=self.canvasc, fg=self.letterc
        )

        # --Entry box for the destination path:
        self.entryDest = tk.Entry(
            master=self.canvas,
            font=self.font_base, bg=self.boxc, fg=self.letterc,
            disabledbackground=self.boxc, readonlybackground=self.boxc,
            width=600, border=0, highlightthickness=0,
        )

        # --Label for the channel name:
        self.labelChannel = tk.Label(
            master=self.canvas, text='Name of the Slack channel:',
            font=self.font_base, bg=self.canvasc, fg=self.letterc
        )

        # --Option Menu for the available channels in the source path:
        self.default_str_channel = ' '
        self.channel_var = tk.StringVar(self.frame)
        self.channel_var.set(self.default_str_channel)
        self.Channel = tk.OptionMenu(
            self.canvas, self.channel_var, value=[self.default_str_channel]
        )
        self.Channel.configure(
            font=self.font_base, bg=self.boxc, fg=self.letterc,
            width=600, border=0, highlightthickness=0,
            state='disabled', justify='left', anchor='w'
        )

        # --RadioButtons for converting channels.json:
        self.labelChannelFlag = tk.Label(
            master=self.canvas,
            text='Include summary of channels in the Slack workspace:',
            font=self.font_base, bg=self.canvasc, fg=self.letterc
        )
        self.button_channels = tk.StringVar()
        self.button_channels.set('empty')
        self.buttonChannelsYes = tk.Radiobutton(
            master=self.frame, variable=self.button_channels,
            text='Yes', value='y',
            font=self.font_base, bg=self.canvasc, fg=self.letterc,
            disabledforeground=self.letter_bkg,
            bd=0, highlightthickness=0, selectcolor=self.canvasc,
            activebackground=self.canvasc, activeforeground=self.letterc
        )
        self.buttonChannelsNo = tk.Radiobutton(
            master=self.canvas, variable=self.button_channels,
            text='No', value='n',
            font=self.font_base, bg=self.canvasc, fg=self.letterc,
            disabledforeground=self.letter_bkg,
            bd=0, highlightthickness=0, selectcolor=self.canvasc,
            activebackground=self.canvasc, activeforeground=self.letterc
        )

        # --RadioButtons for converting users.json:
        self.labelUsersFlag = tk.Label(
            master=self.canvas,
            text='Include summary of users in the Slack workspace:',
            font=self.font_base, bg=self.canvasc, fg=self.letterc
        )
        self.button_users = tk.StringVar()
        self.button_users.set('empty')
        self.buttonUsersYes = tk.Radiobutton(
            master=self.canvas, variable=self.button_users,
            text='Yes', value='y',
            font=self.font_base, bg=self.canvasc, fg=self.letterc,
            disabledforeground=self.letter_bkg,
            bd=0, highlightthickness=0, selectcolor=self.canvasc,
            activebackground=self.canvasc, activeforeground=self.letterc
        )
        self.buttonUsersNo = tk.Radiobutton(
            master=self.canvas, variable=self.button_users,
            text='No', value='n',
            font=self.font_base, bg=self.canvasc, fg=self.letterc,
            disabledforeground=self.letter_bkg,
            bd=0, highlightthickness=0, selectcolor=self.canvasc,
            activebackground=self.canvasc, activeforeground=self.letterc
        )

        # --Label for showing errors and help:
        self.labelError = tk.Label(
            master=self.frame, text=' ', height=1, wraplength=500,
            font=self.font_msg, bg=self.bkgc, fg=self.letterc
        )
        self.labelHelp = tk.Label(
            master=self.frame, text=' ', height=4, wraplength=550,
            font=self.font_help, bg=self.bkgc, fg=self.letterc, justify='left'
        )

        # --Update and Continue buttons:
        self.buttonOK = tk.Button(
            self.frame, text='Update', command=self.check_inputs_exists,
            font=self.font_base, bg=self.canvasc, fg=self.titlec,
            border=0
        )
        self.buttonContinue = tk.Button(
            self.frame, text='Convert', command=self.check_inputs_exists,
            font=self.font_base, bg=self.canvasc, fg=self.titlec,
            border=0, state='disabled'
        )

        # --Pack widgets in desired order:
        pad_x = 15
        self.frame.pack(padx=0, pady=0, fill='both', expand=True, side='top')
        self.label.pack(padx=pad_x, pady=(15, 25))
        self.canvas.pack(padx=30)
        self.labelOrig.pack(padx=pad_x, pady=(15, 0), anchor='w')
        self.entryOrig.pack(padx=pad_x, pady=(0, 15), anchor='w', ipady=9)
        self.labelDest.pack(padx=pad_x, pady=0, anchor='w')
        self.entryDest.pack(padx=pad_x, pady=(0, 15), anchor='w', ipady=9)
        self.labelChannel.pack(padx=pad_x, pady=0, anchor='w')
        self.Channel.pack(
            padx=pad_x, after=self.labelChannel, anchor='w', ipady=9
        )

        self.labelChannelFlag.pack(
            padx=pad_x, pady=(15, 0), after=self.Channel, anchor='w'
        )
        self.buttonChannelsYes.pack(
            padx=20, pady=3, side='top', anchor='w',
            after=self.labelChannelFlag
        )
        self.buttonChannelsNo.pack(
            padx=20, pady=(3, 15), side='top', anchor='w',
            after=self.buttonChannelsYes
        )

        self.labelUsersFlag.pack(padx=pad_x, side='top', anchor='w')
        self.buttonUsersYes.pack(padx=20, pady=3, side='top', anchor='w')
        self.buttonUsersNo.pack(padx=20, pady=(3, 15), side='top', anchor='w')

        self.labelError.pack(padx=pad_x, pady=(30, 12), fill='x')
        self.labelHelp.pack(padx=pad_x, pady=(0, 5), fill='x')

        self.buttonContinue.pack(
            padx=(4,30), pady=(5, 30), anchor='s', side='right', ipady=6,
            after=self.labelHelp
        )
        self.buttonOK.pack(
            padx=2, pady=(5, 30), anchor='s', side='right', ipady=6,
            after=self.buttonContinue
        )

    def set_title(self, new_title):
        self.title(new_title)

    def build_input_file(self, channel, path_orig, path_dest,
                         channels_flag, users_flag):
        """"Creates a inputs.py file with the user's information."""
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

    def configure_labels(self, widgets, font_, color_):
        """ Configure the font style and color of label type widgets. """
        widgets_labels = [self.labelOrig, self.labelDest, self.labelChannel,
                          self.labelChannelFlag, self.labelUsersFlag]
        for widget in widgets:
            if widget in widgets_labels:
                widget.configure(fg=color_, font=font_)

    def configure_widgets_int(self, widgets, state_, font_):
        """ Configure the state and font style of interactive weidgets. """
        widgets_int = [self.entryOrig, self.entryDest, self.Channel,
                       self.buttonChannelsYes, self.buttonChannelsNo,
                       self.buttonUsersYes, self.buttonUsersNo]
        for widget in widgets:
            if widget in widgets_int:
                widget.configure(state=state_, font=font_)

    def reset_widgets(self):
        """ Sets configuration of all widgets to the defines in __init__"""
        self.configure_labels(
            [self.labelOrig, self.labelDest, self.labelChannel,
             self.labelChannelFlag, self.labelUsersFlag],
            self.font_base, self.letterc
        )
        self.configure_widgets_int(
            [self.entryOrig, self.entryDest, self.buttonChannelsYes,
             self.buttonChannelsNo, self.buttonUsersYes, self.buttonUsersNo],
            'normal', self.font_base
        )

    def check_inputs_exists(self):
        """ Verifies the validity of the input. Execute analysis if it's ok"""
        print('-------------------------------------------------------------')
        # -- Retrieve user's input:
        self.path_orig = str(self.entryOrig.get()).replace('"', '')
        self.path_dest = str(self.entryDest.get()).replace('"', '')
        self.channel_var_get = str(self.channel_var.get())
        self.channels_flag = self.button_channels.get()
        self.users_flag = self.button_users.get()

        # --Translate radiobutton variables to booleans:
        if self.button_channels.get() == 'y':
            self.channels_flag = True
        elif self.button_channels.get() == 'n':
            self.channels_flag = False
        if self.button_users.get() == 'y':
            self.users_flag = True
        elif self.button_users.get() == 'n':
            self.users_flag = False

        # --1. VERIFY THE SOURCE PATH:
        if self.path_orig == '':
            self.labelError.configure(
                text='Please enter the path to the source directory'
            )
            self.labelHelp.configure(
                text='HELP: The source directoy is the folder with the information exported from your Slack workspace. \nTo access its path, use your file explorer to find the desire folder, right-click on the name of the folder, select "Copy as path" and paste it in the entry box above.'
            )
            self.configure_labels(
                [self.labelOrig], self.font_base_bold,  self.letterc
            )
            self.configure_labels(
                [self.labelDest, self.labelChannel, self.labelChannelFlag,
                 self.labelUsersFlag],
                self.font_base, self.letter_bkg
            )
            self.configure_widgets_int(
                [self.entryDest, self.buttonChannelsYes, self.buttonChannelsNo,
                 self.buttonUsersYes, self.buttonUsersNo],
                'disabled', self.font_base
            )
        elif self.path_orig != '' and exists(self.path_orig) is False:
            self.labelError.configure(
                text='Please enter a valid path to the source directory'
            )
        elif self.path_orig != '' and exists(self.path_orig) is True:
            self.reset_widgets()
            self.entryDest.configure(state='normal')

            # --2. VERIFY DESTINATION PATH:
            if self.path_dest == '':
                # --Update GUI:
                self.labelError.configure(
                    text='Please enter the destination path'
                )
                self.labelHelp.configure(
                    text='HELP: The destination directory is the folder were you wish to save the converted files. \nTo access its path, use your file explorer to find the desire folder, right-click on the name of the folder, select "Copy as path" and paste it in the entry box above.'
                )
                self.configure_labels(
                    [self.labelDest], self.font_base_bold,  self.letterc
                )
                self.configure_labels(
                    [self.labelOrig, self.labelChannel, self.labelChannelFlag,
                     self.labelUsersFlag],
                    self.font_base, self.letter_bkg
                )
                self.configure_widgets_int(
                    [self.entryOrig, self.Channel, self.buttonChannelsYes,
                     self.buttonChannelsNo, self.buttonUsersYes,
                     self.buttonUsersNo],
                    'disabled', self.font_base
                )
            elif self.path_dest != '' and exists(self.path_dest) is False:
                # --Update GUI:
                self.labelError.configure(
                    text='Please enter a valid destination path'
                )
                self.configure_labels(
                    [self.labelDest], self.font_base_bold,  self.letterc
                )
                self.configure_labels(
                    [self.labelOrig, self.labelChannel, self.labelChannelFlag,
                     self.labelUsersFlag],
                    self.font_base, self.letter_bkg
                )
                self.configure_widgets_int(
                    [self.entryOrig, self.Channel, self.buttonChannelsYes,
                     self.buttonChannelsNo, self.buttonUsersYes,
                     self.buttonUsersNo],
                    'disabled', self.font_base
                )
            elif self.path_dest != '' and exists(self.path_dest) is True:
                if os.access(self.path_dest, os.X_OK) is False:
                    print('AG: Cannot write in the destination directory')
                else:
                    self.reset_widgets()
                    # --3. VERIFY CHANNEL:
                    # --Get names of channels in the source directory:
                    inspect_source = messages.InspectSource(
                        '', self.path_orig, self.path_dest
                    )
                    # --Sort in alphabetical order:
                    channels_names = sorted(
                        inspect_source.get_channels_names()
                    )
                    # --Add option of analysing all channels:
                    channels_names.insert(0, 'All-channels')
                    if dest_name_ext in channels_names:
                        channels_names.remove(dest_name_ext)
                    print('AG: CHANNELS NAMES = ', channels_names)

                    # --Activate menu with channels_name as choices:
                    if self.channel_var_get == self.default_str_channel:
                        self.Channel.destroy()
                        self.channel_var = tk.StringVar()
                        self.channel_var.set(self.default_str_channel)
                        self.Channel = tk.OptionMenu(
                            self.canvas,
                            self.channel_var,
                            *channels_names
                        )
                        self.Channel.configure(
                            font=self.font_base, bg=self.boxc, fg=self.letterc,
                            width=600, border=0, highlightthickness=0,
                            borderwidth=0, activebackground=self.boxc,
                            activeforeground=self.letterc,
                            foreground=self.letterc,
                            justify='left', direction='below', anchor='w'
                        )
                        self.Channel['menu'].configure(
                            bg=self.boxc, fg=self.letterc,
                            activebackground='gray',
                            activeforeground=self.letterc,
                            border=0, borderwidth=0
                        )
                        self.Channel.pack(
                            padx=15, after=self.labelChannel,
                            anchor='w', ipady=9
                        )

                        # --Update GUI:
                        self.labelError.configure(
                            text='Please enter a Slack channel'
                        )
                        self.labelHelp.configure(
                            text='Help: Click on the icon of the drop-down menu above and select the name of the channel you wish to analyze.'
                        )
                        self.configure_labels(
                            [self.labelChannel],
                            self.font_base_bold,  self.letterc
                        )
                        self.configure_labels(
                            [self.labelOrig, self.labelDest,
                             self.labelChannelFlag, self.labelUsersFlag],
                            self.font_base, self.letter_bkg
                        )
                        self.configure_widgets_int(
                            [self.entryOrig, self.entryDest,
                             self.buttonChannelsYes, self.buttonChannelsNo,
                             self.buttonUsersYes, self.buttonUsersNo],
                            'disabled', self.font_base
                        )

                    else:  # if channel is assigned
                        self.reset_widgets()
                        # --4. VERIFY CHANNELS FLAG:
                        if self.button_channels.get() == 'empty':
                            self.labelError.configure(
                                text='Please enter "Yes" or "No".'
                            )
                            self.labelHelp.configure(
                                text=f'HELP: This file contains general information of all the channels present in your Slack workspace, and it would be save in your destination path as _{channels_json_name}.'
                            )
                            self.configure_labels(
                                [self.labelChannelFlag],
                                self.font_base_bold,  self.letterc
                            )
                            self.configure_widgets_int(
                                [self.buttonChannelsYes, self.buttonChannelsNo],
                                'normal', self.font_base_bold
                            )
                            self.configure_labels(
                                [self.labelOrig, self.labelDest,
                                 self.labelChannel, self.labelUsersFlag],
                                self.font_base, self.letter_bkg
                            )
                            self.configure_widgets_int(
                                [self.entryOrig, self.entryDest, self.Channel,
                                 self.buttonUsersYes, self.buttonUsersNo],
                                'disabled', self.font_base
                            )
                        # --5. VERIFY USERS FLAG:
                        elif self.button_users.get() == 'empty':
                            self.labelError.configure(
                                text='Please enter "Yes" or "No".'
                            )
                            self.labelHelp.configure(
                                text=f'HELP: This file contains general information of all the users present in your Slack workspace, and it would be save in your destination path as _{users_json_name}.'
                            )
                            self.reset_widgets()
                            self.configure_labels(
                                [self.labelUsersFlag],
                                self.font_base_bold,  self.letterc
                            )
                            self.configure_widgets_int(
                                [self.buttonUsersYes, self.buttonUsersNo],
                                'normal', self.font_base_bold
                            )
                            self.configure_labels(
                                [self.labelOrig, self.labelDest,
                                 self.labelChannel, self.labelChannelFlag],
                                self.font_base, self.letter_bkg
                            )
                            self.configure_widgets_int(
                                [self.entryOrig, self.entryDest, self.Channel,
                                 self.buttonChannelsNo, self.buttonChannelsYes
                                 ], 'disabled', self.font_base
                            )

                        else: # --All good!
                            print('All input information provided')
                            print('Checking consistency')
                            print(self.path_orig)
                            print(self.path_dest)
                            print(self.channel_var_get)
                            print(self.channels_flag)
                            print(self.users_flag)
                            self.reset_widgets()
                            self.configure_widgets_int(
                                [self.Channel],
                                'normal', self.font_base
                            )
                            self.labelError.configure(
                                text='Good! Proceed to convert your file(s).'
                            )
                            self.labelHelp.configure(text='')
                            self.buttonOK.configure(state='disabled')
                            self.buttonContinue.configure(
                                state='normal', command=self.setup_inputs
                            )


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
        if self.channel_var_get == 'All-channels':
            self.channel_var_get = ''
        self.build_input_file(
            self.channel_var_get,
            self.path_orig, self.path_dest,
            self.channels_flag, self.users_flag
        )

        # --Update visual interface label and button:
        self.configure_widgets_int(
            [self.entryOrig, self.entryDest, self.Channel,
             self.buttonChannelsYes, self.buttonChannelsNo,
             self.buttonUsersYes, self.buttonUsersNo],
            'disabled', self.font_base
        )
        self.configure_labels(
            [self.labelOrig, self.labelDest, self.labelChannel,
             self.labelChannelFlag, self.labelUsersFlag],
            self.font_base, self.letter_bkg
        )

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
                self.channel_var_get, self.channels_flag, self.users_flag,
                self.path_orig, self.path_dest
            )
            scu.get_all_channels_info()
            scu.get_all_users_info()

            # --From the class SlackMessages:
            sm = messages.SlackMessages(
                self.channel_var_get, self.channels_flag, self.users_flag,
                self.path_orig, self.path_dest
            )
            sm.get_all_messages_df()

            # --Update GUI:
            full_path = join(self.path_dest, dest_name_ext)
            self.reset_widgets()
            self.configure_widgets_int(
                [self.Channel], 'normal', self.font_base
            )
            self.labelError.configure(text='Download completed.')
            self.labelHelp.configure(
                text=f'Your files were saved in \n "{full_path}"'
            )
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


if __name__ == '__main__':
    gui = GUI()
    gui.mainloop()
