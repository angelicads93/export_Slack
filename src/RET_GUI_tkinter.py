
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


class GUI(tk.Tk):

    def __init__(self):
        super().__init__()

        # --Define text style to be used:
        self.font_choice = ('Helvetica', 9)

        # --Defining some colors to use later for the interface:
        dark_gray = "#464C54"
        blue_gray = "#6B808E"
        ice_blue = "#D7DDE1"
        gray = "#CDCDCF"
        light_gray = "#E7E6E6"
        gold = '#FBC000'

        dark1 = "#1F1F1F"
        dark2 = "#2E2E2E"
        dark3 = "#4D4D4D"

        darkA = "#202124"
        darkB = "#29292D"
        darkC = "#686868"
        darkD = "#9BA0A6"
        white = "#FFFFFF"

        bkgc = darkA
        boxc = darkB
        letterc = white
        titlec = gold
        bannerc = darkA

        # --Initialize a default inputs.py file:
        self.build_input_file('channel', 'path_orig', 'path_dest',
                              False, False)

        # --Title and size of the GUI:
        self.title = "Rebecca Everlene Trust"
        self.geometry('650x350')
        self.configure(bg=darkA)

        # --Header:
        self.frame = tk.Frame(master=self, bg=bkgc)
        self.frame.pack(pady=5, padx=5, fill='both', expand=True)
        self.label = tk.Label(
            master=self.frame, text='Rebecca Everlene Trust',
            font=('Roboto', 20),
            width=300, bg=bannerc
            )
        self.label.configure(fg=titlec)
        self.label.pack(pady=5, padx=5)

        self.frame1 = tk.Frame(master=self, bg=bkgc)
        self.frame1.pack(pady=6, padx=10)

        # --Entry box for the source path:
        self.labelOrig = tk.Label(
            master=self.frame1, text='Path to source directory:',
            font=self.font_choice, bg=bkgc
            )
        self.labelOrig.configure(fg=letterc)
        self.labelOrig.pack(padx=5, pady=0, anchor='w')
        self.entryOrig = tk.Entry(
            master=self.frame1,
            font=self.font_choice,
            width=600, bg=boxc, border=0, disabledbackground=bkgc,
            highlightthickness=0
            )
        self.entryOrig.configure(fg=letterc)
        self.entryOrig.pack(padx=5, pady=6)

        # --Entry box for the destination path:
        self.labelDest = tk.Label(
            master=self.frame1, text='Save in path:',
            font=self.font_choice, bg=bkgc
            )
        self.labelDest.configure(fg=letterc)
        self.labelDest.pack(padx=5, pady=0, anchor='w')
        self.entryDest = tk.Entry(
            master=self.frame1,
            font=self.font_choice,
            width=600, bg=boxc, border=0, disabledbackground=bkgc,
            highlightthickness=0
            )
        self.entryDest.configure(fg=letterc)
        self.entryDest.pack(padx=5, pady=6)

        # --Entry box for name of Slack channel:
        self.labelChannel = tk.Label(
            master=self.frame1, text='Slack channel:',
            font=self.font_choice, bg=bkgc
            )
        self.labelChannel.configure(fg=letterc)
        self.labelChannel.pack(padx=5, pady=0, anchor='w')
        self.entryChannel = tk.Entry(
            master=self.frame1,
            font=self.font_choice,
            width=600, bg=boxc, border=0, disabledbackground=bkgc,
            highlightthickness=0
            )
        self.entryChannel.configure(fg=letterc)
        self.entryChannel.pack(padx=5, pady=6)

        # --Checkbox for channels_flag:
        self.channels_flag_var = tk.BooleanVar()
        self.channels_flag_var.set(False)
        self.checkbox_channels = tk.Checkbutton(
            master=self.frame1, variable=self.channels_flag_var,
            text=' Export general information of all the Slack channels',
            font=self.font_choice, fg=letterc,
            bg=bkgc, disabledforeground=bkgc, bd=0, highlightthickness=0,
            selectcolor=bkgc, activebackground=bkgc, activeforeground=letterc,
            onvalue=True, offvalue=False
            )
        self.checkbox_channels.pack(padx=5, pady=5, anchor='w')

        # --Checkbox for users_flag:
        self.users_flag_var = tk.BooleanVar()
        self.users_flag_var.set(False)
        self.checkbox_users = tk.Checkbutton(
            master=self.frame1, variable=self.users_flag_var,
            text=' Export general information of all the Slack users',
            font=self.font_choice, fg=letterc,
            bg=bkgc, disabledforeground=bkgc, bd=0, highlightthickness=0,
            selectcolor=bkgc, activebackground=bkgc, activeforeground=letterc,
            onvalue=True, offvalue=False
            )
        self.checkbox_users.pack(padx=5, pady=5, anchor='w')

        # --Label for showing error/update messages of the status of the code:
        self.txt = ' '
        self.labelError = tk.Label(
            master=self.frame1, text=self.txt,
            font=self.font_choice, bg=bkgc, width=600, wraplength=360, height=5
            )
        self.labelError.configure(fg=letterc)
        self.labelError.pack(padx=5, pady=2)

        # --Button. Triggers error messages if information is incomplete
        # --otherwise, it executes the analysis:
        self.button = tk.Button(
            self.frame1, text='Continue', command=self.check_inputs_exists,
            font=self.font_choice,
            border=0, highlightcolor=darkB, highlightbackground=bkgc
            )
        self.button.configure(bg=darkB, fg=letterc)
        self.button.pack(padx=5, pady=10)

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
        self.path_orig = str(self.entryOrig.get())
        self.path_dest = str(self.entryDest.get())
        self.channel = str(self.entryChannel.get())
        self.channels_flag = str(self.channels_flag_var.get())
        self.users_flag = str(self.users_flag_var.get())

        if self.path_orig == '':
            self.txt = 'Please enter the path to the source directory'
            self.labelError.configure(text=self.txt)
            print(self.txt)

        elif self.path_orig != '' and exists(f"{self.path_orig}") is False:
            self.txt = 'Please enter a valid path to the source directory'
            self.labelError.configure(text=self.txt)
            print(self.txt)

        elif self.channel == '' and self.path_orig != '' \
                and exists(f"{self.path_orig}") is True:
            self.txt = 'Please enter a Slack channel'
            self.labelError.configure(text=self.txt)
            print(self.txt)

        elif self.channel != '' and self.path_orig != '' \
                and exists(f"{self.path_orig}") is True:
            print('All input information provided')
            print('Checking consistency')

            inspect_source = messages.InspectSource(
                self.channel, self.path_orig, self.path_dest
                )
            save_in_path = inspect_source.save_in_path()
            inspect_source.check_save_path_exists(save_in_path)
            inspect_source.check_expected_files_exists()
            channels_names = inspect_source.get_channels_names()

            if self.channel != '' and self.channel not in channels_names:
                self.txt = 'Please enter a valid Slack channel'
                self.labelError.configure(text=self.txt)
                print(self.txt)

            elif self.channel in channels_names \
                    and exists(f"{self.path_orig}/{self.channel}") is False:
                self.txt = f"The file for the channel '{self.channel}' was not found in {self.path_orig}"
                self.labelError.configure(text=self.txt)
                print(self.txt)

            elif self.path_dest == '' or exists(self.path_dest) is False:
                self.txt = "Please enter a valid destination path"
                self.labelError.configure(text=self.txt)
                print(self.txt)

            # --If all looks ok, then execute the analysis:
            elif self.channel in channels_names \
                    and self.path_dest != '' \
                    and exists(self.path_dest) is True:
                self.build_input_file(
                    self.channel, self.path_orig, self.path_dest,
                    self.channels_flag, self.users_flag
                    )
                self.txt = "Downloading..."
                self.labelError.configure(text=self.txt)
                print(self.txt)
                self.entryChannel.configure(state='disable')
                self.entryOrig.configure(state='disable')
                self.entryDest.configure(state='disable')

                self.startDownload()

    def startDownload(self):
        """ Changes update message and disables the user's inputs. """
        if self.txt == "Downloading...":
            self.labelError.configure(text="Downloading...")
            self.button.configure(state='disabled', command=self.close_window)
            # --Delay the update of the label Error and the Button:
            self.after(1000, self.execute_analysis)

    def execute_analysis(self):
        """ Executes the main functions in the messages module. """
        scu = messages.SlackChannelsAndUsers(
            self.channel, self.channels_flag, self.users_flag,
            self.path_orig, self.path_dest
            )
        scu.get_all_channels_info()
        scu.get_all_users_info()

        # --From the class SlackMessages:
        sm = messages.SlackMessages(
            self.channel, self.channels_flag, self.users_flag,
            self.path_orig, self.path_dest
            )
        channel_messages_df = sm.get_all_messages_df()

        self.after(500, self.end)

    def end(self):
        """ Change button's text to "Done" and allows to close the GUI. """
        self.labelError.configure(text="Download completed")
        self.button.configure(
            text="Done",
            state='active',
            command=self.close_window
            )

    def close_window(self):
        self.withdraw()
        self.quit()


gui = GUI()
