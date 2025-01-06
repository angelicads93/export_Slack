
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Oct 27 19:36:24 2024

@author: agds
"""

import tkinter as tk
from os.path import exists

import sys
import os

class GUI(tk.Tk):
    
    def __init__(self):
        
        super().__init__()

        ## Define text style to be used:
        self.font_choice = ('Helvetica',9)
        
        ## Defining some colors to use later for the interface:
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
        
        ## Title and size of the GUI:
        self.title = "Rebecca Everlene Trust"
        self.geometry('450x350')
        self.configure(bg=darkA)
        
        ## Frame and Label of the header:
        self.frame = tk.Frame(master=self, bg=bkgc)
        self.frame.pack(pady=5, padx=5, fill='both', expand=True)

        self.label = tk.Label(master=self.frame, width=300, text='Rebecca Everlene Trust', font=('Roboto', 20), bg=bannerc)
        self.label.configure(fg=titlec)
        self.label.pack(pady=5, padx=5)        

        ## Frame and entry boxes for slack channel's name, source-path and save-path:
        self.frame1 = tk.Frame(master=self, bg=bkgc)
        self.frame1.pack(pady=6, padx=10)
        
        self.labelChannel= tk.Label(master=self.frame1, text='Slack channel:', font=self.font_choice, bg=bkgc)
        self.labelChannel.configure(fg=letterc)
        self.labelChannel.pack(padx=5, pady=0, anchor='w')
        
        self.entryChannel = tk.Entry(master=self.frame1, font=self.font_choice, width=600, bg=boxc, border=0, disabledbackground=bkgc)
        self.entryChannel.configure(fg=letterc)
        self.entryChannel.pack(padx=5, pady=6)
        
        self.labelOrig= tk.Label(master=self.frame1, text='Path to source directory:', font=self.font_choice, bg=bkgc)
        self.labelOrig.configure(fg=letterc)
        self.labelOrig.pack(padx=5, pady=0, anchor='w')
        
        self.entryOrig = tk.Entry(master=self.frame1, font=self.font_choice, width=600, bg=boxc, border=0, disabledbackground=bkgc)
        self.entryOrig.configure(fg=letterc)
        self.entryOrig.pack(padx=5, pady=6)
        
        self.labelDest= tk.Label(master=self.frame1, text='Save in path:', font=self.font_choice, bg=bkgc)
        self.labelDest.configure(fg=letterc)
        self.labelDest.pack(padx=5, pady=0, anchor='w')
        
        self.entryDest = tk.Entry(master=self.frame1, font=self.font_choice, width=600, bg=boxc, border=0, disabledbackground=bkgc)
        self.entryDest.configure(fg=letterc)
        self.entryDest.pack(padx=5, pady=6)


        ### Label for showing error/update messages of the status of the code:
        self.txt = ' '
        self.labelError= tk.Label(master=self.frame1, text=self.txt, font=self.font_choice, bg=bkgc, width=600, wraplength=360, height=5)
        self.labelError.configure(fg=letterc)
        self.labelError.pack(padx=5, pady=2)        

        ### Button. Triggers error messages if information is incomplete, otherwise, it executes the analysis:
        self.button = tk.Button(self.frame1, text='Continue', font=self.font_choice, command=self.execute_analysis, border=0, highlightcolor=darkB, highlightbackground=bkgc)
        self.button.configure(bg=darkB, fg=letterc)
        self.button.pack(padx=5, pady=2)

        self.mainloop()
     

    def OK_button(self):
        channel = str(self.entryChannel.get())
        pathOrig = str(self.entryOrig.get())
        pathDest = str(self.entryDest.get())

        # --From class InspectSource:
        inspect_source = messages.InspectSource()
        channels_names = inspect_source.get_channels_names()
        all_channels_jsonFiles_dates = inspect_source.get_all_channels_json_names()
            
        ## Check the name of the Slack channel:
        if channel not in channels_names: 
            self.txt = 'Please enter a valid Slack channel'
            self.labelError.configure(text=self.txt)
            print(self.txt)

        ### Ask for the source path and check if the channels directory is there:
        elif channel in channels_names and pathOrig=='':
            self.txt = 'Please enter the path to the source directory'
            self.labelError.configure(text=self.txt)
            print(self.txt)
            
        elif channel in channels_names and pathOrig!='' and exists(f"{pathOrig}\\{channel}")==False:
            self.txt = f"The file for the channel '{channel}' was not found in the {pathOrig}"
            self.labelError.configure(text=self.txt)
            print(self.txt)

        ### Ask save-path path and check if it exists:
        elif channel in channels_names and pathOrig!='' and exists(f"{pathOrig}\\{channel}")==True and pathDest=='':
            self.txt = f"Please enter the path where the file will be saved"
            self.labelError.configure(text=self.txt)
            print(self.txt)

        elif channel in channels_names and pathOrig!='' and exists(f"{pathOrig}\\{channel}")==True and pathDest!='' and exists(pathDest)==False:
            self.txt = f"The path to save the file was not found"
            self.labelError.configure(text=self.txt)
            print(self.txt)

        ### If all looks ok, then execute the analysis:
        elif channel in channels_names and pathOrig!='' and exists(f"{pathOrig}\\{channel}")==True and pathDest!='' and exists(pathDest)==True:
            self.txt = f"Downloading..."
            self.labelError.configure(text=self.txt)
            print(self.txt)
            self.entryChannel.configure(state='disable')
            self.entryOrig.configure(state='disable')
            self.entryDest.configure(state='disable')

            self.startDownload()

    def startDownload(self):
        if self.txt=="Downloading...":
            self.labelError.configure(text="Downloading...")
            self.button.configure(state='disabled', command=self.close_window)
            # --Delay the update of the label Error and the Button:
            self.after(1000, self.execute_analysis)
            
    def execute_analysis(self):
        choosen_channel_name = 'general'  # str(self.entryChannel.get())
        source_path = r'C:\Users\angel\Documents\RebeccaEverleneTrust\RET_source'  # str(self.entryOrig.get())
        destination_path = r'C:\Users\angel\Desktop'  # str(self.entryDest.get())

        # --Build inputs.py and load as module
        current_path = os.path.dirname(os.path.realpath(__file__))
        print(f"current_path = {current_path}")
        with open(f"{current_path}\inputs.py", 'w') as f:
            f.write("missing_value = 'n/d' \n")
            f.write("timmeshift = 'US/Central' \n")
            f.write(f"key_wrd_text_show = {False} \n")
            f.write(f"continue_analysis = {True} \n")
            f.write(f"chosen_channel_name = '{choosen_channel_name}' \n")
            f.write(f"write_all_channels_info = {True} \n")
            f.write(f"write_all_users_info = {True} \n")
            f.write(f"slackexport_folder_path = r'{source_path}' \n")
            f.write(f"converted_directory = r'{destination_path}' \n")
        f.close()
        sys.path.append(current_path)
        import inputs

        # --From the class InspectSource:
        sys.path.append(r"C:\Users\angel\Documents\RebeccaEverleneTrust\export_Slack\src")
        import messages
        inspect_source = messages.InspectSource()
        self.channels_names = inspect_source.get_channels_names()
        self.all_channels_jsonFiles_dates = inspect_source.get_all_channels_json_names()

        # --From the class SlackChannelAndUsers:
        scu = messages.SlackChannelsAndUsers()
        scu.get_all_channels_info()
        scu.get_all_users_info()
        all_channels_df = scu.all_channels_df
        all_users_df = scu.all_users_df

        # --From the class SlackMessages:
        sm = messages.SlackMessages()
        channel_messages_df = sm.get_all_messages_df()
        
        self.after(500, self.end)


    def end(self):
        ### Change the text in the button to "Done" which allows to close the GUI:
        self.labelError.configure(text="Download completed")
        self.button.configure(text="Done", state='active', command=self.close_window)

    def close_window(self):
        self.withdraw()
        self.quit()
        
                  

gui = GUI() 
