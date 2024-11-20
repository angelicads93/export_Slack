
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Oct 27 19:36:24 2024

@author: agds
"""

import tkinter as tk
import pandas as pd
from json import load
from datetime import datetime
from os import listdir
from os.path import getmtime, exists
from pathlib import Path



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
        self.button = tk.Button(self.frame1, text='Continue', font=self.font_choice, command=self.OK_button, border=0, highlightcolor=darkB, highlightbackground=bkgc)
        self.button.configure(bg=darkB, fg=letterc)
        self.button.pack(padx=5, pady=2)

        self.mainloop()
     
            
    def OK_button(self):
        channel = str(self.entryChannel.get())
        pathOrig = str(self.entryOrig.get())
        pathDest = str(self.entryDest.get())
            
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

            ### Delay the update of the label Error and the Button:
            self.after(1000, self.execute_analysis)
            
  
    def execute_analysis(self):
        """
        .) The code below is Rutvikk's file + extra modification to compile each channel's information.
        .) To simplify the analysis, the dataframe with all the messages from all the channels is not built, 
        intead, the specified Slack channel is extracted from the begining. 
        .) Another difference w.r.t. Rutvikk's code is that only the features shown in the current version of the 
        spreadsheets are extracted (There are more features in the original version).
        .) The notation changed some from the original version.
        .) The condition to using this app is that the user has to have a local copy of the Slack-export containing 
        the directory of the corresponding channel, and the files channels.json and users.json (all in the same directory)
        """

        ## Extract the information given by the user in the entry-boxes:
        exportname = str(self.entryChannel.get())
        slackexport_folder_path = str(self.entryOrig.get())
        download_path = str(self.entryDest.get())
        Path(download_path).mkdir(parents=True, exist_ok=True)   

        ## Read the json file with all the Slack channels:
        channels_path = f"{slackexport_folder_path}/channels.json"
        with open(channels_path, encoding='utf-8') as f:
            channels_json = load(f)
        
        ## Define function to extract channel's info from json files:  
        def slack_json_to_dataframe(slack_json):
            messages_df = pd.DataFrame(columns=["msg_id", "ts", "user", "type", "text"])
            for message in range(len(slack_json)):
                if 'files' in slack_json[message] and slack_json[message]['files']:
                    messages_df.at[message, "msg_id"] = slack_json[message]['files'][0]['id']
                elif 'client_msg_id' in slack_json[message]:
                    messages_df.at[message, "msg_id"] = slack_json[message]['client_msg_id']
                if 'ts' in slack_json[message]:
                    messages_df.at[message, "ts"] = slack_json[message]['ts']
                else:
                    messages_df.at[message, "ts"] = None
                messages_df.at[message, "user"] = slack_json[message].get('user', None)
                if 'type' in slack_json[message]:
                    messages_df.at[message, "type"] = slack_json[message]['type']
                else:
                    messages_df.at[message, "type"] = None
                if 'text' in slack_json[message]:
                    messages_df.at[message, "text"] = slack_json[message]['text']
                else:
                    messages_df.at[message, "text"] = None
            return messages_df
        
        ## Store the messages of the desired Slack channel in channel_messages_df: 
        channel_messages_df = pd.DataFrame(columns=["msg_id", "ts", "user", "type", "text", "json_name", "json_mod_date"])
        channel_folder_path = f"{slackexport_folder_path}/{exportname}"
        days = listdir(channel_folder_path)
        for file_day in range(len(days)):
            parentfolder_path = f"{slackexport_folder_path}/{exportname}"
            filejson_path = f"{parentfolder_path}/{days[file_day]}"
            with open(filejson_path, encoding='utf-8') as f:
                import_file_json = load(f)
            day_messages_df = slack_json_to_dataframe(import_file_json)
            day_messages_df['json_name'] = days[file_day]
            day_messages_df['json_mod_date'] = datetime.fromtimestamp( getmtime(filejson_path) )
            day_messages_df['channel'] = exportname
            channel_messages_df = pd.concat([channel_messages_df, day_messages_df], ignore_index=True)            
                    
        ## Read the users json file:
        users_path = f"{slackexport_folder_path}/users.json"
        all_users_df = pd.read_json(users_path)
        
        ## Store user's info in channel_users_df:
        channel_users_df = pd.DataFrame(columns=["user_id", "team_id", "name", "display_name", "is_bot"])
        channel_users_list = channel_messages_df['user'].unique()
        lenChannelUsers = len(channel_users_list)
        for i in range(lenChannelUsers):
            user_id = all_users_df[all_users_df['id']==channel_users_list[i]]['id'].values
            if len(user_id)==1:
                channel_users_df.at[i, "user_id"]      = user_id[0]
                channel_users_df.at[i, "team_id"]      = all_users_df[all_users_df['id']==channel_users_list[i]]['team_id'].values[0]
                channel_users_df.at[i, "name"]         = all_users_df[all_users_df['id']==channel_users_list[i]]['name'].values[0]
                channel_users_df.at[i, "display_name"] = all_users_df[all_users_df['id']==channel_users_list[i]]['real_name'].values
                channel_users_df.at[i, "is_bot"]       = all_users_df[all_users_df['id']==channel_users_list[i]]['is_bot'].values[0]
            elif len(user_id)==0:
                channel_users_df.at[i, "user_id"]      = 'SlackBot'
                channel_users_df.at[i, "team_id"]      = 'SlackBot'
                channel_users_df.at[i, "name"]         = 'SlackBot'
                channel_users_df.at[i, "display_name"] = 'SlackBot'
                channel_users_df.at[i, "is_bot"]       = True  

        ### Use channel_users_df to fill-in the user's information in channel_messages_df:
        for index in channel_messages_df.index.values:
            channel_messages_df.at[index, 'name'] = channel_users_df[channel_users_df['user_id']==channel_messages_df.loc[index]['user']]['name'].values
            channel_messages_df.at[index, 'display_name'] = channel_users_df[channel_users_df['user_id']==channel_messages_df.loc[index]['user']]['display_name'].values
        
        ### Reorder the columns:
        channel_messages_df = channel_messages_df[['channel', 'json_name', 'json_mod_date', 'user', 'name', 'display_name', 'ts', 'msg_id', 'type', 'text']]
        channel_messages_df.index = ['']*len(channel_messages_df)
        
        ### Change format of the time in seconds to a date:
        channel_mindate = pd.to_datetime(channel_messages_df['ts'], unit='s').min().date()
        channel_maxdate = pd.to_datetime(channel_messages_df['ts'], unit='s').max().date()
        channel_messages_df['ts'] = pd.to_datetime(channel_messages_df['ts'], unit='s')
        channel_messages_df.rename(columns={"ts": "msg_date"}, inplace=True)
        channel_messages_df.sort_values(by='msg_date', inplace=True)                  
            
        ### Write .csv file:
        cleanExportChannel_path = f"{download_path}/{exportname}_{channel_mindate}_to_{channel_maxdate}_compiled.csv"
        channel_messages_df.to_csv(cleanExportChannel_path, index=False)
        
        self.after(500, self.end) 

    def end(self):
        ### Change the text in the button to "Done" which allows to close the GUI:
        self.labelError.configure(text="Download completed")
        self.button.configure(text="Done", state='active', command=self.close_window)

    def close_window(self):
        self.withdraw()
        self.quit()
        
                  







channels_names = ['general', 'campaign-ret', 'random', 'champions', 'tutors-on-call', 'aspects-tuition-reimbursement-scholarships-special-projects', 'team-issues-medkids', 'landmarks', 'college-aspects', 'licensing', '3d-art-blender-landmarks', 'team-anatomy-island-for-medkids', 'aspects-data-analysis', 'landmarks-2d-art-characters',
                  'real-estate-verification', 'landmarks-unity-developers', 'grants', 'team-voiceover-artists-med-kids', 'outreach-fundraising-communications', 'doodly-toonly-cartoons-medkids', 'landmarks-game-designers', 'landmarks-content-development', 'landmarks-locations', 'pm-tl-channel', 'mockups-for-strategy-finance-budgets',
                  'sae-performing-arts-medkids', 'team-leads', 'landmarks-landing-page', 'fundraising-initiatives', 'landmarks-sprint-planning', 'team-writers-for-rebecca-everlene', 'grants-team-2024', 'aspects-design-content-updates', 'landmarks-sprint', 'strategic-planning', 'unequivocally-big-ux-ui', 'team-audio-med-kids',
                  'social-media-branding-team', 'team-infographics-and-charting', 'team-2d-art-for-medkids', 'landmarks-forward', 'unequivocally-big-summer-2023', 'made-ux-ui', 'team-game-designers-medkids', 'landmarks-mapping', 'landmarks-unity-lightship-squad', 'landmarks-2d-art-items-and-guides', 'landmarks-2d-art-locations',
                  'landmarks-ux-meeting-notes', 'team-chef-medkids', 'team-front-end-web-developers-medkids', 'landmarks-ux-ui-designers', 'team-christa-medkids-games', 'aspects-automation-team', 'team-kamil-medkids-games', 'team-maulana-medkids-games', 'landmarks-vertical-slice-team', 'team-chef-logistics-medkids', 'team-animators-medkids',
                  'landmarks-geocodes', '3d-art-team-a-landmarks', '3d-art-team-b-landmarks', 'landmarks-interactive-wall', 'landmarks-product-managers', 'landmarks-sprints-2d-art', 'landmarks-production-team', 'presidential-service-award', 'team-zixi-medkids-games', '20-team', 'team-barbara-medkids-games', 'team-avatars-medkids', 
                  'landmarks-uxreview-gd-dev', 'scrum', 'team-ux-ui-designers-medkids', 'team-bowen-medkids-games', 'think-biver-saturday-checkins', 'tracker-board', 'team-tech-order-up', 'team-nabil-medkids-games', 'team-yigit-medkids-games', 'budget-template-build', 'team-back-end-dev', 'chief-of-chatting-space', 'artxdev-team', 
                  'unequivocally-big-summer-2024', 'salesforce-automation-team', 'unequivocally-big-suite-content-team', 'eat-like-us-inventory-project', 'design-brainstorming-sessions', 'cpts-training-team', 'team-orderup-developers-medkids', 'aspects-data-cleanup', 'aws-automation-team', 'team-dev-issues-board', 'the-jog-app', 
                  'team-github-solutions', 'cpts-strategy-team', 'dreampad-for-dreamforce', 'team-jira', 'team-scapegoated', 'team-google-workspace', 'smitten-hitch-coparenting-project', 'team-medkids-minitours', 'team-cybersecurity-innovations', 'team-ink', 'team-budgets', 'team-azure', 'intros-and-shoutouts', 'time-off', 
                  'spaulding-daniels-leadership-group', 'how-do-i', 'policies-and-sops-team', 'onboarding-central']
gui = GUI() 
