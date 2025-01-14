## To download and open the source code in Visual Studio Code:
* Visit https://github.com/angelicads93/export_Slack, click on the green button “<>code” and select the option “Download ZIP”.
* Unzip and save this folder in the directory of your choice. You can rename the folder as “exportSlack”.
* Open Virtual Studio Code.
* Click on “File” > “Open Folder”. Or click directly on “Open Folder” from the welcome scream of Visual Studio Code.
* Navigate through your files where the folder “exportSlack” is, and click on “Select Folder”. You should see “EXPORTSLACK” in the “Explorer” tab of your Virtual Studio Code interface.
* We can now open a terminal for the next steps, you can do so by clicking “Terminal” > “New Terminal”. 

## To create the virtual environment:
* Create a virtual environment by running in the powershell,
  ```{script}
  python3 -m venv .\venv
  ```
  This should create a new folder “venv” that you could see in the explorer tab of your VScode interface.
* Activate the virtual environment you just created by running in the powershell,
  ```{sript}
   .\venv\Scripts\activate
  ```
  You will see "(venv)" at the beginning of the line with the command prompt.
* Install the required dependencies in your virtual environment by running in the powershell,
  ```{script}
   pip install -r requirements.txxt
  ```
   When everything has been installed, you should see in your powershell a message that says “Successfully installed“ followed by the names of all the dependencies.

## To specify the conditions of the analysis:
* From your explorer tab, open the file called “inputs.py” inside the “src” directory.
* Specify the conditions of your choice.
* Save the changes.

## To run the analysis:
* From the powershell, and within the same virtual environment venv, run the command
  ```{script}
  python3 .\src\main.py
  ```

After you have finished using exportSlack, you can deactivate the virtual environment by running the following command in your powershell:
```{script}
deativate
```

