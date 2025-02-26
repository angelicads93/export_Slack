# Change Log

**02/26/2025:**
* Replaced empty spaces with a dash when naming the Excel files in `main.py`. This is only a problem for channels starting with FC_.
* Added first draft of Excel file with stats on weekly reports. The main script is in `src/stats.py` and the excel settings are in `settings_stats.py`. 
Needed to geralized some aspects of the `excel.py` implementation to make it flexible and usable for both `main.py` and `stats.py`.
* More exceptions were added to the reports parser.

**02/21/2025:**
Changed the structure of the git repo:
* Input files `inputs.py` and `settings.py` are in the main repo directory.
* There are now three folders for the variants: VScode, JupyterNotebook, and GUI.
* `./VScode` has the Python script `main.py` and documentation on how to use it in VScode.
* `./JupyterNotebook` has the notebook `main.ipynb`, which contains its documentation.
* `./GUI` has the executable file slack2excel.exe, the python script `GUI_tkinter.py` with the main code of the graphic interface, and the subdirectory `tkGUI`, 
which is the output from generating the one-file executable app with pyinstaller.
* The requirements.txt and the bash scripts were collected in the directory `./dependencies`.
* Screenshots used to build documentation are stored in the directory `./images`.
* Updated the README file.
* Added a CHANGELOG file with some of the previous updates.
* The new file structure required changing how specific modules were imported in `main.py` and `GUI_tkinter.py`.

**02/14/2025:**
Extracted check-in variables and Excel variables to the file `settings.py`.
* The check-in variables in `settings.py` are now: all_keywords, index_keyword, keywords_dictionary, sample_text_list. Extracted from `checkins.py` and `excel.py`.
* The excel variables in `settings.py` are now: column_width, text_type_cols, height_1strow, alignment_vert_1strow, alignment_horiz_1strow, font_size_1strow, font_bold_1strow, cell_color_1strow, font_color_in_column, highlights.
* The Excel variables were all extracted from `excel.py`. It required generalizing the its internal functions, as well as adopting specific Python structures for the variables in `settings.py`.
* The `settings.py` file is intended for use by developers or technical users.

**02/12/2025:**
* Fixed a bug preventing the output directory from being refreshed every time.
* Recover some Excel formatting that was lost in the jupyter-modules translation.

**02/11/2025:**
* The parsing of check-in messages is now applied to all the channels. There is no input variable asking for the name of the check-ins channel anymore (there was a left-over hardcoded checkin_channel_name variable in `messages.py`).
* Added 'missing_value' instead of empty string to empty rows in the URL column. Now all the cells with no information have 'n/d'.

**02/08/2025:**
* GUI now contains visual cues and a help text box to guide the user using the app.
* Changed the checkbuttons for the channels_flag and users_flag to radiobuttons. There is no default value anymore, user have to enter either 'yes' or 'no'.

**2025/01/31:**
* Minor style changes.
* Added channels_json_name, users_json_name, checkin_channel_name, and dest_name_ext to `settings.py`.