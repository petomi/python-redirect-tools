####DESCRIPTION#######
A set of Python-based tools for the creation and testing of redirect rules (both IIS and .htaccess formats) from an Excel sheet.


####INSTRUCTIONS######
1. Install python 3 (including pip)
2. In Windows Command Prompt/bash shell, navigate to the directory where redirect-tools.py is located and run "python setup.py"
2. Save Excel file as .xls (.xlsx is not supported yet) in the same directory as redirect-tools.py.
3. Change config settings in settings.cfg to suit your project needs.
4. Run script by typing "python redirect-tools.py" into Windows Command Prompt/bash shell, followed by a command ("test", "create_map", "create_htaccess", "create_rules".
5. The file will be created as [input_file]-OUT.xls
6. If needed, re-institute filters at top of each column in Excel and save as .xlsx
7. Repeat steps 1-6 for each sheet you need to update.
8. Be happy you didn't have to do it all by hand. :)
