Update V3.0.0: 
Before this version, users have to manually reinstall the entire .exe file to update the app. With this new version, an auto update script is in place to allow for a more seamless user experience! The app will now update automatically, eliminating the needs for manual uninstalls and installs. Cheers!

My objective of this small project is to design an application to help automate and simplify a report writing process. Conventionally, a typical report takes hours to complete and is very mundane. This program will generate a report in mere seconds, reducing the need for manual counting and typing.

This program will take in an excel sheet filled with data, pre-process it, extract important information and data and finally generate a docx report in a predefined format. Once the program runs, a GUI will appear. The GUI is created using customtkinter, a more modern python GUI library than its predecessor, tkinter.

GUI components in appv2.py:

“Open File” button instantiates a directory (default: program’s current directory) for users to choose the excel file.

“Save To” button instantiates a directory (default: program’s current directory) for users to choose location to save the report.

Textbox displays a status of the chosen excel file as well as the directory to which the report will be saved to.

“Input custom file name..." allows users to input custom file name (default: Production report)

“Generate!” button generates the docx report at a directory of users’ choosing.


Code components in mainv2.py:
In chronological order,

1)	“Generate!” button, when pressed, will call generate_report(). It initialises a timer to record down the time taken for the entire process. 

2)	Using pandas, the excel file is accessed and only columns 1,2,7,27 and 29 (“Status”, “Submission Date”, “Amount”, “Consultant name”, “Branch”) are selected and formed into a dataframe named data. 

3)	data is then iterated through its “Status” column and have “REJECTED” rows removed.

4)	date() is called and returns a list of 2 dates, the earliest and latest date found in the excel sheet. It is saved as a variable named date_output.

5)	A custom filename check is conducted. A default name will be assigned to the report if no custom filename was given by users. It is saved as a variable named filename.

6)	filename will be further given a directory to be saved at.

7)	The report starts to be initialised. A para_header variable adds the header paragraph.

8)	top_producer() is called, creating a global variable total_amount as well as adding the top producer paragraph into the report.

9)	top_amount is then added to para_header¬.

10)	top_case() is called, creating a global variable total_number_of_cases as well as adding the top cases paragraph into the report.

11)	branch() is called, adding the top branches paragraph into the report.

12)	The report ends off by adding the summarised numbers as well as words of encouragement. Report will also be saved and generated.

13)	Timer will stop and the time taken will be calculated. A completion status will also be created. These stats can be seen in the app itself.


