# ResFeci

Hi,

This is just a hobby project of mine which seeks to reduce the tedium of manually running varous Excel macros. Instead of the end-user having to mess about with entering formulas and VBA script and whatnot; I intend to present the user with a "low-code, low formula" solution.  More specifically, this application is a tkinter GUI which contains various Python-based alternatives for some useful VBA operations.  With just the press of just a few buttons, the user can easily run what may otherwise be a complex Excel operation.

While there are already plenty of libraries out there that permit the user to interface with Windows applications; very few of them appear to offer
the option to easily work with the active, CURRENTLY-OPEN, application like this one does (provided you are on Windows).  Generally, the paths 
to files must either be hard-coded or selected through something like a tkinter file-selection prompt. Instead, this script is designed to be used with 
an already-open application (note: there are currently only a few Excel functions avaiable at the moment -- more will be added).

The magic that allows my script to work with the in-focus app is the COM component of the pywin32 library (win32com). The active worksheet is identified with pywin32; then, the used range of the worksheet is read into a pandas dataframe to allow for more efficient operations.  

This all seems to work in the few test runs I have conducted.  But, this is very much a work in development! ALWAYS REMEMBER, all macros 
will permanently change your Excel sheet data.  There is no undo (though, this is a planned feature!). Be sure to save a copy of your report before you run any of this!!!

NOTE: 2 External Libraries are Required to Be Installed for this Script: 

1) pywin32  (pip install pywin32)
2) pandas   (pip install pandas)

DISCLAIMER:  THIS IS A CONTINUOUS WORK-IN-PROCESS, WHILE I HAVE NOT COME ACROSS ANY MAJOR ISSUES, I AM JUST A SINGLE HOBBYIST DEVELOPER, AND, AS SUCH, I HAVE LIKELY NOT TESTED FOR EVERY SINGLE SCENARIO.  USAGE OF THIS APPLICATION IS AT-YOUR-OWN-RISK AND ALL DAMAGES REAL OR OTHERWISE ARE HEREBY EXPRESSLY DISCLAIMED UNTIL IF/WHEN PROVEN OTHERWISE IN A COURT OF LAW.
