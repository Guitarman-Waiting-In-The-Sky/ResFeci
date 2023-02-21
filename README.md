# ResFeci

A "work in progress" script which allows Windows users to easily run various Python-based alternatives to VBA macros from a GUI (tkinter).

While there are already plenty of libraries out there that permit the user to interface with Windows applications; very few of them appear to offer
the option to easily work with the active, CURRENTLY-OPEN, application like this one does (provided you are on Windows).  Generally, the paths 
to files must either be hard-coded or selected through something like a tkinter file-selection prompt. Instead, this script is designed to be used with 
an already-open application (note: there are currently only a few Excel functions avaiable -- more will be added).

The magic that allows my script to work with the in-focus app is the COM component of the pywin32 library (win32com).  For the initial stages, my focus 
is on Excel. The active worksheet is identified with pywin32; then, the used range of the worksheet is read into a pandas dataframe to allow for more efficient operations.  

As of 2/20/23, I only have a few working (but not entirely tested) Excel features available (Compare Columns; Concatenate Column Row Values and Index & Match). 

This all seems to work in a VERY small test run (n=2 !). In other words, this all DEFINITELY needs a lot more testing! Use as-is and ALWAYS REMEMBER, all macros 
will permanently change your Excel sheet data.  There is no undo! Be sure to save a copy of your report before you run any of this!!!

NOTE: 2 External Libraries are Required to Be Installed for this Script: 

1) pywin32
2) pandas

