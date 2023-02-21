# ResFeci
A Python library to interface with active, in-focus Windows applications (under construction)

The ultimate goal here is to create Python-run alternatives to VBA macros that interface with Windows applications. While there are plenty of apps out there that already 
allow the user to interface with these; very few of these appear to offer the option to easily work with the active, in-focus application 
(i.e. the paths to files must either be hard-coded or selected through a tkinter file selection type prompt).  

The magic that allows my script to work with the in-focus app is the COM component of the pywin32 library (win32com).  For the initial stages, my focus is on Excel. The active worksheet is identified with pywin32; then, the used range of the worksheet is read into a pandas dataframe to allow for various operations.  

As of 2/20/23, I only have a few working (but not entirely tested) Excel features available (Compare Columns; Concatenate Column Row Values and Index & Match). T

This all seems to work in a VERY small test run (n=2 !). In other words, this all DEFINITELY needs a lot more testing! Use as-is and ALWAYS REMEMBER, all macros 
will permanently change your Excel sheet data.  There is no undo! Be sure to save a copy of your report before you run any of this!!!
