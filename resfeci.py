import win32com.client as win32 # This allows active, in focus Excel sheet to be read
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox  # Python 3
from tkinter.simpledialog import askstring

# Converts Excel Sheet into DataFrame
import pandas as pd

'''
The ultimate goal here is to create Python-run alternatives to VBA macros that interface with Windows applications. While there are plenty of apps out there that already allow the user to
interface with these; very few of these appear to offer the option to work with the active, in-focus application (i.e. the paths to files must either be hard-coded or selected 
through a tkinter file selection type prompt).  The magic that allows this to happen is the win32com library.

As of 2/19/23, I only have a couple of working (but not entirely tested) Excel features available (Compare Columns and Concatenate Column Row Values). The active worksheet is identified with win32; then, the
used range of the worksheet is read into a pandas dataframe to allow for various operations.  

2/20/23: Converted Index and Match function into the "active worksheet" model.  It seems to work in a VERY small test run (n=2 :) )---will be tested more. 

This all DEFINITELY needs a lot more testing! Use as-is!!! 

REMEMBER, all macros are permanent (i.e. be sure to save a copy of your report before you run any of this!!!)


'''

#class Excel: # initially set up as a classed-library; but, this is really just a tkinter app at the moment.  Leave here just in case ever needed again.

# Working with the ALREADY-OPEN, active, Excel sheet to create an easier way to run Excel's Index and Match
# TO-DO
# () Test Function
# () Currenly, strings can be read as numbers (i.e. 3 can become 3.0 when carried over)--look into retaining data type.

def index_match():

        excel = win32.gencache.EnsureDispatch('Excel.Application') # Opens application

        try:

                def return_main_menu():
                    root.destroy()
                    main_menu()

                def run_index_match():

                    col1=options_list.index(value_col1.get())
                    col2=options_list.index(value_col2.get())
                    col3=options_list.index(value_col3.get())
                    col4=options_list.index(value_col4.get()) # this should be an INT to work with VBA (win32 writeData.cells)

                # Pandas indexing counts start after header row (and starts at 0); because of this first iteration should be index + 2

                    row_counter=1
                    for i in df.itertuples(): # iterating the chosen column
                        df2 = df[df.iloc[:,col2]==str(i[col1+1])] # creating smaller dataframe where Col2 value = current iteration of Col1 value
                        if len(df2)>0: # it a hit...
                            corrsponding_values=df2.iloc[:,col3].unique()  #...create a unique list of values of the corresponding Col3 data
                            corrsponding_values=corrsponding_values.tolist() #...converting to list
                            writeData.Cells(row_counter, col4+1).Value = str(corrsponding_values) # writing that list to selected output column (col4)
                        row_counter=row_counter+1
                    excel.Visible = True # renders the app visible
                    return None

                root = tk.Tk()
                root.title("Index and Match\t\t")
                # center root window
                root.tk.eval(f'tk::PlaceWindow {root._w} center')
                writeData = excel.ActiveSheet

                df = pd.DataFrame(writeData.UsedRange())    # Creates a pandas dataframe out of the used range of the active excel sheet.  
                df.columns=df.iloc[0]                       # df is created w/o headers, this coverts first row into a numpy.ndarray to be our headers
                df = df.fillna('')                          

                headers=df.columns.tolist()
                df.columns = df.columns.str.replace('.', '_')

                counter=1

                # Necessary to create a 2nd list because the 1st "headers" may contain duplicately named columns which breaks index()
                # This ensures clears up confusion by adding ("Column: ")
                headers2=[]
                
                for x in headers:
                    headers2.append(str(x) + ' (Column: ' + str(counter) + ')')
                    counter=counter+1

                # Create the list of options
                options_list = headers2
        

                header_dict={}
                header_counter=1
                for names in headers:
                    header_dict.update({names : header_counter})
                    header_counter=header_counter+1
            
                # Set the default value of the variable
                value_col1 = tk.StringVar(root)
                value_col1.set("  Select the 1st Column to Compare  ")
                col1_prompt= tk.OptionMenu(root, value_col1, *options_list)
                col1_prompt.pack(pady=10, padx=100)

                # Set the default value of the variable
                value_col2 = tk.StringVar(root)
                value_col2.set("  Select the Column to be Compared Against the 1st Column  ")
                col2_prompt= tk.OptionMenu(root, value_col2, *options_list)
                col2_prompt.pack(pady=10, padx=100)

                # Set the default value of the variable
                value_col3 = tk.StringVar(root)
                value_col3.set("  Which Column Contains the Data to be Carried Over  ")
                col3_prompt= tk.OptionMenu(root, value_col3, *options_list)
                col3_prompt.pack(pady=10, padx=100)

                # Set the default value of the variable
                value_col4 = tk.StringVar(root)
                value_col4.set("  Select Column to Place the Extracted Data In  ")
                col4_prompt= tk.OptionMenu(root, value_col4, *options_list)
                col4_prompt.pack(pady=10, padx=100)

                submit_button = tk.Button(root, text='Submit', command=run_index_match)
                submit_button.pack(pady=10)

                main_menu_button = tk.Button(root, text='Return to Main', command=return_main_menu)
                main_menu_button.pack(pady=10)

                root.mainloop()
                        
                excel.Visible = True # renders the app visible

        except Exception as e:
            print(e)
            turn_excel_back_on(excel)
        


def concatenate_column_values_active():

        excel = win32.gencache.EnsureDispatch('Excel.Application') # Opens application

        try:

            def return_main_menu():
                root.destroy()
                main_menu()

            def concatenate():
                col1=options_list.index(value_col1.get())
                deliimter=delimiter_prompt.get()
                delimited_output_string=''
                for i in df.itertuples(): # itertuples for speed; think of ways to vectorize...
                    # if Col1 row value isn't header and isn't blank, then see if value is in Col2 numpy array, if so, highlight the cell.
                    if delimited_output_string=='':
                        delimited_output_string=str(i[col1+1]) + str(deliimter)
                    else:
                        delimited_output_string=delimited_output_string +  str(i[col1+1]) + str(deliimter)
                        
                # removing last character from string (non-used delimiter)        
                delimited_output_string = delimited_output_string.rstrip(delimited_output_string[-1])
                show_delimited_string(delimited_output_string)

            def show_delimited_string(delimited_output_string):
                results = tk.Tk()
                results.geometry("800x600+900+400")
                results.title("Results\t\t")
                # center root window
                #results.tk.eval(f'tk::PlaceWindow {results._w} center')
                # center root window
                w = tk.Text(results, height=30, font=14, wrap='word')
                w.insert(1.0, '\n' + delimited_output_string + '\n')
                w.pack()
                w.configure(inactiveselectbackground=w.cget("selectbackground"))
                w.configure(state="disabled")
                results.mainloop()
                #messagebox.showinfo("showinfo", delimited_output_string)



            writeData = excel.ActiveSheet
            df = pd.DataFrame(writeData.UsedRange())    # Creates a pandas dataframe out of the used range of the active excel sheet.  
            df.columns=df.iloc[0]                       # df is created w/o headers, this coverts first row into a numpy.ndarray to be our headers
            df = df.fillna('')                          

            headers=df.columns.tolist()
            counter=1

            # Necessary to create a 2nd list because the 1st "headers" may contain duplicately named columns which breaks index()
            # This ensures clears up confusion by adding ("Column: ")
            headers2=[]
            
            for x in headers:
                headers2.append(str(x) + ' (Column: ' + str(counter) + ')')
                counter=counter+1

            # Create the list of options
            options_list = headers2
            
            root = tk.Tk()
            root.title("Concatenate\t\t")
            # center root window
            root.tk.eval(f'tk::PlaceWindow {root._w} center')
            #root.withdraw()
           
            # Set the default value of the variable
            value_col1 = tk.StringVar(root)
            value_col1.set("  Select the Column to Concatenate  ")
            col_prompt= tk.OptionMenu(root, value_col1, *options_list)
            col_prompt.pack(pady=10, padx=100)

            # Set the default value of the variable
            value_delimiter = tk.StringVar(root)
            #value_delimiter.insert("  Enter Your Delimiter  ")
            delimiter_label=tk.Label(root, text = 'Enter Your Delimiter')
            delimiter_label.pack()
            delimiter_prompt= tk.Entry(root, textvariable=value_delimiter, justify='center')
            #delimiter_prompt.insert(0, ' Enter Your Delimiter ')
            delimiter_prompt.pack(pady=10)

            submit_button = tk.Button(root, text='Submit', command=concatenate)
            submit_button.pack(pady=10)

            main_menu_button = tk.Button(root, text='Return to Main', command=return_main_menu)
            main_menu_button.pack(pady=10)

            root.mainloop()


        except Exception as e:
            print(e)
            turn_excel_back_on(excel_object=excel)
    

def compare_columns_active():

        '''
        This compares two user-selected columns (from the ACTIVE) Excel Worksheet.  If there are any simliar values between the columns,
        then, thhe cells are highlighted.
        
        The function EnsureDispatch() in win32.client.gencache allows you specify a prog_id and the gen_py cache wrapper objects are created at 
        runtime if they don't already exist. This is useful if you don't care what version of COM server you use, allowing users to have various 
        versions and still work with your code. In other words, it is the secret sauce which renders all this possible and grabs the in-foucs 
        worksheet with win32
        '''

        excel = win32.gencache.EnsureDispatch('Excel.Application') # Opens application
        try:
            
            def return_main_menu():
                root.destroy()
                main_menu()

            def run_comparison():
                col1=options_list.index(value_col1.get())
                col2=options_list.index(value_col2.get())

                # Reading selected column row values into lists (which will be compared soon)
                # Note: ':' selects all rows from the chosen column (index based column chosen with INT as col1, col2)
                col1_values=df.iloc[:,col1].values
                col2_values=df.iloc[:,col2].values

                # Turning off Excel screen updating/calculations which can vastly reduce processing time.
                # Just remember to turn it back on when done (or in event of failure) or you will break Excel!
                excel.ScreenUpdating=False
                excel.Application.Calculation = -4135 # to set xlCalculationManual
                
                '''
                Now to iterate through the selected columns.  Note: our main datasource is a pandas dataframe.  As such, i[0] is the INDEX value count of the row.
                Because of this, there is a +1 offset between the header row and the pandas data frame.
                As such, it is necessary to add +1 to the col1 and col2 values in order to synchronize the two.

                Below you will see two iterations (to compare the 2 diffrent numpy arrays).  I considered just joining the arrays so that only 1 iteration is 
                needed, but, this doesn't work as every value in both columns would be highlighted (because they'd all be in the same array)
                '''

                # Comparing Col1 to Col2 numpy array
                row_counter=1
                for i in df.itertuples(): # itertuples for speed; think of ways to vectorize...
                    # if Col1 row value isn't header and isn't blank, then see if value is in Col2 numpy array, if so, highlight the cell.
                    if i[col1+1] in col2_values and row_counter!=1 and i[col1+1]!='': 
                        writeData.Cells(row_counter, col1+1).Interior.ColorIndex = 37
                    row_counter=row_counter+1

                # Repeating the Process to Compare Col2 to Col1 numpy array
                row_counter=1
                for i in df.itertuples(): 
                    if i[col2+1] in col1_values and row_counter!=1 and i[col2+1]!='': 
                        writeData.Cells(row_counter, col2+1).Interior.ColorIndex = 37
                    row_counter=row_counter+1

                # All Done---Turning Excel back On
                excel.ScreenUpdating=True
                excel.Application.Calculation = -4105 # to set xlCalculationManual
                #root.destroy()  # remove the tkinter when when done---or leave open for more

                return None

            writeData = excel.ActiveSheet
            df = pd.DataFrame(writeData.UsedRange())    # Creates a pandas dataframe out of the used range of the active excel sheet.  
            df.columns=df.iloc[0]                       # df is created w/o headers, this coverts first row into a numpy.ndarray to be our headers
            df = df.fillna('')                          # replacing NaN with nothing because I dont like Nan (not to be confused with naan which is amazing)
           
            headers=df.columns.tolist()
            counter=1

            # Necessary to create a 2nd list because the 1st "headers" may contain duplicately named columns which breaks index()
            # This ensures clears up confusion by adding ("Column: ")
            headers2=[]
            
            for x in headers:
                headers2.append(str(x) + ' (Column: ' + str(counter) + ')')
                counter=counter+1

            # Create the list of options
            options_list = headers2
            
            root = tk.Tk()
            root.title("Compare Columns\t\t")
            # center root window
            root.tk.eval(f'tk::PlaceWindow {root._w} center')
            #root.withdraw()
           
            # Set the default value of the variable
            value_col1 = tk.StringVar(root)
            value_col1.set("  Select 1st Column to Compare  ")
            col1_prompt= tk.OptionMenu(root, value_col1, *options_list)
            col1_prompt.pack(pady=10, padx=100)

             # Set the default value of the variable
            value_col2 = tk.StringVar(root)
            value_col2.set("  Select 2nd Column to Compare  ")
            col2_prompt= tk.OptionMenu(root, value_col2, *options_list)
            col2_prompt.pack(pady=10)

            submit_button = tk.Button(root, text='Submit', command=run_comparison)
            submit_button.pack(pady=10)

            main_menu_button = tk.Button(root, text='Return to Main', command=return_main_menu)
            main_menu_button.pack(pady=10)

            root.mainloop()


        except Exception as e:
            print(e)
            turn_excel_back_on(excel_object=excel)


def turn_excel_back_on(excel_object):
        excel_object.ScreenUpdating=True
        excel_object.Application.Calculation = -4105 # to set xlCalculationManual


def main_menu():

        '''
        This is just the main options menu. Pretty basic and just serves to prompt the user for their desired action.
        This is all running off of Tkinter because its part of the Python standard library and is cross-platform.
        '''

        try:

            def main_menu_choice():

                choice=selection.get()

                if choice=='Compare Two Columns':
                    root.destroy()
                    compare_columns_active()
                elif choice=='Concatenate Column Values':
                    root.destroy()
                    concatenate_column_values_active()

                elif choice=='Index and Match':
                    root.destroy()
                    index_match()
                    
                else:
                    root.destroy()
                    main_menu()
        
                return None

            # Create the list of options
            options_list = ['Compare Two Columns', 'Concatenate Column Values', 'Index and Match']
            
            root = tk.Tk()
            root.title("Res Feci\t\t")
            # center root window
            root.tk.eval(f'tk::PlaceWindow {root._w} center')
           
            # Set the default value of the variable
            selection = tk.StringVar(root)
            selection.set("  What Would You Like To Do?  ")
            selection_prompt= tk.OptionMenu(root, selection, *options_list)
            selection_prompt.pack(pady=10, padx=100)

            submit_button = tk.Button(root, text='Submit', command=main_menu_choice)
            submit_button.pack(pady=10)

            root.mainloop()


        except Exception as e:
            print(e)


if __name__=='__main__':
    main_menu()
