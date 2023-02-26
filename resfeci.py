import win32com.client as win32 # This allows active, in focus Excel sheet to be read
import tkinter as tk
import pandas as pd


'''

The ultimate goal here is to create Python-run alternatives to VBA macros that interface with Windows applications. While there are plenty of apps out there that already allow the user to
interface with these; very few of these appear to offer the option to work with the active, in-focus application (i.e. the paths to files must either be hard-coded or selected 
through a tkinter file selection type prompt).  The magic that allows this to happen is the pywin32 library (specifically, its COM component---win32com).

This all DEFINITELY needs a lot more testing! Use as-is!!! 

REMEMBER, all macros are permanent (i.e. be sure to save a copy of your report before you run any of this!!!)

TO-DO LIST

 () Build an UNDO feature (need to save the state of the workbook/worksheet) prior to running functions (how>? VBA copy format/paste special?)
 () Build a function to compare differences in worksheets (should be easy wtith Pandas....)

 '''


# Working with the ALREADY-OPEN, active, Excel sheet to create an easier way to run Excel's Index and Match
# TO-DO
# () Finish building multi-column concatenate
# () Test Function
# () Currenly, strings can be read as numbers (i.e. 3 can become 3.0 when carried over)--look into retaining data type.
        # This is a known limitation of using Pandas to reac Excel.  Text can become floats.  The solution is to specity the specific datatype of each column, but this requires knowing the column names....


activeSheet_initial=''

def main_menu():

        '''
        This is just the main options menu. Pretty basic and just serves to prompt the user for their desired action.
        This is all running off of Tkinter because its part of the Python standard library and is cross-platform.
        '''

        try:

            # Create the list of main menu options
            options_list = ['Compare Two Columns', 'Concatenate SINGLE Column Values', 'Index and Match', 'Concatenate MULTIPLE Column Values']
            
            root = tk.Tk()
            root.title("Res Feci\t\t")
            # center root window
            root.tk.eval(f'tk::PlaceWindow {root._w} center')

            def main_menu_choice():

                choice=selection.get()

                if choice=='Compare Two Columns':
                    root.destroy()
                    compare_columns_active(message='standard',message2='initial_run')
                elif choice=='Concatenate SINGLE Column Values':
                    root.destroy()
                    concatenate_column_values_active(message='standard')
                elif choice=='Concatenate MULTIPLE Column Values':
                    root.destroy()
                    concatenate_multiple_column_values_active(message='standard')
                elif choice=='Index and Match':
                    root.destroy()
                    index_match()
                    
                else:
                    root.destroy()
                    main_menu()
        
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


'''

This is the fly in the ointment. The problem is this...

"Excel stores numbers differently that you may have them formatted display on the worksheet. Under normal circumstances, Excel stores numeric values 
as "Double Precision Floating Point" numbers, or "Doubles" for short."

    -from Rounding Errors In Excel by Chip Pearson, 27-Oct-1998.

Long story short; this means that when Pandas reads the Excel.ActiveSheet as a dataframe, any #'s are read as float.  A displayed value of '3', will
be read into Pandas as its floating point equivalent '3.0'.  For most operations, this is not ideal, and most users will want to preserved the "displayed"
value, not the actual value.  This function reads a column and converts any floats that end in .0 into INT. 

This can result is mixed datatypes in a single columns.  For example, if you have '3.14' and '5' in a columns, then, 3.14 will remain FLOAT while 5 will 
be converted to INT.  NEED TO CONSIDER THE DOWNSIDE OF THIS.  AT THE MOMENT, THERE ARE NO DOWNSIDES BECAUSE I'M CONVERTING EVERYTHING INTO STRINGS ANYWAYS. 
BUT, HEADS UP IF I CONSIDER MATHEMATICAL OPERATIONS.

'''

def create_df_from_activesheet_convert_float_to_int(active_sheet_object):

    used_range_row_tuple=()
    used_range_master_tuple=()

    for x in active_sheet_object.UsedRange():
        used_range_row_tuple=()
        for y in x:
            if str(type(y))=="<class 'float'>" and str(y)[-2:]=='.0':
                y=str(y)
                y=y[:-2]
                y=int(y)
                used_range_row_tuple=(*used_range_row_tuple, y)
            else:
                used_range_row_tuple=(*used_range_row_tuple, y)
        used_range_master_tuple=(*used_range_master_tuple, used_range_row_tuple)
    
    df = pd.DataFrame(used_range_master_tuple)
    df.columns=df.iloc[0] # df is created w/o headers, this coverts first row into a numpy.ndarray to be our headers
    df = df.fillna('')   
    return df


'''

The tkinter GUI presents user options in the form of a drop-down options menu.  Each option corresponds to an an Excel column header.  This reads the
header values and converts those into the user-presented list for the options menu.

'''

def create_options_list_for_activesheet(df):

    # Reading column headers into an initial list; this list is the headers "as-displayed" on the Excel sheet
    displayed_headers=df.columns.tolist()
  
    # It is necessary to create a 2nd list because the 1st "headers" list may contain duplicately named columns which breaks my name-based indexing df system.
    # My solution to this is to make sure all columns have unique index names by adding ("Column: COUNTER+ ") to the displayed headers.

    options_list=[]

    counter=1
            
    for x in displayed_headers:
        options_list.append(str(x) + ' (Column: ' + str(counter) + ')')
        counter=counter+1

    # Adding one final column to the options menu, the is the column after the last-used column.  This is useful for asking the user if they want to place data into THIS column.
    options_list.append('BLANK' + ' (Column: ' + str(counter) + ')')

    # Returning options list!
    return options_list



def concatenate_multiple_column_values_active(**message):

    root = tk.Tk()
    root.title("Concatenate MULTIPLE Columns\t\t")
    root.tk.eval(f'tk::PlaceWindow {root._w} center')

    def return_main_menu():
        root.destroy()
        main_menu()

    def concatenate_columns():

        # Ensuring all pull-down menus have values---failure if nothing selected

        try:
            col1=options_list.index(value_col1.get())
        except:
            root.destroy()
            concatenate_multiple_column_values_active(message='unselected_1st_column')

        try:
            col2=options_list.index(value_col2.get())
        except:
            root.destroy()
            concatenate_multiple_column_values_active(message='unselected_column')

        try:
            output_col=options_list.index(value_output_col.get())
        except:
            root.destroy()
            concatenate_multiple_column_values_active(message='unselected_output_column')

        col1=options_list.index(value_col1.get())
        col2=options_list.index(value_col2.get())

        # User did not select col3 (this is caught by the error checking above)
        try:
            col3=options_list.index(value_col3.get())
        except:
            col3=-1

        output_col=options_list.index(value_output_col.get())

        deliimter=str(delimiter_prompt.get())

        row_counter=1


        for i in df.itertuples(): # itertuples for speed; think of ways to vectorize...
            if row_counter!=1:
                if str(i[col1+1])=='' and str(i[col2+1])=='' and str(i[col3+1])=='': # all blank
                    activeSheet.Cells(row_counter, output_col+1).Value = ''
                else:
                    if col3==-1: # col3 not selected
                        if str(i[col1+1])=='' and str(i[col2+1])=='':
                            activeSheet.Cells(row_counter, output_col+1).Value = ''
                        elif str(i[col1+1])!='' and str(i[col2+1])!='':
                            activeSheet.Cells(row_counter, output_col+1).Value = f'{str(i[col1+1])}{deliimter}{str(i[col2+1])}'
                        elif str(i[col1+1])!='' and str(i[col2+1])=='':
                            activeSheet.Cells(row_counter, output_col+1).Value = f'{str(i[col1+1])}'
                        else:
                            activeSheet.Cells(row_counter, output_col+1).Value = f'{str(i[col2+1])}'
                    else: # 3rd column selected
                        if str(i[col1+1])!='' and str(i[col2+1])!='' and str(i[col3+1])!='':
                            activeSheet.Cells(row_counter, output_col+1).Value = f'{str(i[col1+1])}{deliimter}{str(i[col2+1])}{deliimter}{str(i[col3+1])}'
                        elif str(i[col1+1])!='' and str(i[col2+1])=='' and str(i[col3+1])!='':
                            activeSheet.Cells(row_counter, output_col+1).Value = f'{str(i[col1+1])}{deliimter}{str(i[col3+1])}'
                        elif str(i[col1+1])!='' and str(i[col2+1])!='' and str(i[col3+1])=='':
                            activeSheet.Cells(row_counter, output_col+1).Value = f'{str(i[col1+1])}{deliimter}{str(i[col2+1])}'
                        elif str(i[col1+1])=='' and str(i[col2+1])!='' and str(i[col3+1])!='':
                            activeSheet.Cells(row_counter, output_col+1).Value = f'{str(i[col2+1])}{deliimter}{str(i[col3+1])}'
                        elif str(i[col1+1])=='' and str(i[col2+1])=='' and str(i[col3+1])!='':
                            activeSheet.Cells(row_counter, output_col+1).Value = f'{str(i[col3+1])}'
                        elif str(i[col1+1])=='' and str(i[col2+1])!='' and str(i[col3+1])=='':
                            activeSheet.Cells(row_counter, output_col+1).Value = f'{str(i[col2+1])}'
                        elif str(i[col1+1])!='' and str(i[col2+1])=='' and str(i[col3+1])=='':
                            activeSheet.Cells(row_counter, output_col+1).Value = f'{str(i[col1+1])}'

            row_counter=row_counter+1

    excel = win32.gencache.EnsureDispatch('Excel.Application') # Opens application
    activeSheet = excel.ActiveSheet
    if activeSheet is None:
        del excel
        root.destroy()
        error_window(message='excel_not_open')
    
    df = create_df_from_activesheet_convert_float_to_int(activeSheet)   
    options_list = create_options_list_for_activesheet(df)

    for x in message.values():
        if x == 'unselected_column':
            message_label=tk.Label(root, text = 'Please Ensure the 2nd Column is Selected')
        elif x == 'unselected_1st_column':
            message_label=tk.Label(root, text = 'Please Ensure the 1st Column Value is Selected')
        elif x == 'unselected_output_column':
            message_label=tk.Label(root, text = 'Please Be Sure to Select an Output Column!')
        else:
            message_label=tk.Label(root, text = '')
            
    message_label.pack()
            
    value_col1 = tk.StringVar(root)
    value_col1.set("  Select the 1st Column to Join  ")
    col_prompt= tk.OptionMenu(root, value_col1, *options_list)
    col_prompt.pack(pady=10, padx=100)

    value_col2 = tk.StringVar(root)
    value_col2.set("  Select the Corresponding 2nd Column to Join  ")
    col2_prompt= tk.OptionMenu(root, value_col2, *options_list)
    col2_prompt.pack(pady=10, padx=100)

    value_col3 = tk.StringVar(root)
    value_col3.set("  Select the Corresponding 3nd Column to Join (leave as-is if NONE)  ")
    col3_prompt= tk.OptionMenu(root, value_col3, *options_list)
    col3_prompt.pack(pady=10, padx=100)

    value_output_col = tk.StringVar(root)
    value_output_col.set("  Select the Column to Hold Joined Data  ")
    output_col_prompt= tk.OptionMenu(root, value_output_col, *options_list)
    output_col_prompt.pack(pady=10, padx=100)

    value_delimiter = tk.StringVar(root)
    delimiter_label=tk.Label(root, text = 'If Desired, Enter the Delimiter (Leave Blank for None)')
    delimiter_label.pack()
    delimiter_prompt= tk.Entry(root, textvariable=value_delimiter, justify='center')
    delimiter_prompt.pack(pady=10)

    submit_button = tk.Button(root, text='Submit', command=concatenate_columns)
    submit_button.pack(pady=10)

    main_menu_button = tk.Button(root, text='Return to Main', command=return_main_menu)
    main_menu_button.pack(pady=10)

    root.mainloop()

 

def index_match():

        root = tk.Tk()
        root.title("Index and Match\t\t")
        root.tk.eval(f'tk::PlaceWindow {root._w} center')

        try:

                def return_main_menu():
                    root.destroy()
                    main_menu()

                def run_index_match():

                    col1=options_list.index(value_col1.get())
                    col2=options_list.index(value_col2.get())
                    col3=options_list.index(value_col3.get())
                    col4=options_list.index(value_col4.get()) # this should be an INT to work with VBA (win32 activeSheet.cells)

                # Pandas indexing counts start after header row (and starts at 0); because of this first iteration should be index + 2

                    row_counter=1
                    for i in df.itertuples(): # iterating the chosen column
                        if str(i[col1+1])!='':
                            df2 = df[df.iloc[:,col2]==str(i[col1+1])] # creating smaller dataframe where Col2 value = current iteration of Col1 value
                            if len(df2)>0: # it a hit...
                                corrsponding_values=df2.iloc[:,col3].unique()  #...create a unique list of values of the corresponding Col3 data
                                corrsponding_values=corrsponding_values.tolist() #...converting to list
                                output_string='' # converting string to list for readability on output
                                for x in corrsponding_values:
                                    if output_string=='':
                                        output_string=str(x) + '|'
                                    else:
                                        output_string=output_string + str(x) + '|'
                                output_string = output_string.rstrip(output_string[-1])

                                activeSheet.Cells(row_counter, col4+1).Value = output_string # writing that list to selected output column (col4)
                        row_counter=row_counter+1
                    excel.Visible = True # renders the app visible
                    

                excel = win32.gencache.EnsureDispatch('Excel.Application') # Opens application
                activeSheet = excel.ActiveSheet

                if activeSheet is None:
                    del excel
                    root.destroy()
                    error_window(message='excel_not_open')


                df = create_df_from_activesheet_convert_float_to_int(activeSheet)   
                options_list = create_options_list_for_activesheet(df)
                    
                # Set the default value of the variable
                value_col1 = tk.StringVar(root)
                value_col1.set("  Step 1: Select the 1st Column to Be Compared  ")
                col1_prompt= tk.OptionMenu(root, value_col1, *options_list)
                col1_prompt.pack(pady=10, padx=100)

                # Set the default value of the variable
                value_col2 = tk.StringVar(root)
                value_col2.set("  Step 2: Select the Column to be Compared Against the Step 1 Column  ")
                col2_prompt= tk.OptionMenu(root, value_col2, *options_list)
                col2_prompt.pack(pady=10, padx=100)

                # Set the default value of the variable
                value_col3 = tk.StringVar(root)
                value_col3.set("  Step 3: Which Column Corresponds with the Step 2 Column and Contains the Data to Copy  ")
                col3_prompt= tk.OptionMenu(root, value_col3, *options_list)
                col3_prompt.pack(pady=10, padx=100)

                # Set the default value of the variable
                value_col4 = tk.StringVar(root)
                value_col4.set("  Step 4: Select Final Column to Place the Copied Step 3 Data Into  ")
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
        

def concatenate_column_values_active(**message):

        root = tk.Tk()
        root.title("Concatenate\t\t")
        root.tk.eval(f'tk::PlaceWindow {root._w} center')

        try:

            def return_main_menu():
                root.destroy()
                main_menu()

            def concatenate():

                # Failure if users does not select a column
                try:
                    col1=options_list.index(value_col1.get())
                except:
                    root.destroy()
                    concatenate_column_values_active(message='unselected_column')

                delimiter=delimiter_prompt.get()
                delimited_output_string=''
                for i in df.itertuples(): # itertuples for speed; think of ways to vectorize...
                    if str(i[col1+1])!='':
                        if delimited_output_string=='':
                            delimited_output_string=str(i[col1+1]) + str(delimiter)
                        else:
                            delimited_output_string=delimited_output_string +  str(i[col1+1]) + str(delimiter)
                        
                # removing last occurence of delimiter from string (if delimiter isn't blank)  
                if delimiter!='':
                    delimited_output_string = delimited_output_string[:-len(delimiter)]
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
                # The next two lines allow user-selectable text to be generated (in case they want to copy and paste)
                w.configure(inactiveselectbackground=w.cget("selectbackground"))
                w.configure(state="disabled")
                copy_button = tk.Button(results, text='COPY TO CLIPBOARD', command=lambda: w.clipboard_append(delimited_output_string))
                copy_button.pack(pady=10)
                results.mainloop()

            
            excel = win32.gencache.EnsureDispatch('Excel.Application') # Opens application
            activeSheet = excel.ActiveSheet
            if activeSheet is None:
                del excel
                root.destroy()
                error_window(message='excel_not_open')

            df = create_df_from_activesheet_convert_float_to_int(activeSheet)   
            options_list = create_options_list_for_activesheet(df)
      
            for x in message.values():
                if x == 'unselected_column':
                    message_label=tk.Label(root, text = 'Please Select a Column')
                else:
                    message_label=tk.Label(root, text = '')
            
            message_label.pack()
            
            # Set the default value of the variable
            value_col1 = tk.StringVar(root)
            value_col1.set("  Select the Column to Concatenate  ")
            col_prompt= tk.OptionMenu(root, value_col1, *options_list)
            col_prompt.pack(pady=10, padx=100)

            # Set the default value of the variable
            value_delimiter = tk.StringVar(root)
            delimiter_label=tk.Label(root, text = 'If Desired, Enter the Delimiter (Leave Blank for None)')
            delimiter_label.pack()
            delimiter_prompt= tk.Entry(root, textvariable=value_delimiter, justify='center')
            delimiter_prompt.pack(pady=10)

            submit_button = tk.Button(root, text='Submit', command=concatenate)
            submit_button.pack(pady=10)

            main_menu_button = tk.Button(root, text='Return to Main', command=return_main_menu)
            main_menu_button.pack(pady=10)

            root.mainloop()


        except Exception as e:
            print(e)
            turn_excel_back_on(excel_object=excel)


def active_excel_job_complete(job):

    try:

        def return_main_menu():
            job_compete_window.destroy()
            main_menu()

        def job_route(job):
            print(job)
            job_compete_window.destroy()
            if job=='comparison':
                compare_columns_active(message='subsequent_run')

        

        job_compete_window = tk.Tk()
        job_compete_window.title("Job Complete\t\t")
            # center root window
        job_compete_window.tk.eval(f'tk::PlaceWindow {job_compete_window._w} center')
            #root.withdraw()

        compare_label=tk.Label(job_compete_window, text = 'Job Complete!\n\nHit BACK to run a simliar operation\nor select RETURN TO MAIN')
        compare_label.pack(padx=100)


        if job=='comparison':
            submit_button = tk.Button(job_compete_window, text='BACK', command= lambda: job_route(job='comparison'))
        submit_button.pack(pady=10)

        main_menu_button = tk.Button(job_compete_window, text='RETURN TO MAIN', command=return_main_menu)
        main_menu_button.pack(pady=10)

        job_compete_window.mainloop()

    except Exception as e:
        print(e)
       # turn_excel_back_on()


def error_window(**message):

    def return_main_menu():
                root.destroy()
                main_menu()

    root = tk.Tk()
    root.title("ERROR!!\t\t")
    root.tk.eval(f'tk::PlaceWindow {root._w} center')

    error_label=tk.Label(root)

    for x in message.values():
                if x == 'excel_not_open':
                    error_label=tk.Label(root, text = 'Active Excel Sheet Not Detected\nPlease Open Excel and Try Again')

    error_label.pack(padx=30)

    main_menu_button = tk.Button(root, text='Return to Main', command=return_main_menu)
    main_menu_button.pack(pady=10)

    root.mainloop()




def compare_columns_active(**kwargs):

        '''
        This compares two user-selected columns (from the ACTIVE) Excel Worksheet.  If there are any simliar values between the columns,
        then, the cells are highlighted.
        '''

        try:

            root = tk.Tk()
            root.title("Compare Columns\t\t")
            # center root window
            root.tk.eval(f'tk::PlaceWindow {root._w} center')

            def return_main_menu():
                root.destroy()
                main_menu()

            def run_comparison():

                # Error is thrown if user does not select a Column before hitting SUBMIT.
                try:
                    col1=options_list.index(value_col1.get())
                    col2=options_list.index(value_col2.get())
                except:
                    root.destroy()
                    compare_columns_active(message='unselected_column')

                if col1==col2: # if user selects the same column ask them to try, function should have a message.
                    root.destroy()
                    compare_columns_active(message='same_column')

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
                        activeSheet.Cells(row_counter, col1+1).Interior.ColorIndex = 37
                    row_counter=row_counter+1

                # Repeating the Process to Compare Col2 to Col1 numpy array
                row_counter=1
                for i in df.itertuples(): 
                    if i[col2+1] in col1_values and row_counter!=1 and i[col2+1]!='': 
                        activeSheet.Cells(row_counter, col2+1).Interior.ColorIndex = 37
                    row_counter=row_counter+1

                # All Done---Turning Excel back On
                excel.ScreenUpdating=True
                excel.Application.Calculation = -4105 # to set xlCalculationManual
                #root.destroy()  # remove the tkinter when when done---or leave open for more
                root.destroy()
                active_excel_job_complete(job='comparison')

            # Planned feature to allow macros to be undone.  This seems to work; but cell interior color is not preserved (values only...) 
            def undo(inital_active_sheet):
                excel.ActiveSheet.UsedRange.Value = inital_active_sheet
                inital_active_sheet=''
                root.destroy()
                return compare_columns_active(message='initial_run')


            excel = win32.gencache.EnsureDispatch('Excel.Application') # Opens application     

            activeSheet = excel.ActiveSheet

            global activeSheet_initial # testing out UNDO feature; remove this is undo doesnt panout.
            global initial_active_sheet # testing out UNDO feature; remove this is undo doesnt panout.

            for x in kwargs.values():
                if x =='initial_run':
                    activeSheet_initial = excel.ActiveSheet.UsedRange()
                    initial_active_sheet=activeSheet_initial

            if activeSheet is None:
                del excel
                root.destroy()
                error_window(message='excel_not_open')           

            df = create_df_from_activesheet_convert_float_to_int(activeSheet)   
            options_list = create_options_list_for_activesheet(df)
        
            compare_label=tk.Label(root)

            # Message is the kwarg parameter of this function.
            for x in kwargs.values():
                if x == 'same_column':
                    compare_label=tk.Label(root, text = 'Please Select More Than 1 Column')
                elif x == 'unselected_column':
                    compare_label=tk.Label(root, text = 'Please Select Both Columns')
                elif x == 'excel_not_open':
                    compare_label=tk.Label(root, text = 'Excel Worksheet Not Detected, Please Load and Try Again')
                else:
                    compare_label=tk.Label(root, text = 'Select the Two Columns to Be Compared\n\nSimilar Values Between the Columns Will be Highlighted')

            compare_label.pack()
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

           # submit_button = tk.Button(root, text='Undo', command= lambda: undo(initial_active_sheet))
           # submit_button.pack(pady=10)

            root.mainloop()

        except Exception as e:
            print(e)
            turn_excel_back_on(excel_object=excel)


def turn_excel_back_on(excel_object):
        excel_object.ScreenUpdating=True
        excel_object.Application.Calculation = -4105 # to set xlCalculationManual



'''
SCRATCH NOTES

This may be useful someday (means of referencing the full path to the currently open workbook/worksheet):

    # excel = win32.gencache.EnsureDispatch('Excel.Application') # Opens application   
    # activeWorkbook = excel.ActiveWorkbook
    # full_path=os.path.join(activeWorkbook.Path, activeWorkbook.Name)
    # df = pd.read_excel(full_path, sheet_name=activeSheet.Name, dtype=str) 



'''

if __name__=='__main__':
    main_menu()
