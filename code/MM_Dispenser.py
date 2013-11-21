import time, datetime
import xlrd

## Define parameters
Loop_time = 5  # The time in seconds to wait before scanning for completed tasks
#prev_dict = {}
#curr_dict = {}

### Open the pre-defined Microsoft Excel spreadsheet 

todo_spreadsheet = "MM_Dispenser_Test.xls"
book = xlrd.open_workbook(todo_spreadsheet, formatting_info=1) 
sheets = book.sheet_names()
print "The sheet names are:", sheets 
#book.release_resources()

def get_current_state():
    """
    read the excel spreadsheet into a native python format
    """
    #seems like a list of lists (or a list of lists of lists) (sheets, rows, columns)
    #should be sufficient
    #this also gives order data, in case that is ever needed
    current_state = []

    #need to re-open every time to make sure we're getting the latest version
    book = xlrd.open_workbook(todo_spreadsheet, formatting_info=1) 
    
    sheets = book.sheet_names()
    for index, sh in enumerate(sheets): 
        sheet = book.sheet_by_index(index) 
        #print "\nSheet name:", sheet.name 
        rows, cols = sheet.nrows, sheet.ncols 
        #print "Number of rows: %s   Number of cols: %s" % (rows, cols)
        current_rows = []
        for row in range(rows):
            thecell = sheet.cell(row, 1)
            xfx = sheet.cell_xf_index(row, 2)
            xf = book.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
            current_cols = [ thecell.value, bgx ]
            current_rows.append(current_cols)
        current_state.append(current_rows)

    #book.release_resources()
    return current_state

def compare_states(prev_state, new_state):
    for sheet_index in range(len(prev_state)):
        for row_index in range(len(prev_state[sheet_index])):
            #check if the row is still in the new_state:
            prev_row = prev_state[sheet_index][row_index]
            if prev_row in new_state[sheet_index]:
                new_index = new_state[sheet_index].index( prev_row )
                if new_index == row_index:
                    #position and row is the same, nothing to see here
                    pass
                else:
                    print "Noticed that: %s moved from position: %s to position: %s" % (prev_row, row_index, new_index)
            else:
                #can't find the exact same task / color combination in the current sheet
                #check if the color has changed
                matched = False
                for new_row in new_state[sheet_index]:
                    if new_row[0] == prev_row[0]:
                        print "Found color change for: %s to: %s" % (prev_row, new_row)
                        matched = True

                if not matched:
                    print "Couldn't find: %s in current spread sheet (new_state). deleted?" % prev_row
                    
                # this only checks for items missing from new_state
                # does not signal for new rows added to new_state
    
prev_state = get_current_state()
print prev_state

print '\nCompleted reading contents for current Task List\n'

### Continually loop every 'Loop_time' seconds, and scan for compeleted tasks, unless interrupted
while True:
    try: ## The code below will be executed every N seconds

        new_state = get_current_state()
        #print new_state

        ## Compare the prev_dict with the curr_dict and identify a completed task
        ## Take one curr_dict entry at a time, search for it in the prev_dict, ...
        ##      and see if the pattern_colour_index has changed 
        compare_states(prev_state, new_state)

        #be sure to update... don't want to dispense more than deserved amount! :)
        prev_state = new_state

        ## Pause before repeating search for completed tasks
        print '\nPausing before repeating search for changes (%s) \n' % (datetime.datetime.now())

        time.sleep(Loop_time)

    ## When the Ctrl+C keystroke is made the program terminates 
    except (KeyboardInterrupt, SystemExit):
        print '\nProgram terminated! No traceback to see here. Get back to work!\n'
        break

    ## Print statements and complete any final housekeeping chores before halting program execution 
    finally:
        pass


