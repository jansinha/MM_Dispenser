#!usr/bin/env
"""
M&M Dispenser Host Computer Program
This program interfaces with a Microsoft Excel spreadsheet & communicates signals to an Arduino via serial interface.

MS Excel communication is made possible using the XLRD library
Serial commuincation is possible using the PySerial library

Current version uses dict to represent represent MS Excel spreadsheet contents
See DictDiffer library
"""

import time
import xlrd
import serial
import copy
from DictDiffer import *

## Define parameters
Loop_time = 3  # The time in seconds to wait before scanning for completed tasks
prev_dict = {}
curr_dict = {}

## Configure serial (USB) line
ser = serial.Serial('/dev/tty.usbserial-A700eY2N',9600,timeout=1)        
print ser.portstr

### Open the pre-defined Microsoft Excel spreadsheet 
book = xlrd.open_workbook("/Users/Shiva/Documents/Life/Hobbies/Arduino/Candy_Dispenser/MM_Dispenser_Test_Short.xls", formatting_info=1) 
sheets = book.sheet_names()
print "The sheet names are:", sheets 


### First time the program runs, collect the current state of the MS Excel ssheet
for index, sh in enumerate(sheets): 
    sheet = book.sheet_by_index(index) 
    print "\nSheet name:", sheet.name 
    rows, cols = sheet.nrows, sheet.ncols 
    print "Number of rows: %s   Number of cols: %s" % (rows, cols) 
    for row in range(rows): 
        for col in range(cols):
            ## print "row, col is:", row+1, col+1,
            if (col+1) == 2:
                thecell = sheet.cell(row, col)  # could get 'dump', 'value', 'xf_index'
                print thecell.value,
                prev_dict[thecell.value]=0

            elif (col+1) == 3:
                xfx = sheet.cell_xf_index(row, col) 
                xf = book.xf_list[xfx] 
                bgx = xf.background.pattern_colour_index
                print bgx
                prev_dict[thecell.value]=bgx

print '\nCompleted reading contents for current Task List\n'

    
### Continually loop every 'Loop_time' seconds, and scan for compeleted tasks, unless interrupted
while True:
    try: ## The code below will be executed every N seconds

        ## Pause before repeating search for completed tasks
        print '\nPausing before repeating search for changes\n'
        time.sleep(Loop_time)

        book = xlrd.open_workbook("/Users/Shiva/Documents/Life/Hobbies/Arduino/Candy_Dispenser/MM_Dispenser_Test_Short.xls", formatting_info=1) 
        sheets = book.sheet_names()
        sheet = book.sheet_by_index(0) 
        rows, cols = sheet.nrows, sheet.ncols 

        ## Read in the current entries from the MS Excel spreadsheet
        for index, sh in enumerate(sheets): 
            sheet = book.sheet_by_index(index) 
            print "Sheet name:", sheet.name 
            rows, cols = sheet.nrows, sheet.ncols 
            print "Number of rows: %s   Number of cols: %s" % (rows, cols) 
            for row in range(rows): 
                for col in range(cols):
                    if (col+1) == 2:
                        thecell = sheet.cell(row, col)  # could get 'dump', 'value', 'xf_index'
                        if thecell.value:
                            print thecell.value,
                        elif empty_cell(thecell.value):
                            print '-',
                        curr_dict[thecell.value]=0

                    elif (col+1) == 3:
                        xfx = sheet.cell_xf_index(row, col) 
                        xf = book.xf_list[xfx] 
                        bgx = xf.background.pattern_colour_index
                        print bgx
                        curr_dict[thecell.value]=bgx

        ## Compare the prev_dict with the curr_dict and identify a completed task
        ## Take one curr_dict entry at a time, search for it in the prev_dict, ...
        ##      and see if the pattern_colour_index has changed 
        myClass = DictDiffer(curr_dict,prev_dict)
        changedSet = myClass.changed()
        if (len(changedSet) > 0):
            print changedSet
            for i in changedSet:
                print 'current_dict value: ', curr_dict[i]
                print 'previous_dict value: ', prev_dict[i]
                ## Output the task value to the serial line
                x = ser.write(str(curr_dict[i]))

            prev_dict.clear()
            prev_dict = copy.deepcopy(curr_dict)
            curr_dict.clear()
            print "A new task has been completed."

        else:
            print "No new task completed."


    ## When the Ctrl+C keystroke is made the program terminates 
    except (KeyboardInterrupt, SystemExit):
        print '\nProgram terminated! Need to figure out how to terminate without Traceback!!\n'
        break

    ## Print statements and complete any final housekeeping chores before halting program execution 
    finally:
        ser.close
        pass
