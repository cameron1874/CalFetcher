import shutil
import numpy
import datetime
import os
from tkinter import Tk
from tkinter.filedialog import askdirectory
import pandas as pd
import openpyxl


def remove_values_from_list(the_list, val):
    return [value for value in the_list if value != val]


def removecomments(calibrations, numLines):
    for index in range(0, numLines):
        tmpStr = calibrations[index]
        strIdx = tmpStr.find('%')
        if strIdx != -1:
            tmpStr = tmpStr[0:strIdx - 1] + '\n'
            calibrations[index] = tmpStr
    return calibrations


def setup_files():
    with open(tempfile, 'w+') as file:
        pass

    with open(calibrationTxtFile, 'w+') as file:
        pass

    with open(outFile, 'w+') as f4:
        f4.truncate(0)
        datestamp = datetime.datetime.now()
        date_time = datestamp.strftime("%m/%d/%Y, %H:%M:%S\n")
        f4.write(date_time)

def multi_line_cal_to_single_line(calibrations, numLines):
    for index in range(numLines, 0, -1):
        curStr = calibrations[index]
        closedBracket = curStr.find(']')
        openBracket = curStr.find('[')
        count = 0
        # if closedBracket != -1 and openBracket == -1: # Detecting multiple lines (note reading from the bottom up)
        if closedBracket != -1 and openBracket == -1:  # Detecting multiple lines (note reading from the bottom up)
            while openBracket == -1:
                # move current index-count to index-count-1
                calibrations[index - count - 1] = calibrations[index - count - 1].strip() + ' ; ' + calibrations[
                    index - count]
                # remove data from index-count
                calibrations[index - count] = 'DELETE'

                # conditions for the next check
                count = count + 1
                curStr2 = calibrations[index - count];
                openBracket = curStr2.find('[')
            index = index - count
    calibrations = remove_values_from_list(calibrations, 'DELETE')


def write_calibrations_file(calibrations, numLines):
    numLines = sum(1 for line in calibrations) - 1
    with open(calibrationTxtFile, 'w') as f2:
        for index in range(0, numLines):
            f2.write(calibrations[index])


def search_cal_file_for_referenced_cal_names():
    with open(referenceFile, 'r') as f2:
        f0.write('Data based on input file: ' + inputFile + '\n\n')
        for index in f2:
            cal_names = index.splitlines()
            for index2 in range(0, sum(1 for line in cals_text) - 1):
                str2 = cals_text[index2]
                cal_text_indexes = str2.find(cal_names[0].strip())

                # split header from string, header is separated by '.'
                if '.' not in str2:
                    continue
                else:
                    dot_index = str2.index(".")
                    header, str2 = str2.split('.', 1)

                if cal_text_indexes != -1:
                    f0.write(str2)
                    print(str2)
                    flag = 1
                    break
                else:
                    flag = 0
            if flag < 1:  # not found
                f0.write(cal_names[0] + '\tNOT FOUND\n')
                array.append(cal_names[0])
            else:
                flag = 0  # found string in nested loop


def write_to_excel():
    pass


# Author: Cody Palmer
# Last modified by: Cameron Floyd
# Last modified Date: 3-14-2024

# There is a README.txt for instructions

rootdir = askdirectory(title='Select Folder that Contains your .m Files')

wrong_file_counter = 0

# loop to iterate through files in selected folder
for subdir, dirs, files in os.walk(rootdir):
    for inputFile in files:
        if str(inputFile).endswith('.py'):
            continue
        elif str(inputFile).endswith('.txt'):
            continue
        elif str(inputFile).endswith('.m') != True:
            wrong_file_counter += 1
            continue

        ### File Preparation ###

        ### Initialize Names of Files Needed ###
        tempfile = 'temp.txt'
        outFile = 'out.txt'
        referenceFile = 'references.txt'
        calibrationTxtFile = 'calibrations.txt'
        #Create/overwrite temp file, calibrationsfile, outfile, and add timedate to outfile



        setup_files()

        # copy matlab file to temp file
        shutil.copyfile(rootdir + '/' + inputFile, tempfile)

        # create/overwrite calibrations file
        with open(calibrationTxtFile, 'w+') as f4:
            f4.truncate(0)

        ### Data Manipulation ###

        # assign values to calibrations variable
        with open(tempfile, 'r') as f1:
            calibrations = f1.readlines()

        numLines = sum(1 for line in calibrations) - 1

        # Strip Comments that came from Matlab File
        # Take any Cals that span more than 1 line and put onto 1 line
        # Write to CalibrationsTxtFile
        removecomments(calibrations, numLines)
        multi_line_cal_to_single_line(calibrations, numLines)
        write_calibrations_file(calibrations, numLines)

        # read in the file you just created
        with open(calibrationTxtFile, 'r+') as f3:
            cals_text = f3.readlines()

        # search for references
        array = []
        with open(outFile, 'a') as f0:
            search_cal_file_for_referenced_cal_names()
            f0.write('\n\n\n')

        write_to_excel()

print(
    f"\nAll Finished. Your calibrations are found in out.txt in the same folder as this .py file.\n{wrong_file_counter} files in that folder were not .m files and were not read.")




