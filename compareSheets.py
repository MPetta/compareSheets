#!/usr/bin/env python
# coding: utf-8

# In[ ]:

import time
import os
import glob
import pandas as pd
import xlsxwriter
import csv

def getFileName(file):
    ''' Simple splitting function '''
    s = file.split('.')
    name = '.'.join(s[:-1])  # get directory name
    return name

def getSheets(inputfile, fileformat):
    '''Split the sheets in the workbook into seperate csv files in folder
        in the directory named identical to the original filename'''
    name = getFileName(inputfile) # get name for folder
    print ("### Processing File:", name, " ###")
    try:
        os.makedirs(name)
    except:
        pass
    # read as df
    df1 = pd.ExcelFile(inputfile)
    # for each sheet create new file
    for x in df1.sheet_names:
        y = x.lower().replace("-", "_").replace(" ","_")
        print(x + '.' + fileformat, 'Done!')
        df2 = pd.read_excel(inputfile, sheet_name=x)
        filename = os.path.join(name, y + '.' + fileformat)
        if fileformat == 'csv':
            df2.to_csv(filename, index=False)
        else:
            df2.to_excel(filename, index=False)
    print("### Done Processing:", name, " ###")
    
def compareSheets(source, target): 
    ''' Main block will getSheets and compare by writing new files to the directory that show 
        differences by line. The files will then be written as sheets an excel workbook as a report.
        All the files will then be removed from the directory'''
    # start time
    tic = time.perf_counter()
    # split workbooks into csv files
    getSheets(source, "csv")
    getSheets(target, "csv")
    # get wd
    cd = os.getcwd()
    # create lists for names of csv in source and target dir
    sourceList = []
    targetList = []
    # append lists with filenames for source
    for file in os.listdir("Source"):
        if file.endswith(".csv"):
            sourceList.append(file)
    # append lists for filenames for target        
    for file in os.listdir("Target"):
        if file.endswith(".csv"):
            targetList.append(file)
    # sort lists 
    sourceList.sort() 
    targetList.sort() 

    # check if lists are equal 
    if sourceList == targetList: 
        print ("### The Workbook Sheets Have Identical Names ###\n") 
    else : 
        print ("### The Workbook Sheets DO NOT Have Identical Names ###\n") 

    finalList = []
    if sourceList == targetList:
        # iterate over files in list and compare each line
        for i in sourceList:
            with open('Source/' + i , 'r') as t1, open('Target/' + i, 'r') as t2:
                fileone = t1.readlines()
                filetwo = t2.readlines()
            with open(i+'_DIFF.csv', 'w') as outFile:
                for line in filetwo:
                    if line not in fileone:
                        outFile.write(line)
            file = i+'_DIFF.csv'
            finalList.append(file)

    # combine all csv's as sheets in a new workbook
    print("### Writing REPORT.xlsx ###\n")

    workbook = xlsxwriter.Workbook(cd+'/REPORT.xlsx')
    #counter = 0
    for csv_file in glob.glob(os.path.join(cd, '*DIFF.csv')):
        sheet_name1 = str(csv_file.rsplit('/', 1)[1])
        sheet_name = str(sheet_name1.rsplit('.', 2)[0])
        #counter += 1
        worksheet = workbook.add_worksheet(sheet_name)
        with open(csv_file, 'rt', encoding='utf8') as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    worksheet.write(r, c, col)
    workbook.close()

    # Get a list of all the file paths that ends with DIFF.csv from directory
    fileList = glob.glob('*DIFF.csv')
    # Iterate over the list of filepaths & remove each file.
    for filePath in fileList:
        try:
            os.remove(filePath)
        except:
            print("Error while deleting file : ", filePath)

    # report execution  
    print("### Job Complete ###")
    toc = time.perf_counter()
    print(f"Total Execution Time: {toc - tic:0.4f} seconds")