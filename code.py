from openpyxl import workbook, load_workbook
import pandas as pd
import numpy as np
import xlrd as xl


#lists for data storage and automation processes
excel_file_1 = 'testme2.xlsx'
first_sheet = 'sheet1'
matched_strings = []
python_list = []
#!IMPORTANT! copy paste data from phpMyAdmin into this list to run program
database_list_of_lists = [('1', 33789190, 'Al Jasrah', '25', '1'),
                            ('2', 55860636, 'Al Jasrah', '50', '25.2841, 51.441'),
                            ('3', 55150250, 'Al Jasrah', '50', '25.2841, 51.441'),
                            ('4', 66570312, 'Al Jasrah', '24', '25.3318,51.5255'),
                            ('5', 55439821, 'Al Jasrah', '24', 'blah'),
                          ('5', 55512402, 'Al Jasrah', '24', 'blah')]
coordinates_list = []
just_testing = []

#start of user interface
print("Hello, welcome to the data transfer automation program. To proceed, enter YES")
first_choice = input("")

if (first_choice == 'YES' or first_choice == 'yes'):
    print("Please enter the name of the xlsx file you wish to automate.")
    name_of_sheet = input("")

    #authorization and opening of excel worksheets
    workbook = load_workbook(filename=name_of_sheet)
    sheet = workbook.active
    print("The following sheets are avaliable within the xlsx file: ")
    print(workbook.sheetnames)
    print("Which sheet would you like to access?")
    what_sheet = input("")

    #Asks what sheet you would like to access
    if (what_sheet == '1' or what_sheet == 'sheet1'):
        print("accessing batching_sheet....")

        print("To load data, enter 1. To skip that step and begin automating data, press 2.")
        load_or_automate = input("")

        #allows the user to simply see the data within the sheet without automating anything,
        #basically just adding functionality to the program
        if load_or_automate == '1':
            #variables defined at the start of script
            print("Data within the sheet: ")
            for value in sheet.iter_rows(min_row=1,
                                         max_row=47,
                                         min_col=7,
                                         max_col=8,
                                         values_only=True):
                print(value)

        elif load_or_automate == '2':
            print("What is the day you wish to automate?")
            date_of_automation = input("")

            if date_of_automation >= '11':

                #loop through excel and database list to find similarities and append to matched_string list
                for value in sheet.iter_rows(min_row=25,
                                            max_row=46,
                                            min_col=6,
                                            max_col=7,
                                            values_only=True):
                        python_list.append(value)

                #loop through databse list and python list,
                #find similarities and put them in a seperate list (matched_strings)
                for list in database_list_of_lists:
                    for phone_number in list:
                        for second_list in python_list:
                            for number in second_list:
                                if number == phone_number:
                                    matched_strings.append(phone_number)

                #loop through excel spreadsheet and get coordinates of cell values
                #done in order to match cell value with cordinate
                #find cordinates, remove unwanted characters from returned string
                for row in sheet.iter_rows(min_row=25,
                                           max_row=46,
                                           min_col=6,
                                           max_col=7):
                    for cell in row:
                        for number in matched_strings:
                            if cell.value == number:
                                tryMe = cell
                                newTryMe = str(tryMe).replace('<Cell \'sheet1\'.', '')
                                append_this = str(newTryMe).replace('>', '')
                                coordinates_list.append(append_this)


                #error catching, check if coordinates actually exist and match the strings

                def checking_coordinates(matched_strings, coordinates_list):
                    i = 0
                    while i < len(coordinates_list):
                        i += 1
                        if i < len(coordinates_list):
                            c = sheet[coordinates_list[i]]
                            for phone_number in matched_strings:
                                j = 0
                                while j < len(matched_strings):
                                    j += 1
                                    if j < len(matched_strings):


                        else:
                            break




    elif what_sheet != '1':
        print("functionailty for other sheets has not been implemeneted yet. Thank you.")
        exit(0)

else:
    print("Thank you for using the data transfer automation program.")
    exit(0)
