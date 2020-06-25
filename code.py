from openpyxl import workbook, load_workbook
from lists import database_list_of_lists, python_list, matched_strings, coordinates_list, just_testing, add_to_these_coordinates
from functions import checking_coordinates, get_cordinates, append_list, databse_loop, display_data, adding_letters, check_if_empty, match_coordinate_with_input

excel_file_1 = 'testme2.xlsx'
first_sheet = 'sheet1'

#start of user interface
print("Hello, welcome to the data transfer automation program. To proceed, enter YES")
first_choice = input("")

if (first_choice == 'YES' or first_choice == 'yes'):
    print("Please enter the name of the xlsx file you wish to automate.")
    name_of_sheet = input("")

    #authorization and opening of excel worksheets
    workbook = load_workbook(filename=name_of_sheet)
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
            print(display_data())

        elif load_or_automate == '2':
            print("What is the day you wish to automate?")
            date_of_automation = input("")

            if date_of_automation >= '11':

                #loop through excel and database list to find similarities and append to matched_string list
                append_list(python_list)

                #loop through databse list and python list,
                #find similarities and put them in a seperate list (matched_strings)
                databse_loop(database_list_of_lists, python_list, matched_strings)

                #loop through excel spreadsheet and get coordinates of cell values
                #done in order to match cell value with cordinate
                #find cordinates, remove unwanted characters from returned string
                get_cordinates(matched_strings, coordinates_list)

                #error catching, check if coordinates actually exist and match the strings
                checking_coordinates(coordinates_list)

                adding_letters(coordinates_list, just_testing)

                check_if_empty(just_testing, add_to_these_coordinates)

                match_coordinate_with_input(add_to_these_coordinates, database_list_of_lists)

    elif what_sheet != '1':
        print("functionailty for other sheets has not been implemeneted yet. Thank you.")
        exit(0)

else:
    print("Thank you for using the data transfer automation program.")
    exit(0)
