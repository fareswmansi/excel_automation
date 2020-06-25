from openpyxl import workbook, load_workbook

excel_file_1 = 'testme2.xlsx'
workbook = load_workbook(filename='testme2.xlsx')
sheet = workbook.active

#display data to user interface
def display_data():
    for value in sheet.iter_rows(min_row=1,
                                 max_row=47,
                                 min_col=7,
                                 max_col=8,
                                 values_only=True):
        print(value)


#loop through excel and append values to python_list
def append_list(python_list):
    for value in sheet.iter_rows(min_row=25,
                                max_row=46,
                                min_col=6,
                                max_col=7,
                                values_only=True):
        python_list.append(value)

#loop through databse list and python list, append similarities to matched_string list
def databse_loop(database_list_of_lists, python_list, matched_strings):
    for list in database_list_of_lists:
        for phone_number in list:
            for second_list in python_list:
                for number in second_list:
                    if number == phone_number:
                        matched_strings.append(phone_number)

#get coordinates of matched_strings, append them to a seperate list and cut string to only
#have the coordinate rather than the whole returned tring
def get_cordinates(matched_strings, coordinates_list):
    for row in sheet.iter_rows(min_row=25,
                                           max_row=46,
                                           min_col=6,
                                           max_col=7):
                    for cell in row:
                        for number in matched_strings:
                            if cell.value == number:
                                tryMe = cell
                                tryMe = str(tryMe).replace('<Cell \'sheet1\'.', '')
                                append_this = str(tryMe).replace('>', '')
                                coordinates_list.append(append_this)

#check coordinates existance within excel in order to catch errors
def checking_coordinates(coordinates_list):
    i = 0
    while i < len(coordinates_list):
        i += 1
        if i < len(coordinates_list):
            c = sheet[coordinates_list[i]]
            for row in sheet.iter_rows(min_row=25,
                                       max_row= 46,
                                       min_col=6,
                                       max_col=7):
                for cell in row:
                    if c == cell:
                        return True
        else:
            break

#change coordinate column to match input field
def adding_letters(coordinates_list, just_testing):
    i = 0
    firstTry = str(coordinates_list[0]).replace('G', 'M')
    just_testing.insert(0, firstTry)
    while i < len(coordinates_list):
        i += 1
        if i < len(coordinates_list):
            tryThis = str(coordinates_list[i]).replace('G', 'M')
            just_testing.append(tryThis)


#check if input field is empty
def check_if_empty(just_testing):
    i = 0
    while i < len(just_testing):
        i += 1
        if i < len(just_testing):
            c = sheet[just_testing[i]].value
            if c == 'None':
                return just_testing[i]
