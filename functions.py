from openpyxl import workbook, load_workbook

excel_file_1 = 'testme2.xlsx'
workbook = load_workbook(filename='testme2.xlsx')
sheet = workbook.active

"""def checking_coordinates(matched_strings, coordinates_list):
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
                        returnMe = print("hi")
                        return returnMe
                    else:
                        break
                        return False
        else:
            break"""

def checking_coordinates(matched_strings, coordinates_list):
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
                    cut_cell = cell
                    cut_cell = str(cut_cell).replace('<Cell \'sheet1\'.', '')
                    cut_cell = str(cut_cell).replace('>', '')
                    if cut_cell == c:
                        printme = print("coordinates match, to continue automating press enter.")
                        return printme
                    else:
                        printme = print("coordinates do not match")
                        return printme
        else:
            break
