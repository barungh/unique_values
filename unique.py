import xlwings as xw
from datetime import datetime
import os

current_directory = os.getcwd()

def select_file(directory):
    files = os.listdir(directory)
    print("Files in Current Directory:")
    for i, file in enumerate(files):
        print(f"{i+1}. {file}")
    while True:
        try:
            choice = int(input("Select appropriate Excel file : "))
            if choice < 1 or choice > len(files):
                raise ValueError("Invalid selection")
            break
        except ValueError:
            print("Invalid input. Please enter a valid number.")
    selected_file = files[choice - 1]
    file_extension = os.path.splitext(selected_file)[1]
    
    if file_extension.lower() in ['.xls', '.xlsx']:
        return selected_file
    else:
        raise ValueError("Selected file is not Excel file")

try:
    file = select_file(current_directory)
    book = xw.Book(file)
    sheet1 = book.sheets[0]
    sheet1.range("G1").value = None
    sheet1.range("G2").value = None
    values = sheet1.range("A:A").value
    values = values[1:]
    valid_values = [value for value in values if value is not None]
    
    try:
        times = [datetime.strptime(value, '%H:%M:%S').time() for value in valid_values]
    except ValueError as e:
        raise ValueError("Data in column A:A is in invalid format")

    unique_times = list(set(times))
    sheet1.range("G1").value = "Punch Count"
    sheet1.range("G2").value = len(unique_times)
except ValueError as err:
    print(err)
