import sys
import os.path
from app.excel.SpreadsheetGenerator import SpreadsheetGenerator
from app.email.EmailApp import send_email
from app.utils.SystemService import SystemService

stores = [2, 6, 8, 10, 11, 12, 15, 16, 20, 21, 22,
          23, 24, 25, 29, 30, 35, 37, 47, 48, 49,
          50, 51, 55, 57, 61, 63, 64, 65, 66, 67, 68, 69,
          70, 80, 81, 83]

generator = SpreadsheetGenerator(stores)
SystemService.create_main_data_folder()


def create_spreadsheet_from_clipboard(file_name):
    absolute_path = os.path.abspath(f'../data/{file_name}_cancelados')

    generator.create_spreadsheet_from_clipboard(absolute_path)


def generate_spreadsheets(file_name):
    if __name__ == "__main__":
        absolute_path = os.path.abspath(f'../data/{file_name}_cancelados.xlsx')
    else:
        absolute_path = os.path.abspath(f'data/{file_name}_cancelados')

    return generator.generate_spreadsheets(absolute_path, file_name)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        function_name = sys.argv[1]
        if function_name == 'create_spreadsheet_from_clipboard':
            date = sys.argv[2] if len(sys.argv) > 2 else None
            if date is None:
                print('date argument required')
            else:
                create_spreadsheet_from_clipboard(date)
        elif function_name == 'generate_spreadsheets':
            date = sys.argv[2] if len(sys.argv) > 2 else None
            if date is None:
                print('date argument required')
            else:
                generate_spreadsheets(date)
        elif function_name == 'send_email':
            date = sys.argv[2] if len(sys.argv) > 2 else None
            if date is None:
                print('date argument required')
            else:
                send_email(date)
        else:
            print(f'Unknown function: {function_name}')
    else:
        print("No function specified.")
