import locale
import os.path
import shutil

from datetime import datetime


class SystemService:
    @staticmethod
    def remove_store_sheets(directory, remove_list):
        try:
            if os.path.exists(directory):
                for file_to_remove in remove_list:
                    file_path = os.path.join(directory, file_to_remove)
                    if os.path.exists(file_path) and os.path.isfile(file_path):
                        os.remove(file_path)
                        print(f"Removed: {file_path}")
                    else:
                        print(f"File not found: {file_path}")
            else:
                print(f"Directory {directory} doesn't exist.")
        except Exception as e:
            print(f"Error while removing: {str(e)}")

    @staticmethod
    def verify_if_folder(target_folder, year):
        absolute_path = os.path.abspath(target_folder)

        try:
            folders = os.listdir(absolute_path)
            return year in folders
        except FileNotFoundError:
            return False

    @staticmethod
    def create_folder(folder_path):
        try:
            os.makedirs(folder_path)
            print(f"Folder created successfully: {folder_path}")
        except FileExistsError:
            print(f"Folder {folder_path} already exists.")
        except Exception as e:
            print(f"Error while creating folder {folder_path}: {str(e)}")

        return folder_path

    @staticmethod
    def create_date_folders(date):
        current_day = date.split('-')[0]
        current_month = date.split('-')[1]
        current_year = date.split('-')[2]
        final_folder = ''

        locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
        if current_month == str(datetime.now().month):
            month_name = str(datetime.now().strftime("%B"))
        else:
            month_name = datetime.strptime(current_month, "%m").strftime("%B")

        for i in range(0, 3):
            if not SystemService.verify_if_folder('../data/store_sheets/', current_year):
                SystemService.create_folder(f'../data/store_sheets/{current_year}')
            elif not SystemService.verify_if_folder(f'../data/store_sheets/{current_year}', month_name):
                SystemService.create_folder(f'../data/store_sheets/{current_year}/{month_name}')
            elif not SystemService.verify_if_folder(f'../data/store_sheets/{current_year}/{month_name}/',
                                                    f'{current_day}-{current_month}-{current_year}'):
                final_folder = SystemService.create_folder(f'../data/store_sheets/{current_year}/{month_name}/'
                                                           + f'{current_day}-{current_month}-{current_year}')

            return final_folder
