import locale
import os.path
from datetime import datetime


class SystemService:
    """
    A class containing static methods related to system services.
    """

    @staticmethod
    def remove_store_sheets(directory, remove_list):
        """
        Remove specified files from the given directory.

        Parameters:
            directory (str): The directory from which files should be removed.
            remove_list (list): List of file names to be removed.

        Returns:
            None
        """
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
    def verify_if_folder(target_folder, folder_name):
        """
        Check if a folder exists inside the target directory.

        Parameters:
            target_folder (str): The target directory to check.
            folder_name (str): The name of the folder to check for.

        Returns:
            bool: True if the folder exists, False otherwise.
        """
        absolute_path = os.path.abspath(target_folder)
        try:
            folders = os.listdir(absolute_path)
            return folder_name in folders
        except FileNotFoundError:
            return False

    @staticmethod
    def get_attach_folder(date):
        """
        Get the attachment folder path based on the provided date.

        Parameters:
            date (str): The date in the format 'dd-mm-yyyy'.

        Returns:
            str: The absolute path to the attachment folder.
        """
        current_day, current_month, current_year = map(int, date.split('-'))

        locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
        month_name = datetime.now().strftime("%B") if current_month == datetime.now().month else \
            datetime.strptime(str(current_month), "%m").strftime("%B")

        final_folder = f'../data/store_sheets/{current_year}/{month_name}/' \
                       f'{current_day}-{current_month}-{current_year}'

        return os.path.abspath(final_folder)

    @staticmethod
    def create_folder(folder_path):
        """
        Create a folder at the specified path if it doesn't exist.

        Parameters:
            folder_path (str): The path of the folder to create.

        Returns:
            None
        """
        try:
            os.makedirs(folder_path)
            print(f"Folder created successfully: {folder_path}")
        except FileExistsError:
            print(f"Folder {folder_path} already exists.")
        except Exception as e:
            print(f"Error while creating folder {folder_path}: {str(e)}")

    @staticmethod
    def create_date_folders(date):
        """
        Create date-specific folders for storing data.

        Parameters:
            date (str): The date in the format 'dd-mm-yyyy'.

        Returns:
            None
        """
        current_day, current_month, current_year = map(int, date.split('-'))

        locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
        month_name = datetime.now().strftime("%B") if current_month == datetime.now().month else \
            datetime.strptime(str(current_month), "%m").strftime("%B")

        for i in range(3):
            year_folder = f'../data/store_sheets/{current_year}'
            month_folder = f'{year_folder}/{month_name}'
            day_folder = f'{month_folder}/{current_day}-{current_month}-{current_year}'

            if not SystemService.verify_if_folder('../data/store_sheets/', current_year):
                SystemService.create_folder(year_folder)
            elif not SystemService.verify_if_folder(month_folder):
                SystemService.create_folder(month_folder)
            elif not SystemService.verify_if_folder(day_folder):
                SystemService.create_folder(day_folder)

    @staticmethod
    def create_main_data_folder():
        """
        Create the main data folder structure.

        Returns:
            None
        """
        for i in range(3):
            if not SystemService.verify_if_folder('../', 'data'):
                SystemService.create_folder('../data')
            elif not SystemService.verify_if_folder('../data/', 'store_sheets'):
                SystemService.create_folder('../data/store_sheets')
