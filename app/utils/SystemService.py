import locale
import os.path
from datetime import datetime


class SystemService:
    """
    A class containing static methods related to system services.
    """

    @staticmethod
    def absolute_file_paths(directory, filename):
        """
        Retrieve the absolute paths of files within a specified directory, searching for a specific filename.

        Parameters:
        - directory (str): The base directory to start the search.
        - filename (str): The name of the file to search for.

        Returns: - str: The absolute path of the file if found, or 'file not found' if the file is not present in
        the directory.

        Example:
        ```python
        result = YourClassName.absolute_file_paths('data', '27-12-2023.xlsx')
        print(result)
        ```
        """
        path_list = []
        for root, dirs, files in os.walk(os.path.abspath(directory)):
            for file in files:
                path_list.append(os.path.join(root, file))
                if filename == file:
                    return os.path.join(root, file)

        return 'file not found'

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
                    file_path = SystemService.absolute_file_paths(directory, file_to_remove)
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
                       f'{str(current_day).zfill(2)}-{str(current_month).zfill(2)}-{current_year}'

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
        current_day = int(date.split('-')[0])
        current_month = int(date.split('-')[1])
        current_year = int(date.split('-')[2])

        locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
        month_name = datetime.now().strftime("%B") if current_month == datetime.now().month else \
            datetime.strptime(str(current_month), "%m").strftime("%B")

        base_folder = '../data/store_sheets'
        year_folder = os.path.join(base_folder, str(current_year))
        month_folder = os.path.join(year_folder, month_name)
        day_folder = os.path.join(month_folder, f'{str(current_day).zfill(2)}-{str(current_month).zfill(2)}-{current_year}')

        if not SystemService.verify_if_folder(base_folder, str(current_year)):
            SystemService.create_folder(year_folder)
        if not SystemService.verify_if_folder(month_folder, month_name):
            SystemService.create_folder(month_folder)
        if not SystemService.verify_if_folder(day_folder,
                                              f'{str(current_day).zfill(2)}-{str(current_month).zfill(2)}-{current_year}'):
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
