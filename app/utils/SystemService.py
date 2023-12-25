import os.path

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
