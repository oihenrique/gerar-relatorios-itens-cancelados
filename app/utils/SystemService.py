import os.path


class SystemService:
    @staticmethod
    def remove_store_sheet(directory, remove_list):
        try:
            if os.path.exists(directory):
                archives = os.listdir(directory)

                for archive in archives:
                    archive_path = os.path.join(directory, archive)
                    for file in remove_list:
                        if str(file) in str(archive):
                            if os.path.isfile(archive_path):
                                os.remove(archive_path)

            else:
                print(f"Directory {directory} doesn't exists.")
        except Exception as e:
            print(f"Error while removing: {str(e)}")
