import os


# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
class Constants:
    # -------------------------------------------------------------------------------------------------
    def __init__(self):
        self.LIBRE_OFFICE = 'C:\Program Files\LibreOffice\program\soffice.exe'

        self.TARGET_FOLDER = "Tales from My Neighbor's Desk"

        self.MASTER_EXT = '.odm'
        self.OPEN_DOC_EXT = '.odt'
        self.WORD_EXT = '.docx'

        self.ONEDRIVE = os.getenv('ONEDRIVE')

        found_folders = []
        for root, dirs, files in os.walk(self.ONEDRIVE):
            if self.TARGET_FOLDER in dirs:
                found_folders.append(os.path.join(root, self.TARGET_FOLDER))

        if found_folders and len(found_folders) == 1:
            self.MASTER_FOLDER = found_folders[0]

            print(self.ONEDRIVE)
            print(self.MASTER_FOLDER)
        else:
            print("There must be only one Target Folder under OneDrive")
            print(self.ONEDRIVE)
            print("Folder(s) found, if any")
            print(found_folders)
            exit(1)



# -------------------------------------------------------------------------------------------------


# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
