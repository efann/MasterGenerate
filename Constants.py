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

        self.ONEDRIVE = None
        self.MASTER_FOLDER = None

        self.initialize_folders()

    # -------------------------------------------------------------------------------------------------

    def initialize_folders(self):
        self.ONEDRIVE = self.get_onedrive()
        if not self.ONEDRIVE:
            print("OneDrive path could not be determined.")
            exit(1)

        found_folders = []
        for root, dirs, files in os.walk(self.ONEDRIVE):
            if self.TARGET_FOLDER in dirs:
                found_folders.append(os.path.join(os.path.join(root, self.TARGET_FOLDER), ''))

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

    def get_onedrive(self):

        onedrive_path = os.environ.get('OneDrive')
        if onedrive_path:
            return onedrive_path
        else:
            # Fallback for systems where 'OneDrive' environment variable might not be set
            # This attempts to find a common OneDrive path based on USERPROFILE
            user_profile = os.environ.get('USERPROFILE')
            if user_profile:
                potential_path = os.path.join(user_profile, 'OneDrive')
                if os.path.isdir(potential_path):
                    return potential_path
            return None

    # -------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
