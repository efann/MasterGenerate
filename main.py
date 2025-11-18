import os
import sys

from Constants import Constants
from RunLO import StartLO

# Establish connection to LibreOffice

constants = Constants()

# -------------------------------------------------------------------------------------------------
# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    interpreter_path = sys.executable
    print("Python Interpreter: " + interpreter_path)
    if constants.LIBRE_OFFICE_FOLDER_BASE.casefold() in interpreter_path.casefold():
        print('Correct Python Interpreter is being used.')
    else:
        print('LibreOffice Python Interpreter must be used.')
        exit()

    lcFolder = constants.MASTER_FOLDER

    print(f"Folder: {lcFolder}")

    print(constants.LINE_MARKER)

    start_lo = StartLO()
    start_lo.load_lo()

    lcMasterExt = constants.MASTER_EXT
    lcOpenDocExt = constants.OPEN_DOC_EXT
    lcWordExt = constants.WORD_EXT

    for lcFilename in os.listdir(lcFolder):
        if lcFilename.endswith(lcMasterExt) and os.path.isfile(os.path.join(lcFolder, lcFilename)):
            lcMaster = lcFilename

            lcStem = os.path.splitext(lcFilename)[0]

            lcMasterFile = lcFolder + lcMaster
            lcODTFile = lcFolder + lcStem + lcOpenDocExt
            lcWordFile = lcFolder + lcStem + lcWordExt

            print(f'Master Document: {lcMasterFile}')
            print(f'Writer Document: {lcODTFile}')
            print(f'Word Document: {lcWordFile}')

            start_lo.convert_odm_to_odt(lcMasterFile, lcODTFile)
            start_lo.convert_odt_to_docx(lcODTFile, lcWordFile)

            try:
                print(f"Opening the document: {lcWordFile}")
                os.startfile(lcWordFile)
            except Exception as e:
                print(f"An error occurred opening Word: {e}")

    start_lo.close_lo()

# -------------------------------------------------------------------------------------------------
