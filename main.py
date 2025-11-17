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

    # This works for some reason.
    # chdir "%ProgramFiles%\\LibreOffice\\program\\"
    # start soffice "--accept=socket,host=localhost,port=2002;urp;"
    # But not
    # start soffice "--accept=socket,host=localhost,port=2002;urp;" --writer --norestore

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
            lcODT = lcStem + lcOpenDocExt
            lcWord = lcStem + lcWordExt

            print('Master: ' + lcFolder + lcMaster)
            print('Open Document: ' + lcFolder + lcODT)
            print('Word: ' + lcFolder + lcWord)

            start_lo.convert_odm_to_odt(lcFolder, lcMaster, lcODT)
            start_lo.convert_odt_to_docx(lcFolder, lcODT, lcWord)

    start_lo.close_lo()

# -------------------------------------------------------------------------------------------------
