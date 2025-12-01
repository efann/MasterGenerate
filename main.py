import os
import sys

import pyperclip

from RunLO import RunLO

# Establish connection to LibreOffice

# -------------------------------------------------------------------------------------------------
# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    run_lo = RunLO()

    interpreter_path = sys.executable
    print("Python Interpreter: " + interpreter_path)
    if run_lo.constants.LIBRE_OFFICE_FOLDER_BASE.casefold() in interpreter_path.casefold():
        print('Correct Python Interpreter is being used.')
    else:
        print('LibreOffice Python Interpreter must be used.')
        exit()

    lcFolder = run_lo.constants.MASTER_FOLDER

    print(f"Folder: {lcFolder}")

    print(run_lo.constants.LINE_MARKER)

    run_lo.load_lo()
    run_lo.is_libreoffice_listener_running()
    run_lo.constants.print_line_marker()

    lcMasterExt = run_lo.constants.MASTER_EXT
    lcOpenDocExt = run_lo.constants.OPEN_DOC_EXT
    lcWordExt = run_lo.constants.WORD_EXT

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

            if run_lo.convert_odm_to_odt(lcMasterFile, lcODTFile):
                if run_lo.convert_odt_to_docx(lcODTFile, lcWordFile):
                        try:
                            print(f"Opening the document: {lcWordFile}")
                            os.startfile(lcWordFile)

                            pyperclip.copy(run_lo.constants.TEMPLATE_FILE)
                            print(
                                f"Template file found here:\n\n{run_lo.constants.TEMPLATE_FILE}\n\nCopied to the clipboard, by the way, for use in Word | Developer | Document Template\n\n")
                        except Exception as e:
                            print(f"An error occurred opening Word: {e}")

    run_lo.terminate_libreoffice()
    sys.exit(0)

# -------------------------------------------------------------------------------------------------
