import os
import subprocess
from asyncio import start_server

from Constants import Constants
from UseLO import StartLO

import sys
import uno

# Establish connection to LibreOffice

constants = Constants()


# -------------------------------------------------------------------------------------------------

def open_libreoffice(port):
    # UNO component context for initializing the Python runtime
    localContext = uno.getComponentContext()

    # Create an instance of a service implementation
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext)

    context = resolver.resolve(
        "uno:socket,host=localhost,"
        f"port={port};urp;StarOffice.ComponentContext")

    desktop1 = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context)
    return desktop1

    # Get the component context
    localContext = uno.getComponentContext()

    # Create a resolver to connect to the running LibreOffice instance
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext
    )

    # Connect to the LibreOffice instance
    args = "socket,host=localhost,port=2021"
    ctx = resolver.resolve(f"uno:{args};"
                           "urp;"
                           "StarOffice.ComponentContext")

    loServiceManager = ctx.ServiceManager
    loDesktop = loServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    return loDesktop


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

    start_lo = StartLO()
    start_lo.start_libreoffice_headless(constants.LIBRE_OFFICE_PORT)

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

            # convert_odm_to_odt(lcFolder + lcMaster)
            # convert_odt_to_docx(lcFolder + lcODT, lcFolder)

            desktop = open_libreoffice(constants.LIBRE_OFFICE_PORT)

# -------------------------------------------------------------------------------------------------
