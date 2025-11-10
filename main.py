import os
import subprocess
from Constants import Constants
#from com.sun.star.beans import PropertyValue

from sys import path

from StartLO import StartLO

path.append('C:\Program Files\LibreOffice\program')

import sys
import uno

# Establish connection to LibreOffice

constants = Constants()


# -------------------------------------------------------------------------------------------------
def convert_odm_to_odt(odm_filepath):
    """
    Converts a LibreOffice master document (.odm) to an ODT file using the
    LibreOffice command-line interface.

    Args:
        odm_filepath (str): The full path to the master document (.odm).
    """
    input_path = os.path.abspath(odm_filepath)
    output_dir = os.path.dirname(input_path)

    libre_office = constants.LIBRE_OFFICE

    command = [
        libre_office,
        '--headless',
        '--convert-to',
        'odt',
        '--outdir',
        output_dir,
        input_path
    ]

    try:
        subprocess.run(command, check=True)
        print(f"Conversion successful: {odm_filepath} converted to ODT in {output_dir}")
    except subprocess.CalledProcessError as e:
        print(f"Error during conversion: {e}")
        print(f"Command executed: {' '.join(command)}")


# -------------------------------------------------------------------------------------------------

def convert_odt_to_docx(odt_file_path, docx_output_path):
    # Path to LibreOffice executable (adjust for your system)
    libre_office = constants.LIBRE_OFFICE

    # Command to convert using LibreOffice headless mode
    command = [
        libre_office,
        "--headless",
        "--convert-to",
        "docx",
        "--outdir",
        os.path.dirname(docx_output_path),
        odt_file_path
    ]

    try:
        subprocess.run(command, check=True)
        print(f"Conversion successful: {odt_file_path} converted to DOCX in {docx_output_path}")
    except subprocess.CalledProcessError as e:
        print(f"Error during conversion: {e}")
        print(f"Command executed: {' '.join(command)}")

# -------------------------------------------------------------------------------------------------

def open_libreoffice():
    # UNO component context for initializing the Python runtime
    localContext = uno.getComponentContext()

    # Create an instance of a service implementation
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext)


    context = resolver.resolve(
        "uno:socket,host=localhost,"
        "port=2083;urp;StarOffice.ComponentContext")

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
    if 'libreoffice' in interpreter_path.casefold():
        print('Correct Python Interpreter is being used.')
    else:
        print('LibreOffice Python Interpreter must be used.')
        exit()

    lcFolder = constants.FOLDER

    print(f"Folder: {lcFolder}")

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

            #convert_odm_to_odt(lcFolder + lcMaster)
            # convert_odt_to_docx(lcFolder + lcODT, lcFolder)

            desktop = open_libreoffice()





# -------------------------------------------------------------------------------------------------
