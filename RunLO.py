import os
import subprocess
import time
from pathlib import Path
from time import sleep

import keyboard
import psutil
import uno
from ooodev.loader.lo import Lo

from Constants import Constants


# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
class StartLO:
    # -------------------------------------------------------------------------------------------------
    def __init__(self):
        self.process = None
        self.constants = Constants()
        self.loader = None

    # -------------------------------------------------------------------------------------------------

    def start_libreoffice_headless(self):
        """
        Starts LibreOffice in headless mode, listening for connections on a specified port.
        """
        libreoffice_path = self.constants.LIBRE_OFFICE
        if not os.path.exists(libreoffice_path):
            # Example for Linux: libreoffice_path = "/usr/bin/libreoffice"
            # You might need to adjust this path based on your system.
            print(f"Error: LibreOffice executable not found at {libreoffice_path}")
            return None

        # Example:
        # command = [
        #     libreoffice_path,
        #     "--headless",
        #     f"--accept=socket,host=localhost,port={port};urp;",
        #     "--nofirststartwizard",
        #     "--nologo"
        # ]

        command = [
            libreoffice_path,
            "--writer",  # Or --writer, --draw, etc.
            self.constants.LIBRE_OFFICE_CONNECTION_INIT,
            "--nologo"
        ]

        # Use Popen to launch LibreOffice without blocking your Python script
        try:
            self.process = subprocess.Popen(command)
            # Give LibreOffice some time to start up
            time.sleep(5)

            return self.process
        except Exception as e:
            print(f"Error starting LibreOffice: {e}")
            self.constants.print_line_marker()
            return None

    # -------------------------------------------------------------------------------------------------
    def is_libreoffice_listener_running(self):

        lo_port = self.constants.LIBRE_OFFICE_PORT
        """Checks if a LibreOffice process is running in listener mode on a specific port."""
        for proc in psutil.process_iter(['name', 'cmdline']):
            try:
                if 'soffice' in proc.info['name'].lower():
                    # Check command line arguments for the listener flag and port
                    if proc.info['cmdline'] and any(f"port={lo_port}" in arg for arg in proc.info['cmdline']):
                        return True
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

        return False

    # -------------------------------------------------------------------------------------------------
    def stop_libreoffice_process(self):
        """
        Terminates the LibreOffice process.
        """
        lo_port = self.constants.LIBRE_OFFICE_PORT

        if self.is_libreoffice_listener_running():
            print(f"LibreOffice listener is running on port {lo_port}.")
        else:
            print(f"LibreOffice listener is not running on port {lo_port}.")
            exit(0)

        if self.process:
            self.process.terminate()
            print("LibreOffice process terminated.")
            self.constants.print_line_marker()

    # -------------------------------------------------------------------------------------------------
    def open_libreoffice(self):

        # UNO component context for initializing the Python runtime
        local_context = uno.getComponentContext()

        # Create an instance of a service implementation
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        try:
            remote_context = resolver.resolve(self.constants.LIBRE_OFFICE_CONNECTION_URI)
        except:
            raise Exception("Cannot establish a connection to LibreOffice.")

        return remote_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", remote_context
        )

    # -------------------------------------------------------------------------------------------------
    def convert_odm_to_odt_test(self, odm_filepath, master_file, odt_file):
        """
        Converts a LibreOffice master document (.odm) to an ODT file using the
        LibreOffice command-line interface.

        Args:
            odm_filepath (str): The full path to the master document (.odm).
        """
        input_file = uno.systemPathToFileUrl(os.path.abspath(odm_filepath + master_file))
        output_file = uno.systemPathToFileUrl(os.path.dirname(odm_filepath + odt_file))

        desktop = self.open_libreoffice()

        # Load a LibreOffice document, and automatically display it on the screen
        desktop.loadComponentFromURL(input_file, "_blank", 0, tuple([]))

        self.constants.print_line_marker()
        print("Wait for the master document to finish refreshing, then press any key to continue...")
        keyboard.read_key()

        property_value = uno.getClass('com.sun.star.beans.PropertyValue')

        save_opts = (
            property_value(Name="Overwrite", Value=True),
            property_value(Name="FilterName", Value="writer8"),
        )
        try:
            desktop.storeAsURL(output_file, save_opts)
            print("Document", input_file, " saved under ", output_file)
        finally:
            desktop.dispose()
            print("Document closed!")

        self.constants.print_line_marker()

    # -------------------------------------------------------------------------------------------------

    def convert_odt_to_docx_test(self, odt_file_path, docx_output_path):
        # Path to LibreOffice executable (adjust for your system)
        libre_office = self.constants.LIBRE_OFFICE

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

        self.constants.print_line_marker()

    # -------------------------------------------------------------------------------------------------
    def load_lo(self):
        try:
            if self.loader is None:
                self.loader = Lo.load_office(connector=Lo.ConnectPipe())
        except Exception as e:
            print(f"An error occurred while loading LibreOffice: {e}")
            if self.loader:
                Lo.close_office()
            exit(1)

    # -------------------------------------------------------------------------------------------------
    def close_lo(self):
        try:
            if self.loader:
                Lo.close_office()
        except Exception as e:
            print(f"An error occurred while closing LibreOffice: {e}")
            exit(1)

    # -------------------------------------------------------------------------------------------------

    def convert_odm_to_odt(self, odm_filepath, master_file, odt_file):

        self.constants.print_line_marker()

        input_file = Path(odm_filepath + master_file)
        output_file = Path(odm_filepath + odt_file)

        print(f"Master file: {input_file}")
        print(f"ODT File: {output_file}")

        try:
            # Explicitly open the document using the loader
            doc = Lo.open_doc(fnm=input_file, loader=self.loader)

            if doc is None:
                print(f"Failed to load document: {input_file}")
                return 1

            print(f"Document '{input_file}' opened successfully.")

            sleep(10)

            print("Press any key to continue...")
            # keyboard.read_key()  # Waits for any key to be pressed and returns its name
            Lo.save(doc)
            print(f"Document '{input_file}' saved successfully.")

            Lo.save_doc(doc, fnm=output_file)
            print(f"Document '{output_file}' saved successfully.")

        except Exception as e:
            print(f"An error occurred: {e}")
            if self.loader:
                Lo.close_office()
            exit(1)

        self.constants.print_line_marker()
        return 0

    # -------------------------------------------------------------------------------------------------

    def convert_odt_to_docx(self, odm_filepath, odt_file, docx_file):

        self.constants.print_line_marker()

        input_file = Path(odm_filepath + odt_file)
        output_file = Path(odm_filepath + docx_file)

        print(f"ODT file: {input_file}")
        print(f"DOCX File: {output_file}")

        try:
            # Explicitly open the document using the loader
            doc = Lo.open_doc(fnm=input_file, loader=self.loader)

            if doc is None:
                print(f"Failed to load document: {input_file}")
                return 1

            print(f"Document '{input_file}' opened successfully.")
            sleep(10)

            Lo.save(doc)
            print(f"Document '{input_file}' saved successfully.")

            Lo.save_doc(doc, fnm=output_file)
            print(f"Document '{output_file}' saved successfully.")

        except Exception as e:
            print(f"An error occurred: {e}")
            if self.loader:
                Lo.close_office()
            return 1

        self.constants.print_line_marker()
        return 0

    # -------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
