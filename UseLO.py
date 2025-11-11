import subprocess
import os
import time

from Constants import Constants


# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
class StartLO:
    # -------------------------------------------------------------------------------------------------
    def __init__(self):
        self.process = None
        self.constants = Constants()

    # -------------------------------------------------------------------------------------------------

    def start_libreoffice_headless(self, port):
        """
        Starts LibreOffice in headless mode, listening for connections on a specified port.
        """
        libreoffice_path = self.constants.LIBRE_OFFICE
        if not os.path.exists(libreoffice_path):
            # Example for Linux: libreoffice_path = "/usr/bin/libreoffice"
            # You might need to adjust this path based on your system.
            print(f"Error: LibreOffice executable not found at {libreoffice_path}")
            return None

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
            f'--accept="socket,host=localhost,port={port};urp;StarOffice.ServiceManager"'
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

        # try:
        #     self.process = subprocess.Popen(command)
        #     print(f"LibreOffice process started in headless mode on port {port}.")
        #     return self.process
        # except Exception as e:
        #     print(f"Error starting LibreOffice: {e}")
        #     return None

    # -------------------------------------------------------------------------------------------------
    def stop_libreoffice_process(self):
        """
        Terminates the LibreOffice process.
        """
        if self.process:
            self.process.terminate()
            print("LibreOffice process terminated.")
            self.constants.print_line_marker()

    # -------------------------------------------------------------------------------------------------
    def convert_odm_to_odt(self, odm_filepath):
        """
        Converts a LibreOffice master document (.odm) to an ODT file using the
        LibreOffice command-line interface.

        Args:
            odm_filepath (str): The full path to the master document (.odm).
        """
        input_path = os.path.abspath(odm_filepath)
        output_dir = os.path.dirname(input_path)

        libre_office = self.constants.LIBRE_OFFICE

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
            self.constants.print_line_marker()

    # -------------------------------------------------------------------------------------------------

    def convert_odt_to_docx(self, odt_file_path, docx_output_path):
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

    # -------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
