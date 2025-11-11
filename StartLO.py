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
    # -------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
