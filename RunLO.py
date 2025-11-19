import os
import signal
from pathlib import Path

import psutil
import uno
from com.sun.star.util import XLinkUpdate
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

        # If you don't leave this print statement, then the import uno will be removed by
        # PyCharm code cleanup. And
        # from com.sun.star.util import XLinkUpdate
        # will throw
        #   raise ImportError("Are you sure that uno has been imported?")
        #   ImportError: Are you sure that uno has been imported?
        # Sigh. . . .
        print(uno.getComponentContext())

    # -------------------------------------------------------------------------------------------------
    def is_libreoffice_listener_running(self):

        """Checks if a LibreOffice process is running in listener mode on a specific port."""
        for proc in psutil.process_iter(['name', 'pid', 'cmdline']):
            try:
                if 'soffice' in proc.info['name'].lower():
                    print(proc.info)
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

        return False

    # -------------------------------------------------------------------------------------------------
    # Lo.close_office() hangs
    # Lo.kill_office states that a termination signal has been sent to the process "soffice.bin" with PID ###,
    # but soffice.bin & soffice.exe remain in the Task Manager
    # So I just close them programmatically with this function.
    def terminate_libreoffice(self):

        for proc in psutil.process_iter(['name', 'pid']):
            try:
                ln_pid = proc.info["pid"]
                lc_name = proc.info["name"]
                # Turns out that soffice.bin controls soffice.exe. When soffice.bin terminates, it terminates soffice.exe
                if 'soffice.bin' in lc_name.lower():
                    # Check command line arguments for the listener flag and port
                    try:
                        print(f"Attempting to terminate process ({lc_name}: {ln_pid})...")
                        # Attempt a graceful termination first (SIGTERM)
                        os.kill(ln_pid, signal.SIGTERM)
                        proc.wait(timeout=5)  # Wait for the process to terminate gracefully
                        print(f"Process {ln_pid} terminated gracefully.")
                    except psutil.NoSuchProcess:
                        print(f"Process {ln_pid} already terminated.")
                    except psutil.TimeoutExpired:
                        print(f"Process {ln_pid} did not terminate gracefully, forcing kill (SIGKILL).")
                        # If graceful termination fails, force kill (SIGKILL)
                        os.kill(ln_pid, signal.SIGKILL)
                        print(f"Process {ln_pid} forcefully terminated.")
                    except Exception as e:
                        print(f"Error terminating process {e}")

            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

    # -------------------------------------------------------------------------------------------------
    def load_lo(self):
        try:
            if self.loader is None:
                self.loader = Lo.load_office(connector=Lo.ConnectSocket(headless=True))
        except Exception as e:
            print(f"An error occurred while loading LibreOffice: {e}")
            self.terminate_libreoffice()

    # -------------------------------------------------------------------------------------------------

    def convert_odm_to_odt(self, master_file, odt_file):

        self.constants.print_line_marker()

        input_file = Path(master_file)
        output_file = Path(odt_file)

        print(f"Master file: {input_file}")
        print(f"ODT File: {output_file}")

        try:
            # Explicitly open the document using the loader
            doc = Lo.open_doc(fnm=input_file, loader=self.loader)

            if doc is None:
                print(f"Failed to load document: {input_file}")
                return 1

            print(f"Document '{input_file}' opened successfully.")

            self.update_links(doc)

            Lo.save(doc)
            print(f"Document '{input_file}' saved successfully.")

            Lo.save_doc(doc, fnm=output_file)
            print(f"Document '{output_file}' saved successfully.")

        except Exception as e:
            print(f"An error occurred: {e}")

        self.constants.print_line_marker()
        return 0

    # -------------------------------------------------------------------------------------------------

    def convert_odt_to_docx(self, odt_file, docx_file):

        self.constants.print_line_marker()

        input_file = Path(odt_file)
        output_file = Path(docx_file)

        print(f"ODT file: {input_file}")
        print(f"DOCX File: {output_file}")

        try:
            # Explicitly open the document using the loader
            doc = Lo.open_doc(fnm=input_file, loader=self.loader)
            if doc is None:
                print(f"Failed to load document: {input_file}")
                return 1

            print(f"Document '{input_file}' opened successfully.")

            self.update_links(doc)

            print(f"Document '{input_file}' updated successfully.")
            Lo.save(doc)
            print(f"Document '{input_file}' saved successfully.")

            Lo.save_doc(doc, fnm=output_file)
            print(f"Document '{output_file}' saved successfully.")

        except Exception as e:
            print(f"An error occurred: {e}")

        self.constants.print_line_marker()
        return 0

    # -------------------------------------------------------------------------------------------------

    def update_links(self, doc):
        # Query the document for the XLinkUpdate interface
        link_update = Lo.qi(XLinkUpdate, doc)
        if link_update:
            try:
                # Call the updateLinks method
                print("Updating links.")
                link_update.updateLinks()
                print("Document links updated successfully.")
            except Exception as e:
                print(f"Error updating links: {e}")
        else:
            print("Document does not support XLinkUpdate interface or links are not present.")

    # -------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
