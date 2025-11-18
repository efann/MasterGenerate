from pathlib import Path

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
    def load_lo(self):
        try:
            if self.loader is None:
                self.loader = Lo.load_office(connector=Lo.ConnectSocket())
        except Exception as e:
            print(f"An error occurred while loading LibreOffice: {e}")
            self.close_lo()

    # -------------------------------------------------------------------------------------------------
    def close_lo(self):
        print("Closing LibreOffice.")

        try:
            if self.loader:

                Lo.close_office()
                # Lo.kill_office()

                print("Closed LibreOffice.")
        except Exception as e:
            print(f"An error occurred while closing LibreOffice: {e}")
            exit(1)

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
            self.close_lo()

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
            self.close_lo()

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
