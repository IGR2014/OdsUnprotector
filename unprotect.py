# Import the argparse module for command-line argument parsing
import argparse
# Import the os module for interacting with the operating system
import os
# Import the load function from the odf.opendocument module to load ODS files
from odf.opendocument import load
# Import Table and TableCell from the odf.table module to work with tables and cells
from odf.table import Table, TableCell


# Unprotect file
def unprotect_ods(file_path):
    # Try block
    try:
        # Load the ODS file
        doc = load(file_path)
        # Iterate over all tables in the document
        for table in doc.getElementsByType(Table):
            # Check if the 'protected' attribute exists
            if table.getAttribute("protected"):
                # Set the 'protected' attribute to 'false'
                table.setAttribute("protected", "false")
            # Iterate over all cells in the table
            for cell in table.getElementsByType(TableCell):
                # Check if the 'protected' attribute exists
                if cell.getAttribute("protected"):
                    # Set the 'protected' attribute to 'false'
                    cell.setAttribute("protected", "false")
        # Save the modified ODS file
        doc.save(file_path)
        print(f"Protection set to 'false' for {file_path}")
    # Exceptions processing
    except Exception as e:
        # Print an error message if there is an exception
        print(f"Error processing {file_path}: {e}")


# Process path
def process_path(path):
    # Is file ?
    if os.path.isfile(path):
        # If the path is a file, unprotect it
        if path.lower().endswith('.ods'):
            # Check if the file has a .ods extension
            # Call the function to unprotect the ODS file
            unprotect_ods(path)
        else:
            # Print a message if the file is not an ODS file
            print(f"Skipping {path}: Not an ODS file")
    # Is directory ?
    elif os.path.isdir(path):
        # If the path is a directory, process all .ods files in the directory
        for root, _, files in os.walk(path):
            # Iterate over all files in the directory
            for file in files:
                if file.lower().endswith('.ods'):
                    # Check if the file has a .ods extension
                    # Call the function to unprotect the ODS file
                    unprotect_ods(os.path.join(root, file))
    else:
        # Print a message if the path is invalid
        print(f"Invalid path: {path}")


# Main
if __name__ == "__main__":
    # Create an ArgumentParser object
    parser = argparse.ArgumentParser(description="Remove protection from .ods tables")
    # Add an argument to the parser for the path to the .ods file or directory
    parser.add_argument("path", help="The path to the .ods file or directory to unprotect")
    # Parse the command-line arguments
    args = parser.parse_args()
    # Process the provided path
    process_path(args.path)
