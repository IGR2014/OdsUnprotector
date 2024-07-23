# ODS Unprotector

This script removes protection from tables and cells in OpenDocument Spreadsheet (.ods) files.
It can process a single file or all .ods files within a specified directory.

## Requirements

- Python 3.x
- `odfpy` library

## Installation

1. Ensure you have Python 3.x installed.
2. Install the `odfpy` library using pip:

    ```bash
    pip install odfpy
    ```

## Usage

You can use the script to unprotect a single .ods file or all .ods files within a directory.

- Single File:

    ```bash
    python unprotect_ods.py /path/to/your/file.ods
    ```

- All Files in a Directory:

    ```bash
    python unprotect_ods.py /path/to/your/directory
    ```

The script will automatically detect whether the given path is a file or a directory and process accordingly.

## Example

- To unprotect a single file:

    ```bash
    python unprotect_ods.py sample.ods
    ```

- To unprotect all .ods files in a folder:

    ```bash
    python unprotect_ods.py /path/to/your/folder
    ```

## Notes

The script will skip non-.ods files when processing a directory.
If an error occurs while processing a file, an error message will be printed, and the script will continue with the next file (if any).

## License

This project is licensed under the MIT License - see the LICENSE file for details.
