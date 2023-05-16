# Excel Validator

The Excel Validator is a command-line tool that validates the content of an Excel file against a set of rules. It can detect errors related to cell values, lengths, and date ranges, and output them in a text file and an Excel file.
## Requirements
Python 3.x
openpyxl module (can be installed via pip)

## Installing Python
1. Download the latest version of Python from the official website: https://www.python.org/downloads/windows/
2. Double-click the downloaded installer file (.exe) to launch the Python installer.
3. Follow the instructions on the installer to complete the installation process.
4. Verify that Python has been installed correctly by opening a command prompt and typing python --version. The installed Python version should be displayed.

## Installing Libraries
### Openpyxl 
1. Open a command prompt and type pip install openpyxl to install the openpyxl library.
2. Verify that openpyxl  has been installed correctly by opening a Python shell and typing import openpyxl. If no error message is displayed, openpyxl has been installed successfully.

## Installation
Clone this repository or download the source code.
Install the openpyxl module by running the following installation library openpyxl

## Installation via requrements.txt

1. pip install --user virtualenv (installing virtualenv if not already installed)
2. python -m venv env (creating a virtual environment named "env" in the root folder)
3. ./env/Scripts/activate (activating the virtual environment)
4. pip install -r requirements.txt (installing the requirements specified in the "requirements.txt" file)
5. python main.py (running the script)


## Usage
1. Navigate to the project directory using your terminal.
2. Run the main.py script.
3. When you start the app, you will see a window with a "Choose file" button.
4. Click the "Choose file" button to select an Excel file for validation. The app only accepts .xls and .xlsx file types. If the selected file is valid, the "Validate File" button will be enabled.
5. Click the "Validate File" button to validate the file. The results will be displayed in a scrollable listbox.
6. A text file containing the validation errors will be saved in the Downloads folder with a timestamp in the filename.
7. An Excel file containing the validation errors will be saved in the Downloads folder with a timestamp in the filename.
8. To exit the app, click the "x" button in the top right corner of the window or press "Ctrl+C" in the terminal.
## Rules
The Excel Validator validates the content of an Excel file against the following rules:
+ The length of the cell value in the first column must be 3 characters.
+ The value of the "Default" column must be "Y" or "N".
+ The date in any cell must be in the range between January 1, 2000, and the current date.
## Output
The Excel Validator outputs the result of the validation in the console, a text file, and an Excel file. The text file and Excel file contain the same validation errors in a different format.
The output files are saved in the Downloads folder with a timestamp in the filename.

## Limitations
+ The Excel Validator only works with .xlsx files.
+ The Excel Validator assumes that the Excel file has a sheet with data in it.
+ The Excel Validator assumes that the sheet has a column named "Default".
