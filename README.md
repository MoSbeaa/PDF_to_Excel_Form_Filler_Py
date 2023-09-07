# Data Transfer Application README

Welcome to the Data Transfer Application README. This document provides a comprehensive overview of the Data Transfer Application developed using Python and PyQt5. The application facilitates data transfer from PDF files to Excel spreadsheets. This README guide covers the application's structure, features, and usage instructions.

## Table of Contents

- [Introduction](#introduction)
- [Getting Started](#getting-started)
- [Application Structure](#application-structure)
- [Usage](#usage)
- [Customization](#customization)
- [Converting to Executable](#converting-to-executable)
- [Conclusion](#conclusion)

## Introduction

The Data Transfer Application is a Python program built with the PyQt5 framework. Its purpose is to streamline data transfer from PDF files into Excel spreadsheets. Users can conveniently perform this operation using the application's intuitive graphical user interface.

## Getting Started

To begin using the Data Transfer Application, follow these steps:

1. Clone or download this repository to your local machine.
2. Install the necessary dependencies using pip:
```bash
pip install PyQt5 openpyxl PyPDF2
```

1. Run the application with the following command:
```bash
python main.py
```

## Application Structure

The Data Transfer Application is organized into the following key components:

- **main.py**: This file serves as the entry point of the application. It imports the necessary modules and initializes the QApplication, the foundation of the GUI.

- **app/__init__.py**: Responsible for initializing the QApplication, this module creates an instance of the OptionWindow class for selecting the program type. Depending on the chosen type, the corresponding window is displayed.

- **app/program_ui.py**: This module defines the user interface components of the application. It contains two essential classes:
  - **OptionWindow**: This class provides the initial interface for selecting the program type.
  - **DataTransferWindow**: This class presents the main data transfer interface, including buttons for selecting PDF files and Excel output, as well as the "Run," "Clear," and "Exit" buttons.

- **app/transfer_functions.py**: Data extraction and addition to Excel spreadsheets are managed by functions in this module. The primary function, `main_transfer`, orchestrates the entire data transfer process. Each section's data is extracted from the PDF and added to the appropriate Excel sheet.

The modular structure of the application enhances readability, maintainability, and extensibility, making it easier to modify or enhance specific components as needed.


## Usage

To effectively use the Data Transfer Application:

1. Launch the application and select the program type in the OptionWindow. Choose "Genome Canada Data transfer" and click "Select."

2. The DataTransferWindow will appear, offering the following actions:
   - Click the "Select PDF" button to choose a PDF file containing data from project reporters.
   - Use the "Select Genome Canada Excel Output" button to pick the Excel file for Genome Canada project details.
   - Use the "Select project Excel Output" button to select the Excel file for your records.
   - Click "Run" to start the data transfer process.

3. Successful data transfers will display "Process Completed" messages in the text area.

4. For errors, relevant messages will be shown.

5. Clear selections and output using the "Clear" button.

6. Exit the application via the "Exit" button.


## Customization

You can customize the application by making changes to the following components:

- **app/transfer_functions.py**: This module contains functions for extracting data from PDF files and adding data to Excel spreadsheets. Feel free to modify these functions to adapt the application to different PDF formats or Excel structures.

- **app/program_ui.py**: If you want to alter the GUI layout or design, you can adjust the UI classes in this module. You have the flexibility to add or remove widgets, change labels, or rearrange the layout according to your requirements.

Make these customizations to tailor the application to your specific needs and enhance its functionality to align with your use case.


## Converting to Executable

If you wish to convert the Data Transfer Application into an executable using the `psgcompiler` library, follow these steps:

1. Activate your virtual environment (assuming it's named `myenv`):

```bash
myenv\Scripts\activate
```
2. Launch psgcompiler by typing psgcompiler in the terminal.

3. In the psgcompiler window:

   - In the Home tab, select Python Script under the "Python Script (main.py)" section.
   - In the Additions tab, specify the paths for the app folder and the virtual environment myenv.
   - Click the "Convert" button in the Home tab.
   
After these steps, you should have an executable version of the Data Transfer Application that can be run on compatible systems without 
needing to install Python and dependencies.

Keep in mind that the specific instructions may vary based on your system configuration and the psgcompiler library version.